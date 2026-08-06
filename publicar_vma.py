#!/usr/bin/env python3
"""
Publica al CMMS el VMA (volumen medio del árbol) por FAENA y DÍA.

POR QUÉ EXISTE. Desde 2026-08-05 el jefe de faena ya no estima m³ al ojo en /t/avance:
CUENTA árboles y viajes, y el volumen lo deriva el sistema con  m³ = árboles × VMA  (la
misma fórmula con la que Arauco llena su tablero). Ese VMA no lo sabe el CMMS: sale del
reporte Productividad Genérico del NOC, que vive acá en el pipeline. Este script es el
puente.

CÓMO SE CALCULA. Exactamente igual que en compute_kpis.py:  VMA = Σ Vol ÷ Σ Árboles  del
día y la faena. Se copia la fórmula a propósito, para que el número que ve el jefe en el
teléfono sea EL MISMO que el de la pestaña KPIs y el del informe. Si algún día cambia el
criterio, tiene que cambiar en los dos lados.

⚠ OJO CON EL DATO: el campo del NOC se llama `arboles_madereados` pero son los árboles que
pasó el PROCESADOR, no los que arrastró el skidder. Da igual para el VMA —numerador y
denominador salen los dos del procesador, así que el ratio es sólido y no lo distorsiona el
movimiento del colchón— pero no hay que leerlo como producción del madereo.

NO CRÍTICO. Si falla (sin credenciales, sin red, sin CSV) imprime y sale con 0: el pipeline
sigue. Sin VMA publicado la app de terreno igual deja declarar —el conteo es el dato duro— y
el volumen se completa solo cuando este script corra, porque `vma_faena_publicar` convierte
las filas que quedaron pendientes.
"""

import json
import os
import sys
import urllib.request
from pathlib import Path

import pandas as pd

BASE = Path(__file__).parent
sys.path.insert(0, str(BASE))

from compute_kpis import TEAM_MAP, _num, _fetch_pg_api  # misma fuente de verdad que los KPIs
from normalizar_produccion import normalizar

# ⚠ El CSV es el RESPALDO, no la fuente. En el runner de GitHub NO existe: compute_kpis
# baja el reporte PG por la API del NOC y solo cae al CSV si la API falla. La primera
# versión de este script leía únicamente el CSV, así que el paso corría, avisaba
# "No está ProductividadGenerico.csv" y seguía sin publicar nada — dos corridas en
# verde con la tabla vacía. Misma fuente que los KPIs o los números se separan.
CSV_PROD = BASE / "ProductividadGenerico.csv"

# Código de faena → id del nodo en el CMMS. Espejo de FAENA_ID de generar_informe_faena.py:
# si se agrega una faena, tocar los dos.
FAENA_ID = {'M1.1': 'faena-m1-1', 'M1.2': 'faena-m1-2', 'M1.3': 'faena-m1-3',
            'M1.4': 'faena-m1-4', 'M5': 'faena-m5', 'M7': 'faena-m7',
            'M9': 'faena-m9', 'M11': 'faena-m11'}

# Techo del CHECK de la tabla. Un VMA sobre esto es un dato roto, no un árbol grande: se
# descarta en vez de mandarlo y que el insert reviente entero.
VMA_MAX = 20.0


def _cargar_prod():
    """El reporte PG: primero la API del NOC (lo que usa compute_kpis), CSV de respaldo."""
    prod = _fetch_pg_api()
    if prod is not None and len(prod):
        print(f"   📡 PG por API del NOC: {len(prod)} folios.")
        return prod
    if CSV_PROD.exists():
        print(f"   📄 API sin datos; uso {CSV_PROD.name}.")
        return normalizar(pd.read_csv(CSV_PROD, sep=';', encoding='utf-8-sig'))
    print("   ⚠️  Ni API ni CSV: no hay VMA que publicar.")
    return None


def filas_vma():
    """[{faena_id, fecha, m3, arboles, vma}] por faena y día. Misma fuente que los KPIs."""
    prod = _cargar_prod()
    if prod is None:
        return []
    for c in ['Volumen SSC PU', 'Volumen SSC AS']:
        if c not in prod.columns:
            prod[c] = 0
        prod[c] = _num(prod[c])
    prod['Vol'] = prod['Volumen SSC PU'].fillna(0) + prod['Volumen SSC AS'].fillna(0)
    if 'Árboles Madereados' not in prod.columns:
        print("   ⚠️  La fuente no trae árboles (Base2NOC no los tiene); sin VMA que publicar.")
        return []
    prod['Arb'] = _num(prod['Árboles Madereados']).fillna(0)
    prod['Team'] = prod['Equipo'].map(TEAM_MAP)
    prod['Fecha_dt'] = pd.to_datetime(prod['Fecha NOC'], dayfirst=True, errors='coerce')
    prod = prod[prod['Team'].notna() & prod['Fecha_dt'].notna()]

    filas = []
    for (team, fecha), d in prod.groupby(['Team', prod['Fecha_dt'].dt.date]):
        fid = FAENA_ID.get(team)
        if not fid:
            continue
        vol = float(d['Vol'].sum())
        arb = float(d['Arb'].sum())
        # Sin árboles no hay denominador: ese día simplemente no tiene VMA. Publicar un 0
        # sería peor que no publicar nada (la app multiplicaría por cero).
        if arb <= 0 or vol <= 0:
            continue
        vma = vol / arb
        if not (0 < vma <= VMA_MAX):
            print(f"   ⚠️  VMA fuera de rango en {team} {fecha}: {vma:.3f} — se omite.")
            continue
        filas.append({'faena_id': fid, 'fecha': str(fecha),
                      'm3': round(vol, 2), 'arboles': int(round(arb)), 'vma': round(vma, 4)})
    return filas


def publicar(filas):
    """POST a la RPC vma_faena_publicar. Requiere la clave SECRETA (service_role)."""
    url = os.environ.get('SUPABASE_URL')
    key = os.environ.get('SUPABASE_KEY')
    if not url or not key:
        print("   ⚠️  Sin SUPABASE_URL/SUPABASE_KEY: no se publica el VMA.")
        return False
    req = urllib.request.Request(
        url.rstrip('/') + '/rest/v1/rpc/vma_faena_publicar',
        data=json.dumps({'filas': filas}).encode('utf-8'),
        method='POST',
        headers={'apikey': key, 'Authorization': f'Bearer {key}',
                 'Content-Type': 'application/json'},
    )
    with urllib.request.urlopen(req, timeout=60) as r:
        cuerpo = r.read().decode('utf-8').strip()
    print(f"   ✅ VMA publicado: {len(filas)} días-faena (respuesta: {cuerpo}).")
    return True


def main():
    print("   🌲 Calculando VMA por faena y día para el CMMS…")
    filas = filas_vma()
    if not filas:
        print("   ⚠️  Nada que publicar.")
        return 0
    ultimos = sorted({f['fecha'] for f in filas})[-1:]
    print(f"   {len(filas)} días-faena, último {ultimos[0] if ultimos else '—'}.")
    try:
        publicar(filas)
    except Exception as e:
        # NO crítico: el conteo del jefe se guarda igual y el volumen se completa en la
        # próxima corrida. Que el pipeline caiga por esto sería peor.
        print(f"   ⚠️  No se pudo publicar el VMA ({e}). Se reintenta en la próxima corrida.")
    return 0


if __name__ == '__main__':
    sys.exit(main())
