#!/usr/bin/env python3
"""
Publica al CMMS los TIEMPOS PERDIDOS que la faena declaró al NOC de Arauco.

POR QUÉ EXISTE. El tablero de disponibilidad del CMMS (cmms.millalemu.com/disponibilidad)
solo sabía del tiempo perdido por el pre-uso del operador, y un día sin pre-uso no tenía
causa. El NOC sí la tiene —la gente declara donde le pagan— y este pipeline ya lo baja
todas las noches. Este script es el puente, igual que publicar_vma.py.

TRES TRAMPAS DEL DATO (ver noc_tiempo_perdido en el CMMS):
  · `codigo_equipo` del TP (S246, MG5…) es la FAENA, no una máquina. El proceso sale de la
    causa («Falla equipo de madereo»); la clasificación la hace el CMMS.
  · El día trabajado es `hora_inicio` del folio en el reporte PG. La `Fecha` del TP es la de
    EMISIÓN (pasada la medianoche): usarla corre la mitad de los registros un día.
  · Arauco corrige folios: el CMMS reemplaza todo lo de cada numero_noc que llega. Por eso
    los lotes NUNCA parten un folio en dos (el segundo lote borraría al primero).

USO
  python3 publicar_tp.py                          # año en curso, mes a mes (cada noche)
  python3 publicar_tp.py --desde 2026-10-01 --hasta 2026-10-31   # recarga de un período
  python3 publicar_tp.py --dry-run                # arma y resume, no publica

NO CRÍTICO: si falla imprime y sale con 0; el pipeline sigue.
"""

import argparse
import json
import os
import sys
import urllib.request
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).parent
sys.path.insert(0, str(BASE))

from compute_kpis import TEAM_MAP  # código NOC de faena → M1.4, M5… (fuente única del pipeline)
from publicar_vma import FAENA_ID  # M1.4 → faena-m1-4 (id del nodo en el CMMS)

LOTE = 400  # registros por POST; se corta solo entre folios


def _txt(v):
    return '' if v is None else str(v).strip()


def armar_filas(tp, pg):
    """tp, pg: registros crudos de la API del NOC. → (filas para el CMMS, resumen)."""
    dia_por_folio = {}
    for r in pg:
        folio = r.get('numero_noc') or r.get('folio')
        hi = _txt(r.get('hora_inicio'))[:10]
        if folio and len(hi) == 10:
            # Un folio = un turno; si viniera repetido, gana el primer inicio del día.
            dia_por_folio.setdefault(int(folio), hi)

    filas, sin_dia, sin_faena = [], 0, Counter()
    for r in tp:
        folio = r.get('Numero_noc') or r.get('numero_noc')
        codigo = _txt(r.get('codigo_tiempo_perdido'))
        if not folio or not codigo:
            continue
        folio = int(folio)
        faena_noc = _txt(r.get('codigo_equipo'))
        team = TEAM_MAP.get(faena_noc)
        faena_id = FAENA_ID.get(team) if team else None
        if not faena_id:
            sin_faena[faena_noc] += 1
        dia = dia_por_folio.get(folio)
        if not dia:
            sin_dia += 1
        try:
            minutos = float(str(r.get('tiempo') or 0).replace(',', '.'))
        except ValueError:
            minutos = 0.0
        filas.append({
            'numero_noc': folio, 'faena_noc': faena_noc, 'faena_id': faena_id or '',
            'dia': dia or '', 'codigo': codigo, 'descripcion': _txt(r.get('descripcion')),
            'minutos': max(0.0, min(minutos, 1440.0)), 'observacion': _txt(r.get('observacion'))[:500],
            'estado_noc': _txt(r.get('estado_noc')),
        })
    resumen = {
        'registros': len(filas), 'folios': len({f['numero_noc'] for f in filas}),
        'horas': round(sum(f['minutos'] for f in filas) / 60, 1),
        'sin_dia': sin_dia, 'sin_faena': dict(sin_faena),
    }
    return filas, resumen


def lotes_por_folio(filas, tam=LOTE):
    """Parte en lotes de ~tam sin separar nunca los registros de un mismo folio."""
    por_folio = defaultdict(list)
    for f in filas:
        por_folio[f['numero_noc']].append(f)
    lote = []
    for regs in por_folio.values():
        if lote and len(lote) + len(regs) > tam:
            yield lote
            lote = []
        lote.extend(regs)
    if lote:
        yield lote


def _meses(desde, hasta):
    """[(ini, fin)] por mes calendario dentro de [desde, hasta] (fechas 'YYYY-MM-DD')."""
    from datetime import date, timedelta
    d = date.fromisoformat(desde)
    h = date.fromisoformat(hasta)
    tramos = []
    while d <= h:
        sig = (d.replace(year=d.year + 1, month=1, day=1) if d.month == 12
               else d.replace(month=d.month + 1, day=1))
        tramos.append((d.isoformat(), min(h, sig - timedelta(days=1)).isoformat()))
        d = sig
    return tramos


def descargar(desde, hasta):
    """Baja TP y PG MES POR MES (pedirle un año entero de una vez al NOC es buscarse un
    timeout) y los junta ANTES de cruzar: un TP emitido el día 1 puede ser de un folio
    trabajado el último día del mes anterior, y su día sale del PG de ese otro mes."""
    import requests
    import descargar_noc_api as noc
    s = requests.Session()
    s.verify = False
    s.headers.update({'User-Agent': 'Mozilla/5.0'})
    tok = noc.obtener_token_arcgis(s)
    try:
        noc.establecer_sesion(s, tok)
    except Exception:
        pass
    tp, pg = [], []
    for ini, fin in _meses(desde, hasta):
        t = noc.descargar_reporte(s, tok, 'TP', ini, fin) or []
        p = noc.descargar_reporte(s, tok, 'PG', ini, fin) or []
        print(f"   · {ini[:7]}: {len(t)} TP · {len(p)} folios PG")
        tp.extend(t)
        pg.extend(p)
    return tp, pg


def publicar(filas):
    url = os.environ.get('SUPABASE_URL')
    key = os.environ.get('SUPABASE_KEY')
    if not url or not key:
        print("   ⚠️  Sin SUPABASE_URL/SUPABASE_KEY: no se publican los tiempos perdidos.")
        return
    tot_b = tot_i = 0
    for lote in lotes_por_folio(filas):
        req = urllib.request.Request(
            url.rstrip('/') + '/rest/v1/rpc/noc_tp_publicar',
            data=json.dumps({'filas': lote}).encode('utf-8'), method='POST',
            headers={'apikey': key, 'Authorization': f'Bearer {key}', 'Content-Type': 'application/json'},
        )
        with urllib.request.urlopen(req, timeout=60) as r:
            resp = json.loads(r.read().decode('utf-8') or '{}')
        if not resp.get('ok'):
            raise RuntimeError(f"el CMMS rechazó el lote: {resp}")
        tot_b += resp.get('borradas', 0)
        tot_i += resp.get('insertadas', 0)
    print(f"   ✅ Tiempos perdidos publicados: {tot_i} registros (reemplazó {tot_b}).")


def main(argv=None):
    ap = argparse.ArgumentParser()
    hoy = datetime.now()
    # Por defecto, TODO EL AÑO en curso (y como mínimo desde el mes anterior, para que en
    # enero no quede diciembre a medias). Cuesta ~2 consultas cortas por mes y, como el CMMS
    # reemplaza por folio, recargar no duplica: así entra cualquier corrección tardía de
    # Arauco y un mes que se perdió por una caída del NOC se completa solo la noche siguiente.
    primero = hoy.replace(day=1)
    mes_ant = (primero.replace(year=primero.year - 1, month=12) if primero.month == 1
               else primero.replace(month=primero.month - 1))
    desde = min(mes_ant, hoy.replace(month=1, day=1))
    ap.add_argument('--desde', default=desde.strftime('%Y-%m-%d'))
    ap.add_argument('--hasta', default=hoy.strftime('%Y-%m-%d'))
    ap.add_argument('--dry-run', action='store_true')
    a = ap.parse_args(argv)
    print(f"   ⏱️  Tiempos perdidos NOC → CMMS ({a.desde} → {a.hasta})…")
    try:
        tp, pg = descargar(a.desde, a.hasta)
        if not tp:
            print("   ⚠️  El NOC no devolvió tiempos perdidos; nada que publicar.")
            return 0
        filas, res = armar_filas(tp, pg)
        print(f"   {res['registros']} registros · {res['folios']} folios · {res['horas']} h"
              f" · sin día trabajado {res['sin_dia']} · faena desconocida {res['sin_faena'] or 0}")
        if not a.dry_run:
            publicar(filas)
    except Exception as e:
        print(f"   ⚠️  No se publicaron los tiempos perdidos ({e}). Se reintenta en la próxima corrida.")
    return 0


if __name__ == '__main__':
    sys.exit(main())
