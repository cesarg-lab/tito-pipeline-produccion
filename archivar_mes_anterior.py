#!/usr/bin/env python3
"""
archivar_mes_anterior.py
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Guarda el cierre del MES ANTERIOR en historico_cierres_mensuales.csv.

Pensado para la NUBE (GitHub Actions, sin disco persistente): como cada
corrida arranca de cero y baja solo el mes en curso, el mes que recién
cerró se perdería. Este script lo recupera bajándolo directo de Arauco.

Idempotente: si el mes anterior YA está en el histórico, no hace nada
(termina con código 0). Así puede correr todos los días sin duplicar.

Uso:  python3 archivar_mes_anterior.py
Requiere: ARAUCO_USER / ARAUCO_PASS (o .noc_config.json), pandas, requests.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
import os, sys, csv, calendar
from pathlib import Path
from datetime import datetime
import importlib.util
import pandas as pd
import urllib3
urllib3.disable_warnings()

BASE_DIR = Path(__file__).parent
HIST_CSV = BASE_DIR / "historico_cierres_mensuales.csv"
sys.path.insert(0, str(BASE_DIR))
from normalizar_produccion import normalizar  # noqa: E402

TEAM_MAP = {
    'S123':'Millalemu 1.1','S58':'Millalemu 1.2','S223':'Millalemu 1.3',
    'S246':'Millalemu 1.4','MG5':'Millalemu 5','TEA02':'Millalemu 7',
    'TEA08':'Millalemu 9','T125':'Millalemu 11','TEA30':'Millalemu 1.3',
}
ORDEN = ['Millalemu 1.1','Millalemu 1.2','Millalemu 1.3','Millalemu 1.4',
         'Millalemu 5','Millalemu 7','Millalemu 9','Millalemu 11']
METAS_DEFAULT = {'Millalemu 1.1':7000,'Millalemu 1.2':7000,'Millalemu 1.3':7000,
                 'Millalemu 1.4':7000,'Millalemu 5':7000,'Millalemu 7':7000,
                 'Millalemu 9':7000,'Millalemu 11':6000}
MESES = ["","Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio",
         "Agosto","Septiembre","Octubre","Noviembre","Diciembre"]

def log(m): print(f"[archivar_mes_anterior] {m}")

def metas_desde_excel():
    """Lee metas de la hoja CONFIGURACIÓN si el Excel existe; si no, defaults."""
    metas = dict(METAS_DEFAULT)
    excel = BASE_DIR / "Dashboard_CosechaForestal.xlsx"
    if excel.exists():
        try:
            from openpyxl import load_workbook
            wb = load_workbook(str(excel), data_only=True)
            if "CONFIGURACIÓN" in wb.sheetnames:
                ws = wb["CONFIGURACIÓN"]
                for i, t in enumerate(ORDEN):
                    v = ws.cell(24 + i, 5).value
                    if v:
                        metas[t] = float(v)
            wb.close()
        except Exception as e:
            log(f"aviso: no se pudo leer metas del Excel ({e}); uso defaults")
    return metas

def mes_ya_archivado(mes, anio):
    if not HIST_CSV.exists():
        return False
    with open(HIST_CSV, encoding='utf-8-sig', newline='') as f:
        for row in csv.reader(f, delimiter=';'):
            if len(row) >= 2 and row[0].strip() == str(mes) and row[1].strip() == str(anio):
                return True
    return False

def descargar_mes(mes, anio):
    """Baja PG del mes desde Arauco y devuelve el DataFrame normalizado."""
    spec = importlib.util.spec_from_file_location("dnoc", str(BASE_DIR / "descargar_noc_api.py"))
    dn = importlib.util.module_from_spec(spec); spec.loader.exec_module(dn)
    import requests
    s = requests.Session(); s.verify = False
    s.headers.update({"User-Agent": "Mozilla/5.0"})
    token = dn.obtener_token_arcgis(s)
    dn.establecer_sesion(s, token)
    ult_dia = calendar.monthrange(anio, mes)[1]
    fi, ff = f"{anio}-{mes:02d}-01", f"{anio}-{mes:02d}-{ult_dia:02d}"
    log(f"descargando {MESES[mes]} {anio} ({fi} → {ff})...")
    pg = dn.descargar_reporte(s, token, "PG", fi, ff)
    if not pg:
        return None
    dn.guardar_csv(pg, "PG", Path("/tmp"))
    return pd.read_csv("/tmp/ProductividadGenerico.csv", sep=';', encoding='utf-8-sig')

def cierre(prod, mes, anio, metas):
    prod = normalizar(prod)
    for c in ['Volumen SSC PU', 'Volumen SSC AS']:
        prod[c] = pd.to_numeric(prod[c].astype(str).str.replace(',', '.'), errors='coerce')
    prod['Vol'] = prod['Volumen SSC PU'].fillna(0) + prod['Volumen SSC AS'].fillna(0)
    _te = pd.to_numeric(prod['Tiempo Efectivo'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
    prod['HrsEf'] = _te / (3600 if _te.max() > 1000 else 60)
    prod['Team'] = prod['Equipo'].map(TEAM_MAP)
    prod['Fecha_dt'] = pd.to_datetime(prod['Fecha NOC'], dayfirst=True, errors='coerce')
    prod['Dia'] = prod['Fecha_dt'].dt.day
    prod = prod[prod['Fecha_dt'].dt.month == mes]
    fuente = f"Auto-archivado nube {datetime.now().strftime('%Y-%m-%d')}"
    filas = []
    for t in ORDEN:
        td = prod[prod['Team'] == t]
        vol = float(td['Vol'].sum()); hrs = float(td['HrsEf'].sum())
        dias = int(td['Dia'].nunique()); meta = metas[t]
        cumpl = vol / meta * 100 if meta else 0
        prom = vol / dias if dias else 0
        rend = vol / hrs if hrs else 0
        filas.append(f"{mes};{anio};{MESES[mes]};{t};{meta:.0f};{vol:.1f};{cumpl:.1f};{dias};{dias};{prom:.1f};{rend:.1f};{hrs:.1f};{fuente}")
    return filas

def main():
    hoy = datetime.now()
    primer_dia = hoy.replace(day=1)
    prev = primer_dia.fromordinal(primer_dia.toordinal() - 1)  # último día mes anterior
    mes, anio = prev.month, prev.year

    if mes_ya_archivado(mes, anio):
        log(f"{MESES[mes]} {anio} ya está en el histórico — nada que hacer.")
        return 0

    log(f"{MESES[mes]} {anio} NO está en el histórico — recuperando de Arauco...")
    prod = descargar_mes(mes, anio)
    if prod is None or prod.empty:
        log(f"sin datos de {MESES[mes]} {anio} en Arauco — no se archiva (no es error).")
        return 0

    filas = cierre(prod, mes, anio, metas_desde_excel())
    total = sum(float(r.split(';')[5]) for r in filas)
    with open(HIST_CSV, 'a', encoding='utf-8') as f:
        f.write("\n".join(filas) + "\n")
    log(f"✅ {MESES[mes]} {anio} archivado: {len(filas)} faenas, total {total:,.0f} m³")
    return 0

if __name__ == "__main__":
    sys.exit(main())
