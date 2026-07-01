#!/usr/bin/env python3
"""
generar_snapshots.py
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Genera el dashboard COMPLETO de meses pasados (snapshots) y los sube por FTP,
para que el selector de mes muestre el detalle real de cada mes — no solo el
resumen del Comparativo.

Por cada mes: baja PG + TP de Arauco → arma _snapshots/AAAA-MM/{data_diario,
data_tm, meta} → corre GENERAR_HTML.py --snapshot AAAA-MM → sube
Dashboard_Cosecha_AAAA-MM.html por FTP.

Idempotente: salta los meses cuyo HTML ya está en el servidor (a menos de
pasar --force). Pensado para la nube (sin disco persistente) y para backfill.

Uso:
    python3 generar_snapshots.py 2026-01 2026-02 ...   # meses específicos
    python3 generar_snapshots.py --todos               # todos los del histórico
    python3 generar_snapshots.py --todos --force        # regenerar aunque existan
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
import os, sys, csv, json, calendar, ftplib, subprocess, importlib.util
from pathlib import Path
from datetime import datetime
import pandas as pd
import urllib3
urllib3.disable_warnings()

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))
from normalizar_produccion import normalizar  # noqa: E402

TEAM_MAP = {
    'S123':'Millalemu 1.1','S58':'Millalemu 1.2','S223':'Millalemu 1.3',
    'S246':'Millalemu 1.4','MG5':'Millalemu 5','TEA02':'Millalemu 7',
    'TEA08':'Millalemu 9','T125':'Millalemu 11','TEA30':'Millalemu 1.3',
}
CLASIF = {
    1:'Mantención',2:'Mantención',3:'Mantención',4:'Mantención',5:'Mantención',
    6:'Mantención',7:'Mantención',8:'Mantención',10:'Mantención',12:'Mantención',
    58:'Mantención',69:'Mantención',
    13:'Operacional',14:'Operacional',15:'Operacional',20:'Operacional',21:'Operacional',
    22:'Operacional',31:'Operacional',32:'Operacional',33:'Operacional',38:'Operacional',
    41:'Operacional',
    16:'Proceso',17:'Proceso',18:'Proceso',25:'Proceso',26:'Proceso',61:'Proceso',
    65:'Proceso',66:'Proceso',68:'Proceso',
    42:'Programado',43:'Programado',
}
METAS = {'Millalemu 1.1':7000.0,'Millalemu 1.2':7000.0,'Millalemu 1.3':7000.0,
         'Millalemu 1.4':7000.0,'Millalemu 5':7000.0,'Millalemu 7':7000.0,
         'Millalemu 9':7000.0,'Millalemu 11':6000.0}
MESES = ["","Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio",
         "Agosto","Septiembre","Octubre","Noviembre","Diciembre"]

FTP_HOST = os.environ.get('FTP_HOST') or "186.64.119.70"
FTP_USER = os.environ.get('FTP_USER') or "produccion@millalemu.com"
FTP_PASS = os.environ.get('FTP_PASS') or "Produccion2026x"

def log(m): print(f"[snapshots] {m}", flush=True)

def descargar(mes, anio):
    spec = importlib.util.spec_from_file_location("dnoc", str(BASE_DIR / "descargar_noc_api.py"))
    dn = importlib.util.module_from_spec(spec); spec.loader.exec_module(dn)
    import requests
    s = requests.Session(); s.verify = False
    s.headers.update({"User-Agent": "Mozilla/5.0"})
    tok = dn.obtener_token_arcgis(s)
    ud = calendar.monthrange(anio, mes)[1]
    fi, ff = f"{anio}-{mes:02d}-01", f"{anio}-{mes:02d}-{ud:02d}"
    pg = dn.descargar_reporte(s, tok, "BN", fi, ff)
    tp = dn.descargar_reporte(s, tok, "TP", fi, ff)
    d = Path("/tmp/snap_dl"); d.mkdir(exist_ok=True)
    if pg: dn.guardar_csv(pg, "BN", d)
    if tp: dn.guardar_csv(tp, "TP", d)
    pgdf = pd.read_csv(d / "Base2NOC.csv", sep=';', encoding='utf-8-sig') if pg else None
    tpdf = pd.read_csv(d / "TiemposPerdidos.csv", sep=';', encoding='utf-8-sig') if tp else None
    return pgdf, tpdf

def construir_snapshot(mes, anio, pg, tp):
    snap_dir = BASE_DIR / "_snapshots" / f"{anio}-{mes:02d}"
    snap_dir.mkdir(parents=True, exist_ok=True)

    # ── data_diario.csv (dedup HrsEf por folio, igual que el dashboard) ──
    prod = normalizar(pg)
    for c in ['Volumen SSC PU', 'Volumen SSC AS']:
        prod[c] = pd.to_numeric(prod[c].astype(str).str.replace(',', '.'), errors='coerce')
    prod['Vol'] = prod['Volumen SSC PU'].fillna(0) + prod['Volumen SSC AS'].fillna(0)
    _te = pd.to_numeric(prod['Tiempo Efectivo'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
    _pos = _te[_te > 0]
    prod['HrsEf'] = _te / (3600 if (len(_pos) > 0 and _pos.median() > 1000) else 60)
    prod['Arb'] = pd.to_numeric(prod.get('Árboles Madereados', 0), errors='coerce').fillna(0)
    prod['Ciclos'] = pd.to_numeric(prod.get('Número Ciclos', 0), errors='coerce').fillna(0)
    prod['Team'] = prod['Equipo'].map(TEAM_MAP)
    prod['Fecha_dt'] = pd.to_datetime(prod['Fecha NOC'], dayfirst=True, errors='coerce')
    prod['Dia'] = prod['Fecha_dt'].dt.day
    prod = prod[prod['Fecha_dt'].dt.month == mes].dropna(subset=['Team'])
    if 'Número Noc' in prod.columns:
        pf = prod.groupby(['Dia', 'Team', 'Número Noc']).agg(
            Vol=('Vol', 'sum'), HrsEf=('HrsEf', 'first'),
            Arb=('Arb', 'sum'), Ciclos=('Ciclos', 'sum')).reset_index()
    else:
        pf = prod
    agg = pf.groupby(['Dia', 'Team']).agg(
        Vol_m3=('Vol', 'sum'), Hrs_Ef=('HrsEf', 'sum'),
        Arboles=('Arb', 'sum'), Ciclos=('Ciclos', 'sum')).reset_index()
    agg['Arboles'] = agg['Arboles'].astype(int); agg['Ciclos'] = agg['Ciclos'].astype(int)
    agg['TM_Mant'] = 0; agg['TM_Oper'] = 0; agg['TM_PP'] = 0; agg['TM_Total'] = 0
    agg.to_csv(snap_dir / "data_diario.csv", sep=';', encoding='utf-8-sig', index=False)

    # ── data_tm.csv ──
    tp = tp.copy()
    tp['F'] = pd.to_datetime(tp['Fecha'], dayfirst=True, errors='coerce')
    tp = tp[tp['F'].dt.month == mes]
    tp['Dia'] = tp['F'].dt.day
    tp['Team'] = tp['Código Equipo'].map(TEAM_MAP)
    tp['Clasif'] = pd.to_numeric(tp['Código Tiempo Perdido'], errors='coerce').map(CLASIF).fillna('Operacional')
    out = tp.dropna(subset=['Team'])[['Dia', 'Team', 'Código Tiempo Perdido', 'Descripción',
                                       'Clasif', 'Tiempo (Min)', 'Observación']].copy()
    out.columns = ['Dia', 'Team', 'Codigo', 'Descripcion', 'Clasificacion', 'Minutos', 'Observacion']
    out.to_csv(snap_dir / "data_tm.csv", sep=';', encoding='utf-8-sig', index=False)

    # ── meta.json ──
    meta = {'mes': mes, 'anio': anio, 'mes_nombre': MESES[mes],
            'dias_mes': calendar.monthrange(anio, mes)[1], 'dias_no_trab': 0,
            'dias_trabajados': calendar.monthrange(anio, mes)[1], 'metas': METAS,
            'fuente': 'Re-descarga Arauco', 'archivado_en': datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
    with open(snap_dir / "meta.json", 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    return len(agg), len(out)

def existe_en_ftp(nombre):
    try:
        ftp = ftplib.FTP(FTP_HOST, timeout=30); ftp.login(FTP_USER, FTP_PASS)
        existe = nombre in ftp.nlst()
        ftp.quit(); return existe
    except Exception as e:
        log(f"aviso: no pude listar FTP ({e})"); return False

def subir_ftp(local_path, nombre):
    ftp = ftplib.FTP(FTP_HOST, timeout=60); ftp.login(FTP_USER, FTP_PASS)
    with open(local_path, 'rb') as f:
        ftp.storbinary(f"STOR {nombre}", f)
    ftp.quit()

def generar_mes(mes, anio, force=False):
    nombre = f"Dashboard_Cosecha_{anio}-{mes:02d}.html"
    if not force and existe_en_ftp(nombre):
        log(f"{MESES[mes]} {anio} ya está en el servidor — skip.")
        return
    log(f"Generando snapshot {MESES[mes]} {anio}...")
    pg, tp = descargar(mes, anio)
    if pg is None or tp is None or pg.empty:
        log(f"sin datos de {MESES[mes]} {anio} — skip."); return
    nd, ntm = construir_snapshot(mes, anio, pg, tp)
    res = subprocess.run([sys.executable, str(BASE_DIR / "GENERAR_HTML.py"),
                          "--snapshot", f"{anio}-{mes:02d}"], capture_output=True, text=True)
    out_html = BASE_DIR / nombre
    if res.returncode != 0 or not out_html.exists():
        log(f"❌ GENERAR_HTML falló: {res.stderr[-500:]}"); return
    subir_ftp(str(out_html), nombre)
    log(f"✅ {nombre} generado ({nd} días, {ntm} TM) y subido.")

def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    force = '--force' in sys.argv
    meses = []
    if '--todos' in sys.argv:
        hist = BASE_DIR / "historico_tm_mensual.csv"
        if hist.exists():
            for r in csv.DictReader(open(hist, encoding='utf-8-sig'), delimiter=';'):
                meses.append((int(r['mes']), int(r['anio'])))
    for a in args:
        try:
            y, m = a.split('-'); meses.append((int(m), int(y)))
        except Exception:
            log(f"argumento inválido: {a}")
    meses = sorted(set(meses), key=lambda x: (x[1], x[0]))
    if not meses:
        log("nada que generar (pasa AAAA-MM o --todos)"); return 0
    for mes, anio in meses:
        try:
            generar_mes(mes, anio, force=force)
        except Exception as e:
            log(f"❌ error en {MESES[mes]} {anio}: {e}")
    return 0

if __name__ == "__main__":
    sys.exit(main())
