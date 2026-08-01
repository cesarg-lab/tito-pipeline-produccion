#!/usr/bin/env python3
"""
generar_grafico_desplazamiento.py — imagen diaria del arrastre por faena, para Telegram.

Es el ÚNICO indicador del GPS que no está en ninguna otra salida. La grilla que ya se manda
lleva acumulado, meta, proyección y brecha; esto no las repite.

Sale HTML y lo rasteriza run_pipeline.sh con el Chrome que ya usa para los PDF — así reusa el
mismo lenguaje visual del informe en vez de una segunda librería de gráficos.

Lee wialon_km.json (lo deja descargar_wialon.py) y kpis.json (compute_kpis.py). Sin cualquiera
de los dos no dibuja nada y avisa: una imagen a medias es peor que ninguna.
"""
import json
import sys
from pathlib import Path

BASE = Path(__file__).parent
SALIDA = BASE / "grafico_desplazamiento.html"

# Solo SG y HM: las máquinas para las que desplazarse ES el trabajo. Ver descargar_wialon.py.
MAQUINA = {'SKIDDER': 'SG', 'GRAPPLE': 'SG', 'SHOVEL': 'HM'}
VEL_CAMA_BAJA = 13.0     # km/h — sobre esto el equipo iba en camión, no trabajando
MIN_DIAS = 21            # bajo esta cobertura el promedio se marca como flojo

NOMBRE = {'M1.1': 'M1.1', 'M1.2': 'M1.2', 'M1.3': 'M1.3', 'M1.4': 'M1.4',
          'M5': 'M5', 'M7': 'M7', 'M9': 'M9', 'M11': 'M11'}
ORDEN = ['M1.1', 'M1.2', 'M1.3', 'M1.4', 'M5', 'M7', 'M9', 'M11']


def agregar(wia, mes_key):
    """km/día por faena y máquina, sobre los días CON movimiento y sin los de cama baja."""
    res = {}
    for v in (wia or {}).values():
        m = MAQUINA.get(v.get('tipo'))
        if not m:
            continue
        a = res.setdefault((v['faena'], m), {'km': 0.0, 'dias': 0})
        for fecha, d in (v.get('dias') or {}).items():
            if str(fecha)[:7] != mes_key:
                continue
            km, h = d.get('km', 0.0), d.get('h_mov', 0.0)
            if km <= 0.5 or h <= 0 or km / h > VEL_CAMA_BAJA:
                continue
            a['km'] += km
            a['dias'] += 1
    return {k: {'km_dia': a['km'] / a['dias'], 'dias': a['dias']}
            for k, a in res.items() if a['dias']}


def main():
    wia_p, kpi_p = BASE / "wialon_km.json", BASE / "kpis.json"
    if not wia_p.exists() or not kpi_p.exists():
        print("  ⚠️  falta wialon_km.json o kpis.json; no se genera el gráfico")
        return 0
    try:
        wia = json.loads(wia_p.read_text(encoding='utf-8'))
        kpi = json.loads(kpi_p.read_text(encoding='utf-8'))
    except Exception as e:
        print(f"  ⚠️  no legible ({e}); no se genera el gráfico")
        return 0

    # kpis.json trae `mes` como TEXTO ("Julio 2026") y el número aparte en `mes_num`/`anio`.
    # (Lo aprendí rompiendo producción: probé contra un kpis.json que yo mismo inventé con otra
    # forma, así que el test pasó y el runner falló. Comprobar la estructura REAL, no la
    # supuesta.) Si faltara, el mes se deduce del propio wialon.
    try:
        mes_key = f"{int(kpi['anio'])}-{int(kpi['mes_num']):02d}"
    except Exception:
        fechas = [f for v in wia.values() for f in (v.get('dias') or {})]
        mes_key = max(fechas)[:7] if fechas else None
    if not mes_key:
        print("  ⚠️  no se pudo determinar el mes; no se genera el gráfico")
        return 0

    agg = agregar(wia, mes_key)
    cumpl = {f['team']: f.get('proy_cumpl_pct', 0) for f in kpi.get('faenas', [])}
    filas = []
    for fa in ORDEN:
        sg, hm = agg.get((fa, 'SG')), agg.get((fa, 'HM'))
        if not sg and not hm:
            continue
        filas.append({'f': NOMBRE.get(fa, fa),
                      'sg': sg['km_dia'] if sg else None, 'sgD': sg['dias'] if sg else 0,
                      'hm': hm['km_dia'] if hm else None, 'hmD': hm['dias'] if hm else 0,
                      'cum': cumpl.get(fa, 0)})
    if not filas:
        print("  ⚠️  sin datos de GPS del mes; no se genera el gráfico")
        return 0
    filas.sort(key=lambda x: -(x['sg'] or 0))

    MESES = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto',
             'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    titulo = f"{MESES[int(mes_key[5:7])]} {mes_key[:4]}"

    SALIDA.write_text(PLANTILLA.replace('/*DATOS*/', json.dumps(filas, ensure_ascii=False))
                              .replace('{MES}', titulo)
                              .replace('{MIN_DIAS}', str(MIN_DIAS)), encoding='utf-8')
    print(f"  ✅ grafico_desplazamiento.html — {len(filas)} faenas, {titulo}")
    return 0


# Mismo lenguaje visual del informe. Colores categóricos validados contra daltonismo
# (azul ↔ naranjo, ΔE 24,7 en protanopia) y semáforo del proyecto en la columna de
# cumplimiento: VERDE solo ≥90.
PLANTILLA = """<!doctype html><html lang=es><head><meta charset=utf-8><style>
:root{--surface:#fcfcfb;--ink1:#1c2024;--ink2:#4a5259;--ink3:#7b858d;--rule:#e4e7ea;
      --grid:#eef0f2;--sg:#2a78d6;--hm:#eb6834}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--surface);font-family:'IBM Plex Sans',system-ui,sans-serif;
     color:var(--ink1);padding:30px 34px;width:1080px}
h1{font-size:20px;font-weight:700;letter-spacing:-.01em}
.sub{font-size:12.5px;color:var(--ink2);margin-top:3px}
.leyenda{display:flex;gap:20px;align-items:center;margin:16px 0 2px;font-size:12px;color:var(--ink2)}
.leyenda i{display:inline-block;width:20px;height:9px;border-radius:2px;margin-right:6px;vertical-align:-1px}
table{border-collapse:collapse;width:100%;margin-top:8px}
th{font-size:10.5px;font-weight:600;color:var(--ink3);text-transform:uppercase;
   letter-spacing:.05em;text-align:right;padding:0 0 7px;border-bottom:1px solid var(--rule)}
th.l{text-align:left}th.b{text-align:left;padding-left:2px}
td{font-size:13px;padding:6px 0;border-bottom:1px solid var(--grid);text-align:right;
   font-variant-numeric:tabular-nums}
td.l{text-align:left;font-weight:600;font-size:14px}
td.b{padding:6px 16px 6px 2px;width:52%}
.par{display:flex;flex-direction:column;gap:3px}
.fila{display:flex;align-items:center;gap:7px;height:11px}
.barra{height:100%;border-radius:2px;min-width:2px}
.cifra{font-size:11px;font-variant-numeric:tabular-nums;color:var(--ink2);white-space:nowrap}
.vacio{font-size:10.5px;color:var(--ink3);font-style:italic}
.cob{font-size:11px;color:var(--ink3);font-variant-numeric:tabular-nums}
.flaco{color:#C87A18;font-weight:600}
.nota{margin-top:14px;font-size:11.5px;color:var(--ink2);line-height:1.6;
      border-left:3px solid var(--rule);padding-left:11px}
.nota b{color:var(--ink1)}
</style></head><body>
<h1>Desplazamiento diario · {MES}</h1>
<div class=sub>Kilómetros que recorre cada máquina en un día de faena, del GPS ·
  sin los días de traslado en cama baja</div>
<div class=leyenda>
  <span><i style="background:var(--sg)"></i>SG · skidder (arrastra a cancha)</span>
  <span><i style="background:var(--hm)"></i>HM · shovel (acomoda en el frente)</span>
</div>
<table><tr><th class=l>Faena</th><th class=b>km por día trabajado</th>
  <th>Días medidos</th><th>Cumpl. mes</th></tr><tbody id=filas></tbody></table>
<div class=nota id=nota></div>
<script>
const D=/*DATOS*/, MIN={MIN_DIAS};
const MAX=Math.max(25,...D.map(d=>d.sg||0))*1.05;
const fx=v=>v.toFixed(1).replace('.',',');
document.getElementById('filas').innerHTML=D.map(d=>{
  const b=(v,c)=> v==null?'<span class=vacio>sin GPS en este proceso</span>'
    :`<div class=barra style="width:${v/MAX*100}%;background:${c}"></div><span class=cifra>${fx(v)} km</span>`;
  const dias=d.hm!=null?`${d.sgD} · ${d.hmD}`:`${d.sgD}`, flaco=d.sgD<MIN;
  const col=d.cum>=90?'#1A8060':d.cum>=60?'#C87A18':'#A32A22';
  return `<tr><td class=l>${d.f}</td><td class=b><div class=par>
    <div class=fila>${b(d.sg,'var(--sg)')}</div><div class=fila>${b(d.hm,'var(--hm)')}</div>
    </div></td><td class="cob${flaco?' flaco':''}">${dias}${flaco?' \\u26a0':''}</td>
    <td style="font-weight:600;color:${col}">${d.cum.toFixed(0)}%</td></tr>`}).join('');
// La nota se arma con los datos del día, no con un texto fijo que envejece.
const con=D.filter(d=>d.sg!=null).sort((a,b)=>b.sg-a.sg);
const alto=con[0], bajo=con[con.length-1];
const flacos=D.filter(d=>d.sgD<MIN).map(d=>`${d.f} (${d.sgD} días)`);
document.getElementById('nota').innerHTML =
  (alto&&bajo&&alto!==bajo ? `<b>${alto.f}</b> es la que más arrastra (${fx(alto.sg)} km/día,
     cierra en ${alto.cum.toFixed(0)}%) y <b>${bajo.f}</b> la que menos (${fx(bajo.sg)},
     ${bajo.cum.toFixed(0)}%).<br>` : '')
  + (flacos.length ? `<b>Cobertura floja</b> en ${flacos.join(' · ')}: el promedio se apoya en
     parte del mes.` : 'Cobertura completa en todas las faenas medidas.');
</script></body></html>"""


if __name__ == "__main__":
    sys.exit(main())
