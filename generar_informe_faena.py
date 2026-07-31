#!/usr/bin/env python3
"""
generar_informe_faena.py — INFORME ÚNICO por faena con la FORMA del "Tablero de Gestión
Diaria de Faena" de Arauco (la pizarra física), SOLO la mitad de PRODUCTIVIDAD.

Fusiona en un solo descargable/imprimible lo que hoy son dos pestañas (Tablero + Tablas de
Productividad). 1 hoja A4 por faena (page-break) + botón "Imprimir / Guardar PDF".
Es una GUÍA que el jefe de faena LLENA en terreno: pre-llena lo que el NOC ya sabe (trozado
real por día, metas, ritmo del procesador, VMA) y deja "por reportar" lo que solo el jefe sabe
(volteo, madereo en m³, tiempos perdidos, acta, stock).

QUEDA FUERA (SSO): Fichas IAP, Nivel de Madurez, Peligros/Riesgos, Tarea Crítica, Mapa de Riesgo.

Uso:  python3 generar_informe_faena.py [pg_historico.json]
      (sin argumento baja PG del mes por API vía descargar_noc_api — igual que el tablero)
Salida: Informe_Faena.html
"""
import sys, calendar
from pathlib import Path
import pandas as pd, numpy as np

# Reutiliza TODO lo ya probado del generador del tablero (mismo módulo, mismos datos).
from generar_tablero_faena import (
    base_diaria, tabla_p75, teorico, metas_excel, metas_procesos, cargar_pg,
    CSS, TEC_NORM, TEAM_MAP, NOMBRE, ESP, ESPN, TECN, LB, TR, TRAMO_MID,
    USO, HDISP, METAS_DEFAULT, LOGO,
)

BASE = Path(__file__).parent
FAENA_ORDER = ['M1.1','M1.2','M1.3','M1.4','M5','M7','M9','M11']

# Código de salida cuando el CMMS RECHAZA la credencial (401/403). run_pipeline.sh lo
# distingue de cualquier otra falla del informe.
#
# Por qué existe: el 30-07-2026 un blindaje de seguridad le quitó a `anon` el EXECUTE de las
# tres RPC del informe. Las llamadas son best-effort a propósito (si el CMMS no responde el
# informe igual sirve con lo del NOC), así que el pipeline siguió marcando ✅ durante dos días
# mientras publicaba un informe sin volteo, sin madereo, sin tiempos perdidos y sin horas de
# pre-uso — la mitad del documento. Un 401 no es "un dato que falta": es configuración rota, y
# tiene que hacer ruido.
EXIT_CMMS_AUTH = 3

# Zonas para agrupar los PDF que van por Telegram. MISMO criterio que GENERAR_IMAGEN.py
# (grupos 'aereo' / 'terrestre') — si cambia allá, cambiar acá.
ZONAS = {
    'Aereo':     ['M5', 'M7', 'M9', 'M11'],
    'Terrestre': ['M1.1', 'M1.2', 'M1.3', 'M1.4'],
}

# ── CSS extra del informe (encima del CSS del tablero) ──────────────────────
CSS_INFORME = """
button.noprint{position:fixed;top:12px;right:14px;z-index:99;background:#417505;color:#fff;
  border:0;border-radius:8px;padding:9px 16px;font-size:13px;font-weight:600;cursor:pointer;
  box-shadow:0 2px 8px rgba(0,0,0,.2);font-family:'IBM Plex Sans',sans-serif}
button.noprint:hover{background:#345d04}
.sheet{max-width:194mm}
.ig{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:4px}
.ig div{flex:1;min-width:96px;background:#f6f8fa;border:1px solid #e0e5ea;border-radius:5px;padding:4px 8px}
.ig span{font-size:8.5px;color:#778;text-transform:uppercase;display:block}
.ig b{font-size:12px;color:#1b3a05}
.ig .fill{color:#a06000;font-style:italic;font-weight:400}
td.bl{background:#fffdf5}                 /* celda en blanco para llenar */
td.pr,.pr{color:#a06000;font-style:italic}/* "por reportar" (lo informa el jefe) */
td.nf{background:#eaf3e0;font-weight:700;color:#2d5202} /* pre-llenado del NOC */
td.gu{background:#f4f7fb;color:#1A5276;font-weight:600} /* guía / teórico */
td.tp,.tp{background:#fdecea;color:#a01b0b;font-weight:600} /* tiempo perdido (preuso) */
td.sp{background:#eaf3e0;color:#2d5202;font-weight:700}    /* turno limpio CONFIRMADO */
td.nd{color:#b7bec6;font-size:6.5px;letter-spacing:-.2px}  /* sin pre-uso: nadie declaró */
td.jA{background:#e3ecf5;box-shadow:inset 3px 0 0 #4a7ba7}  /* turno del jefe A */
td.jB{background:#f7efe2;box-shadow:inset 3px 0 0 #b9863e}  /* turno del jefe B */
.leyj{font-size:8.5px;color:#4a5a6a;margin:3px 0 0}
.leyj i{display:inline-block;width:9px;height:9px;border-radius:2px;margin:0 3px -1px 9px}
.tpaclab{font-size:9.5px;color:#778;text-transform:uppercase;letter-spacing:.3px;margin:2px 0}
.ig .jefe{flex:2;min-width:180px}
.cand{display:inline-block;font-size:11px;font-weight:600;color:#1b3a05;margin-right:9px;
  white-space:nowrap}
.cob{font-size:9.5px;color:#4a5a6a;background:#f6f8fa;border-left:3px solid #7f9c5a;
  border-radius:4px;padding:4px 8px;margin-top:3px;line-height:1.45}
.cob b{color:#2d5202}
table.tpac{width:auto;min-width:60%;margin-bottom:6px}
.diaria{font-size:7.6px}
.diaria td,.diaria th{padding:1px 2px}
.diaria th{font-size:7.4px}
.grp1{background:#eef3f8}.grp2{background:#f3f0ea}
table.prod{margin-bottom:6px}
th.proc{background:#1b4f72;color:#fff;letter-spacing:.06em;font-size:10px}
tr.hoy td{background:#fff3cf!important;box-shadow:inset 0 0 0 1px #e0a800}
tr.tot td{background:#eef1f4;border-top:2px solid #1b3a05;font-weight:700}
tr.hoy td.l{font-weight:700;color:#7a5c00}
.two{display:flex;gap:10px;align-items:flex-start}
.two>div{flex:1}
.guia{background:#f4f7fb;border:1px solid #d3ddea;border-left:4px solid #1A5276;border-radius:8px;
  padding:7px 11px;font-size:11px;color:#33475b;line-height:1.5;margin-top:4px}
.guia b{color:#1A5276}
.recu{background:#fff7e6;border:1px solid #e08e0b;border-radius:5px;padding:5px 9px;color:#8a5a00;
  font-size:11px;margin-top:4px}
.q{font-size:10.5px;color:#3a4a5a;margin:5px 0 2px}
.q b{color:#233}
.barra{position:sticky;top:0;z-index:99;display:flex;gap:12px;align-items:center;flex-wrap:wrap;
  background:#1b3a05;color:#fff;padding:9px 16px;font-family:'IBM Plex Sans',sans-serif;
  box-shadow:0 2px 8px rgba(0,0,0,.2);margin-bottom:10px}
.barra label{font-size:13px;font-weight:600;display:flex;align-items:center;gap:6px}
.barra select{font-size:13px;padding:5px 8px;border-radius:6px;border:0;font-family:inherit;min-width:190px}
.barra button{background:#417505;color:#fff;border:0;border-radius:8px;padding:8px 15px;
  font-size:13px;font-weight:600;cursor:pointer;font-family:inherit}
.barra button:hover{background:#5a9e0a}
.barra .hint{font-size:11px;opacity:.85;font-weight:400}
.oculto{display:none!important}
.analisis{page-break-before:always}
/* El registro del mes va en su propia hoja: antes el corte caía a mitad de la tabla
   diaria, partiendo el mes en dos por donde tocara. Ahora el quiebre es a propósito. */
.registro{page-break-before:always}
@media print{button.noprint{display:none!important}.barra{display:none!important}
  .diaria{font-size:7px}.sheet{padding:8px 10px}}
"""

MESES = ['','Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto',
         'Septiembre','Octubre','Noviembre','Diciembre']


# ── Referencia OFICIAL de Arauco: ritmo y carga esperados por tecnología × especie ──────────
# Fuente: "Productividad Faena, KPI.xlsx" (hoja Madereo, bloque "Datos Referenciales").
#
# Decisión de gerencia 2026-07-30: ESTA es la columna "Plan". Antes el plan salía del p75 de
# nuestro propio historial, o sea la faena se comparaba consigo misma y una faena mediocre
# aparecía "sobre el plan" por el solo hecho de repetir su promedio. Ese p75 sigue a la vista
# como "Habitual" — sirve para saber qué se logra hoy, no para fijar la meta.
#
# OJO: ritmo × carga da el mismo rendimiento en las dos especies del skidder (8×5,2 = 13×3,2 =
# 41,6 m³/hr): Arauco reparte distinto entre viajes y carga según el árbol, pero espera la misma
# productividad. Forwarder y Clambunk no están en la flota; van por completitud de la tabla.
REF_ARAUCO = {
    ('TORRE',               'EUGL'): (19.0,  1.2),
    ('TORRE',               'PIRA'): (12.0,  2.0),
    ('TORRE',               'EUNI'): (15.0,  1.8),
    ('SKIDDER 6X6 GRAPPLE', 'EUGL'): (13.0,  3.2),
    ('SKIDDER 6X6 GRAPPLE', 'PIRA'): ( 8.0,  5.2),
    ('SKIDDER 6X6 GRAPPLE', 'EUNI'): ( 8.0,  5.0),
    ('FORWARDER',           'EUGL'): ( 2.0, 16.0),
    ('FORWARDER',           'PIRA'): ( 1.6, 15.0),
    ('FORWARDER',           'EUNI'): ( 2.5, 15.0),
    ('CLAMBUNK',            'PIRA'): ( 2.3, 14.0),
}


def ref_arauco(tec, esp):
    """(ritmo, carga, rendimiento) de la referencia de Arauco. None si esa combinación no está
    en la tabla — el informe muestra "—" y nunca un número inventado."""
    v = REF_ARAUCO.get((str(tec).upper(), str(esp).upper()))
    if not v:
        return None
    ritmo, carga = v
    return ritmo, carga, ritmo * carga


# Feriados IRRENUNCIABLES (los únicos días que la faena no puede trabajar por ley).
# ESPEJO de `_FERIADOS_IRR` en compute_kpis.py.
FERIADOS_IRR = {'01-01', '05-01', '09-18', '09-19', '12-25'}


def dias_operables(anio, mes):
    """**Se trabaja TODOS los días, domingos incluidos** (gerencia 2026-07-25); solo se
    descuentan los feriados irrenunciables.

    MISMO criterio que DT/DD/DR de compute_kpis.py — si cambia uno, cambiar el otro. Antes el
    informe excluía los domingos y los KPIs no: en julio 2026 eso repartía la meta de M7 en 27
    días (299 m³/día) contra 31 (260 m³/día) de la pestaña KPIs, un 15% de diferencia entre dos
    números del mismo tablero."""
    n = calendar.monthrange(anio, mes)[1]
    ops = [d for d in range(1, n+1) if f"{mes:02d}-{d:02d}" not in FERIADOS_IRR]
    return n, ops


def fmt(x, dec=0):
    try:
        return f"{x:,.{dec}f}"
    except Exception:
        return "—"


def prod_general(jul, meta_mes, anio, mes, ult_dia, kpi=None, dr_kpi=None):
    """Producción General. Si hay registro del kpis.json (kpi) usa su AVANCE y PROYECCIÓN para
    CALZAR con la pestaña KPIs; si no, recalcula. Devuelve dict con todos los campos."""
    n_mes, ops = dias_operables(anio, mes)
    n_op = max(len(ops), 1)
    op_hasta = max(len([d for d in ops if d <= ult_dia]), 1)  # días operables transcurridos
    real_dia = float(jul[jul.dia == jul.dia.max()].m3.sum())
    plan_diario = meta_mes / n_op
    avance_plan = plan_diario * op_hasta                       # dónde deberías ir (plan lineal)
    if kpi:                                                    # ← congruente con KPIs
        acum = float(kpi.get('vol_m3', jul.m3.sum()))
        proy = float(kpi.get('proy_cierre_m3', 0))
        dias_rest = max(int(dr_kpi) if dr_kpi else (n_op - op_hasta), 1)
    else:
        acum = float(jul.m3.sum())
        diast = max(int(jul.dia.nunique()), 1)
        dias_rest = max(n_op - op_hasta, 1)
        proy = acum + (acum / diast) * dias_rest               # = ritmo real × días restantes
    cumpl = acum / avance_plan * 100 if avance_plan else 0
    recuperar = max(0.0, (avance_plan - acum)) / dias_rest
    # Meta día DINÁMICA: lo que la meta exige por día de HOY a fin de mes, según lo YA procesado.
    # Cambia cada día con el real acumulado (si vas atrasado sube, si vas adelantado baja).
    meta_dia_req = max(0.0, (meta_mes - acum)) / dias_rest
    return dict(meta_mes=meta_mes, avance_plan=avance_plan, avance_real=acum, cumpl=cumpl,
                proy=proy, plan_diario=plan_diario, real_diario=real_dia,
                recuperar=recuperar, meta_dia_req=meta_dia_req, dias_rest=dias_rest,
                n_op=n_op, op_hasta=op_hasta, cumple_plan=(cumpl >= 100))


def plan_productividad(fa, jul, cell):
    """Plan (guía VMA/especie) + Real (NOC) por proceso de madereo de la faena."""
    jd = jul.dropna(subset=['VMA','rend']); jd = jd[jd.hrs >= 3].copy()
    if len(jd):
        jd['tramo'] = pd.cut(jd.VMA, TR, labels=LB)
        p75s = [cell[k]['p75'] for k in ((r.tec, r.especie, r.tramo) for _, r in jd.iterrows()) if k in cell]
        cargas = [cell[k]['carga'] for k in ((r.tec, r.especie, r.tramo) for _, r in jd.iterrows()) if k in cell]
        ritmos = [cell[k]['ritmo'] for k in ((r.tec, r.especie, r.tramo) for _, r in jd.iterrows()) if k in cell]
    else:
        p75s = cargas = ritmos = []
    plan_rend = round(np.median(p75s), 1) if p75s else round(jul.rend.median(), 1)
    plan_carga = round(np.median(cargas), 2) if cargas else round(jul.carga.median(), 2)
    plan_ritmo = round(np.median(ritmos), 2) if ritmos else round(jul.ritmo.median(), 2)
    # REAL = ACUMULADO del mes (gerencia 2026-07-25), no la mediana de los días: mediana de
    # ratios ≠ ratio de totales, y daba números que contradecían a la hoja de análisis y a la
    # pestaña KPIs (M7: carga 4,62 por mediana contra 1,06 por acumulado). Con el acumulado las
    # tres vistas hablan del mismo número.
    _m3 = float(jul.m3.sum()); _hrs = float(jul.hrs.sum())
    real_rend = (_m3 / _hrs) if _hrs > 0 else float('nan')

    # CARGA y RITMO se calculan descartando los días con ciclos FUERA DE RANGO FÍSICO. Los
    # ciclos los declara la faena al NOC y se detectaron dos problemas reales (2026-07-25):
    # un dedazo en M7 el día 19 (381 ciclos/hr) que solo arrastraba el mes entero de 4 a 18, y
    # una diferencia sistemática de criterio entre los dos turnos de M1.4 (uno declara 7-8
    # ciclos/hr y el otro 20-43, todos los días de su semana). Sin filtro, la "palanca
    # limitante" del análisis apuntaba al problema equivocado.
    #
    # El techo es por TECNOLOGÍA y no relativo a la mediana de la faena: cuando la mayoría de
    # los días viene mal (M1.4), la mediana también está mal y el filtro descartaría los días
    # buenos. La torre tiene ciclos legítimamente más cortos que un skidder (M11 anda en 12-19
    # ciclos/hr y es correcto), por eso su techo es mayor.
    #
    # NO afecta al volumen ni al rendimiento: la producción del mes se calcula con TODOS los
    # días. Solo se limpia el divisor de estos dos indicadores.
    _tope = 30.0 if str(jul.tec.mode().iat[0] if len(jul.tec.mode()) else '').upper() == 'TORRE' else 15.0
    _d = jul[(jul.hrs > 0) & (jul.ciclos > 0)].copy()
    _d['_r'] = _d.ciclos / _d.hrs
    _ok = _d[(_d._r >= 0.5) & (_d._r <= _tope)]
    dias_fuera = int(len(_d) - len(_ok))
    _cic = float(_ok.ciclos.sum()); _hrs_ok = float(_ok.hrs.sum()); _m3_ok = float(_ok.m3.sum())
    real_carga = (_m3_ok / _cic) if _cic > 0 else float('nan')
    real_ritmo = (_cic / _hrs_ok) if _hrs_ok > 0 else float('nan')
    return dict(plan_rend=plan_rend, plan_carga=plan_carga, plan_ritmo=plan_ritmo,
                real_rend=real_rend, real_carga=real_carga, real_ritmo=real_ritmo,
                dias_fuera=dias_fuera, dias_ok=int(len(_ok)))


def horas_preuso(fa, cmms, real_por_dia, mes_key):
    """Uso Real y Rend Real por proceso, medidos con el HORÓMETRO DEL PRE-USO (RPC
    informe_horas_faena: Δ horómetro entre dos pre-usos de días consecutivos = un turno).

    · Uso real [%] = horas trabajadas ÷ (jornada 10,5 h × equipos-día medidos). Es un promedio
      por equipo-día, así que vale aunque solo una parte del proceso haya hecho pre-uso.
    · Rend real [m³/hr] = m³ del NOC ÷ horas, SOLO de los días con dato, y SOLO si ese día se
      midió el proceso COMPLETO (equipos = dotación): el m³ del NOC es de toda la faena, y
      dividirlo por las horas de un equipo de tres inflaría el rendimiento.
      Aplica al procesado/clasificado, que es lo que el NOC mide en m³.

    Devuelve {proceso: {horas, eq_dia, dias, turnos, uso, rend, rend_dias}}; vacío si no hay
    pre-uso con tramo de un día (el informe muestra "rep." y nunca un número reconstruido).
    """
    h = (cmms or {}).get('horas', {}).get(FAENA_ID.get(fa), {})
    out = {}
    # La RPC trae el mes actual Y el anterior: sin filtrar, en agosto se sumarían las horas de
    # julio y el día 15 de los dos meses caería en la misma celda.
    for (fecha, proc), v in h.items():
        if str(fecha)[:7] != mes_key:
            continue
        dia = int(str(fecha)[8:10])
        a = out.setdefault(proc, {'horas': 0.0, 'eq_dia': 0, 'turnos': 0,
                                  'dias': set(), 'dias_full': set(), 'horas_full': 0.0})
        a['horas'] += v['horas']
        a['eq_dia'] += v['equipos']
        a['turnos'] += v['equipos']
        a['dias'].add(dia)
        if v['equipos'] >= v['dotacion']:          # proceso medido completo ese día
            a['dias_full'].add(dia)
            a['horas_full'] += v['horas']
    for proc, a in out.items():
        a['uso'] = (a['horas'] / (HDISP * a['eq_dia']) * 100) if a['eq_dia'] else None
        # El m³ del NOC es TROZADO. Tiene sentido dividirlo por las horas del PROCESADO (el PM
        # que lo produjo) y por las del CLASIFICADO: si la GM no está en pana clasifica TODA la
        # madera que trozó el PM, así que su m³ del día es el mismo del NOC. Contra las horas
        # del volteo o del madereo daría un número sin significado (dio 106 m³/h en M7).
        if proc in ('PROCESADO', 'CLASIFICADO'):
            m3 = sum(real_por_dia.get(d, 0.0) for d in a['dias_full'])
            a['rend'] = (m3 / a['horas_full']) if (a['horas_full'] and m3) else None
        else:
            a['rend'] = None
        a['rend_dias'] = len(a['dias_full'])
    return out


def hoja_kpis(fa, kpi, dr, pp=None, bench=None, tm=None):
    """HOJA 2 — análisis. Lo que ya calcula compute_kpis por faena y el informe no mostraba:

      · cómo viene CADA TURNO (atribución 7×7 por jefe: m³/día y m³/hr) — la hoja 1 dice quién
        está hoy, esta dice cómo le va a cada uno;
      · la PALANCA que más limita (Uso, Ritmo o Carga) contra el mejor de su MISMA tecnología,
        con la brecha y los m³ que se ganarían cerrándola de aquí a fin de mes;
      · los tiempos perdidos según el NOC (mantención / operacional / proceso), que es la vista
        del cliente y contrasta con la del pre-uso de la hoja 1.

    Sin registro en kpis.json devuelve "" y el informe queda con una sola hoja."""
    if not kpi:
        return ""
    # Carga y Ritmo salen de la HOJA 1 (ya sin los días de ciclos mal declarados), NO de
    # kpis.json: si no, el mismo concepto saldría con dos cifras en el mismo documento — que es
    # justo lo que veníamos corrigiendo. El benchmark se recalcula con ese mismo criterio.
    # OJO: sin benchmark (M11 es la única torre, no tiene par de su tecnología) la hoja se
    # muestra IGUAL — solo se omite el bloque de palanca. Antes se devolvía "" y M11 perdía su
    # página entera de análisis, incluidos turnos y tiempos perdidos, que sí tiene.
    bt = dict(bench or {})
    mio = {'Uso': kpi.get('uso_pct'),
           'Ritmo': (pp or {}).get('real_ritmo'),
           'Carga': (pp or {}).get('real_carga')}
    gaps = {}
    for k in ('Uso', 'Ritmo', 'Carga'):
        ref, v = bt.get(k), mio.get(k)
        gaps[k] = round(max((ref - v) / ref * 100, 0), 1) if (ref and v and ref > 0) else 0.0
    pal = max(gaps, key=gaps.get) if gaps else None

    # ── Turnos (jefe A vs jefe B) ──
    turnos = kpi.get('turnos') or []
    if turnos:
        ft = ""
        mejor = max((t.get('m3_dia', 0) for t in turnos), default=0)
        for t in turnos:
            top = " style='background:#eaf3e0'" if t.get('m3_dia') == mejor and len(turnos) > 1 else ""
            ft += (f"<tr{top}><td class=l>{t.get('jefe','—')}</td><td>{t.get('dias','—')}</td>"
                   f"<td class=nf>{fmt(t.get('m3',0))}</td><td class=nf>{t.get('m3_dia','—')}</td>"
                   f"<td class=nf>{t.get('m3_hr','—')}</td></tr>")
        t_turnos = ("<table><tr><th class=l>Jefe de turno</th><th>Días</th><th>m³</th>"
                    "<th>m³/día</th><th>m³/hr</th></tr>" + ft + "</table>"
                    "<div class=cob>Atribución 7×7: cada jornada se asigna al jefe que estaba de "
                    "turno, con la misma rotación que usa la hoja 1. Verde = mejor m³/día del mes.</div>")
    else:
        t_turnos = "<div class=cob>Sin turnos atribuidos este mes.</div>"

    # ── Palanca limitante vs el mejor de su tecnología ──
    if pal and bt and any(bt.values()):
        filas = ""
        for nom, val, ref, dec in (
                ('Uso [%]', mio['Uso'], bt.get('Uso'), 0),
                ('Ritmo [ciclo/hr]', mio['Ritmo'], bt.get('Ritmo'), 2),
                ('Carga [m³/ciclo]', mio['Carga'], bt.get('Carga'), 2)):
            clave = nom.split(' ')[0]
            brecha = gaps.get(clave, 0)
            es_pal = clave == pal
            marca = " style='background:#fdecea;font-weight:700'" if es_pal else ""
            filas += (f"<tr{marca}><td class=l>{nom}{' ←' if es_pal else ''}</td>"
                      f"<td class=nf>{fmt(val, dec)}</td><td class=gu>{fmt(ref, dec)}</td>"
                      f"<td>{brecha:.1f}%</td></tr>")
        t_pal = ("<table><tr><th class=l>Palanca</th><th>Esta faena</th>"
                 "<th>Mejor de su tipo</th><th>Brecha</th></tr>" + filas + "</table>"
                 f"<div class=cob>La que más limita es <b>{pal}</b>: está "
                 f"<b>{gaps.get(pal, 0):.1f}%</b> bajo la mejor faena de su misma tecnología. "
                 f"Carga y Ritmo salen del mismo cálculo de la hoja 1 (sin los días con ciclos "
                 f"mal declarados), así que ambas hojas muestran la misma cifra.</div>")
    else:
        t_pal = ("<div class=cob>Esta faena es la única de su tecnología en la flota, así que no "
                 "hay contra quién comparar las palancas. Su referencia es su propio historial.</div>")

    # ── Tiempos perdidos del NOC (la vista del cliente) ──
    tmm, tmo, tmp = (kpi.get('tm_mant_min', 0), kpi.get('tm_oper_min', 0), kpi.get('tm_proc_min', 0))
    tot = tmm + tmo + tmp
    if tot:
        f_noc = ""
        for nom, mins in (('Mantención', tmm), ('Operacional', tmo), ('Proceso', tmp)):
            f_noc += (f"<tr><td class=l>{nom}</td><td class=tp>{mins/60:.1f}</td>"
                      f"<td>{mins/tot*100:.0f}%</td></tr>")
        t_noc = ("<table><tr><th class=l>Clasificación</th><th>Horas</th><th>%</th></tr>" + f_noc +
                 f"<tr><td class=l><b>Total</b></td><td class=tp><b>{tot/60:.1f}</b></td><td></td></tr>"
                 "</table><div class=cob>Tiempos perdidos <b>según el NOC</b> (lo que ve Arauco). "
                 "Contrastar con los del pre-uso de la hoja 1: si difieren mucho, algo no se está "
                 "declarando en alguno de los dos lados.</div>")
    else:
        t_noc = "<div class=cob>El NOC no registra tiempos perdidos de esta faena en el mes.</div>"

    # ── Evolución día a día (lo que salió de la hoja 1, acá en clave de análisis) ──
    dd = (kpi.get('dias_detalle') or [])
    if dd:
        rends = [x.get('rend', 0) for x in dd if x.get('rend')]
        mejor = max(rends) if rends else 0
        def _filas(items):
            out = ""
            for x in items:
                r = x.get('rend', 0)
                marca = " style='background:#eaf3e0;font-weight:700'" if r and r == mejor else ""
                out += (f"<tr{marca}><td class=l>{x.get('d',''):02d}</td><td class=nf>{fmt(x.get('vol',0))}</td>"
                        f"<td>{x.get('hrs','')}</td><td class=nf>{r}</td></tr>")
            return ("<table class=diaria><tr><th class=l>Día</th><th>m³</th><th>Hrs</th>"
                    "<th>Rend</th></tr>" + out + "</table>")
        # En DOS COLUMNAS: 24 filas seguidas desbordaban la hoja a una tercera página.
        mitad = (len(dd) + 1) // 2
        t_dd = (f"<div class=two><div>{_filas(dd[:mitad])}</div>"
                f"<div>{_filas(dd[mitad:])}</div></div>"
                "<div class=cob>Verde = el mejor rendimiento del mes: la jornada a igualar.</div>")
    else:
        t_dd = ""

    # ── Qué la detiene: top causas del NOC ──
    top = ((tm or {}).get(NOMBRE.get(fa, fa)) or {}).get('top_causas') or {}
    if top:
        tot_min = sum(top.values()) or 1
        f_top = ""
        for causa, mins in sorted(top.items(), key=lambda y: -y[1])[:6]:
            f_top += (f"<tr><td class=l>{causa[:46]}</td><td class=tp>{mins/60:.1f}</td>"
                      f"<td>{mins/tot_min*100:.0f}%</td></tr>")
        t_top = ("<table><tr><th class=l>Causa</th><th>Horas</th><th>% del total</th></tr>"
                 + f_top + "</table>")
    else:
        t_top = ""

    bloque_dd = f"<h2>Día a día del mes</h2>{t_dd}" if t_dd else ""
    bloque_top = f"<h2>Qué la detuvo — causas del NOC</h2>{t_top}" if t_top else ""

    return f"""<div class="sheet faena analisis" data-faena="{fa}">
<header>{'<img src="'+LOGO+'">' if LOGO else ''}<div>
<h1>Análisis de Faena · {NOMBRE.get(fa, fa)}</h1>
<div class=sub>KPIs Uso · Ritmo · Carga — comparación entre turnos y contra el mejor de su tecnología</div>
</div></header>
<h2>Cómo viene cada turno</h2>{t_turnos}
<h2>Palanca que más limita</h2>{t_pal}
<h2>Tiempos perdidos según el NOC</h2>{t_noc}
{bloque_top}
{bloque_dd}
<div class=foot>Hoja de ANÁLISIS: no se llena, se mira. Los datos salen del mismo kpis.json que
la pestaña KPIs del dashboard, así que ambos muestran las mismas cifras.</div>
</div>"""


def texto_tp(t):
    """Qué mostrar de un tiempo perdido. Con causa "Otro" la etiqueta no informa nada (todas las
    filas dirían "Otro"): manda la NOTA del operador, que es donde está el hecho. Con una causa
    real se muestra la causa, y la nota se agrega solo si aporta algo distinto."""
    causa = (t.get('causa') or '').strip()
    det = (t.get('detalle') or '').strip()
    if causa.lower() == 'otro':
        return det or causa
    return f"{causa} — {det}" if det else causa


def cargar_bn():
    """Lee Base2NOC.csv (reporte BN del NOC, que el pipeline ya descarga en el paso 1) y lo deja
    agrupado por FAENA. El PG que usa el resto del informe agrega por equipo/día y NO trae el
    desglose por producto, el acta, el stock ni el estado de la madera: eso solo viene en el BN.

    Sin archivo o ilegible → {} y las secciones quedan en blanco (best-effort: no rompe nada)."""
    import csv
    p = BASE / "Base2NOC.csv"
    if not p.exists():
        print("  ⚠️  Base2NOC.csv no está; acta y stock quedan para llenar a mano")
        return {}

    def num(s):
        try:
            return float(str(s).replace(',', '.'))
        except Exception:
            return 0.0

    out = {}
    try:
        with open(p, encoding='utf-8-sig', newline='') as f:
            for r in csv.DictReader(f, delimiter=';'):
                fa = TEAM_MAP.get(str(r.get('EQUIPO', '')).strip())
                if not fa:
                    continue
                out.setdefault(fa, []).append({
                    'fecha': str(r.get('FECHA', '')).strip(),          # dd-mm-yyyy
                    'producto': (r.get('PRODUCTO') or '').strip().upper(),
                    'destino': (r.get('NOMBRE_DESTINO') or r.get('CODIGO_DESTINO') or '').strip(),
                    'm3': num(r.get('M3SSC')),
                    'stock': num(r.get('STOCK')),
                    'acta': str(r.get('NUMERO_ACTA', '')).strip(),
                    # Condición del RODAL (Fresca / Quemada / Manchada). Llega en el BN desde
                    # siempre y no se usaba: es un dato de ENTRADA —con qué te tocó trabajar—,
                    # distinto del PRODUCTO, que es lo que salió. Explica por qué dos faenas
                    # con el mismo equipo y la misma especie rinden distinto, y relativiza la
                    # comparación contra la referencia de Arauco, que supone madera sana.
                    'estado': (r.get('ESTADO_MADERA') or '').strip(),
                })
    except Exception as e:
        print(f"  ⚠️  Base2NOC.csv no legible ({e}); acta y stock quedan en blanco")
        return {}
    return out


def tabla_acta(regs):
    """Cumplimiento Acta · UNA FILA POR ACTA, no un promedio del mes.

    El acta es la unidad del RODAL: es el trato con Arauco sobre un paño concreto. Una faena
    trabaja entre 2 y 6 actas en el mismo mes —M1.1 en dos predios distintos, M11 en seis
    actas— y esos rodales NO son comparables entre sí. Medido en el BN:

        M1.1 · acta 2078005 (predio 10073) → 98% fresca
              acta 2074817 (predio 11983) → 60% QUEMADA

    Promediarlos daba "81% fresca" para la faena, escondiendo que la mitad del mes se cosechó
    un paño quemado. Lo mismo con el mix de producto: M7 va de 61% pulpable en un acta a 100%
    aserrable en otra. Por eso acá se abre por acta y el total queda al pie como referencia.

    El **Plan del acta** (el mix comprometido con el cliente) sigue sin mostrarse: no es un dato
    de producción, el NOC no lo trae y no se inventa.
    """
    por_acta = {}
    for x in regs:
        if x['m3'] <= 0:
            continue
        a = por_acta.setdefault(x['acta'] or '—',
                                {'m3': 0.0, 'predio': x.get('destino', ''), 'prod': {}, 'est': {}})
        a['m3'] += x['m3']
        p = x['producto']
        if p in ('PODADO', 'ASERRABLE', 'PULPABLE'):
            a['prod'][p] = a['prod'].get(p, 0.0) + x['m3']
        e = (x.get('estado') or '').strip()
        if e:
            e = 'Quemada' if e.lower().startswith('quemada') else e
            a['est'][e] = a['est'].get(e, 0.0) + x['m3']
    if not por_acta:
        return "<div class=cob>El NOC no informa actas para esta faena en el período.</div>"

    def mix(prod, total):
        """Pod/Ase/Pul en tres números. En una columna de ~60 mm no cabe una tabla anidada."""
        if not total:
            return "—"
        return " · ".join(f"{prod.get(p, 0.0)/total*100:.0f}" for p in ('PODADO', 'ASERRABLE', 'PULPABLE'))

    def rodal(est, total):
        if not est or not total:
            return "<td>—</td>"
        nom, m3 = max(est.items(), key=lambda y: y[1])
        col = 'nf' if nom.lower().startswith('fresca') else 'tp'
        return f"<td class={col}>{m3/total*100:.0f}% {nom.lower()[:6]}</td>"

    filas = ""
    suma = 0.0
    tprod = {}
    for acta, a in sorted(por_acta.items(), key=lambda y: -y[1]['m3']):
        suma += a['m3']
        for k, v in a['prod'].items():
            tprod[k] = tprod.get(k, 0.0) + v
        filas += (f"<tr><td class=l>{acta}</td><td class=nf>{fmt(a['m3'])}</td>"
                  f"{rodal(a['est'], a['m3'])}<td>{mix(a['prod'], sum(a['prod'].values()))}</td></tr>")
    filas += (f"<tr class=tot><td class=l><b>TOTAL</b></td><td class=nf>{fmt(suma)}</td>"
              f"<td></td><td>{mix(tprod, sum(tprod.values()))}</td></tr>")
    return ("<table><tr><th class=l>Acta</th><th>m³</th><th>Rodal</th>"
            "<th title='Podado · Aserrable · Pulpable, en % del acta'>Pod·Ase·Pul</th></tr>"
            + filas + "</table>"
            "<div class=cob>Una fila por <b>acta</b>: es la unidad del rodal, y una faena "
            "trabaja varias en el mismo mes con condiciones distintas. El <b>plan</b> del acta "
            "(mix comprometido) no lo trae el NOC.</div>")


def tabla_stock(regs, hasta_iso):
    """Stock en Bosque por producto y destino.

    OJO: `STOCK` del NOC es un NIVEL de inventario, no un flujo — viene repetido en cada registro
    de esa combinación y **sumarlo da 7,5× la producción**. Se toma el ÚLTIMO valor por
    (producto, destino) y la antigüedad se cuenta desde esa fecha."""
    from datetime import date
    ult = {}
    for x in regs:
        if not x['producto'] or x['stock'] <= 0:
            continue
        try:
            d, m, a = x['fecha'].split('-')[:3]
            f = date(int(a), int(m), int(d))
        except Exception:
            continue
        k = (x['producto'], x['destino'])
        if k not in ult or f > ult[k][0]:
            ult[k] = (f, x['stock'])
    if not ult:
        return "<div class=cob>El NOC no informa stock para esta faena en el período.</div>"
    try:
        hoy = date.fromisoformat(hasta_iso[:10])
    except Exception:
        hoy = max(f for f, _ in ult.values())
    filas = ""
    for (prod, dest), (f, st) in sorted(ult.items(), key=lambda y: -y[1][1])[:6]:
        dias = (hoy - f).days
        col = "tp" if dias > 15 else "nf"          # stock viejo = riesgo de deterioro
        filas += (f"<tr><td class=l>{prod.title()}</td><td class=l>{dest[:22]}</td>"
                  f"<td class=nf>{fmt(st)}</td><td class={col}>{dias}</td></tr>")
    return ("<table><tr><th class=l>Producto</th><th class=l>Destino</th><th>Stock [m³]</th>"
            "<th>Antig. [días]</th></tr>" + filas + "</table>"
            "<div class=cob>Último nivel informado al NOC. Rojo sobre 15 días.</div>")


def jefe_de_turno(fa, dia_iso):
    """Jefe que está DE TURNO en esa faena ese día, según la matriz de turnos
    (`turnos_config.json`): rotación 7×7 correlativa, ciclo de 14 días desde `ref`,
    pos < 7 → jefeA, si no → jefeB. MISMA convención que `_turnos_schedule` de compute_kpis
    (si cambia una, cambiar la otra) — así el informe y la pestaña KPIs nombran al mismo jefe.

    La matriz manda sobre el roster del CMMS: es la que dice quién está el día del informe.
    Devuelve None si no hay config o la faena no está en ella (el campo queda en blanco)."""
    import json
    from datetime import date
    try:
        cfg = json.loads((BASE / "turnos_config.json").read_text(encoding='utf-8'))
        c = cfg.get('faenas', {}).get(NOMBRE.get(fa, fa))
        if not c:
            return None
        ref = date.fromisoformat(cfg.get('ref', '2026-07-01'))
        ciclo = int(cfg.get('ciclo_dias', 14))
        d = date.fromisoformat(dia_iso[:10])
        pos = (c['posR'] + (d - ref).days) % ciclo
        return c['jefeA'] if pos < ciclo // 2 else c['jefeB']
    except Exception as e:
        print(f"  ⚠️  matriz de turnos no legible ({e}); el campo Jefe de Faena queda en blanco")
        return None


_TURNOS_CACHE = {}


def turno_de(fa, dia_iso):
    """(nombre del jefe, 'A'|'B') para ese día, o (None, None). Misma matriz 7×7 que
    `jefe_de_turno`, pero devolviendo también CUÁL de los dos es, que es lo que permite pintar
    el bloque de turno en la tabla diaria. Cachea el JSON: si no, son 248 lecturas por corrida."""
    import json
    from datetime import date
    if 'cfg' not in _TURNOS_CACHE:
        try:
            _TURNOS_CACHE['cfg'] = json.loads((BASE / "turnos_config.json").read_text(encoding='utf-8'))
        except Exception:
            _TURNOS_CACHE['cfg'] = None
    cfg = _TURNOS_CACHE['cfg']
    if not cfg:
        return None, None
    try:
        c = cfg.get('faenas', {}).get(NOMBRE.get(fa, fa))
        if not c:
            return None, None
        ref = date.fromisoformat(cfg.get('ref', '2026-07-01'))
        ciclo = int(cfg.get('ciclo_dias', 14))
        pos = (c['posR'] + (date.fromisoformat(dia_iso[:10]) - ref).days) % ciclo
        return (c['jefeA'], 'A') if pos < ciclo // 2 else (c['jefeB'], 'B')
    except Exception:
        return None, None


def campo_jefes(fa, dia_iso):
    """Campo 'Jefe de Faena': UN nombre, el que está de turno ese día según la matriz.
    Sin matriz → línea en blanco para llenar a mano (el informe nunca deja de servir)."""
    j = jefe_de_turno(fa, dia_iso)
    return f"<b>{j}</b>" if j else "<b class=fill>____________</b>"


def cruce_clasificado(av, tp_faena, hp):
    """CRUCE de dos fuentes independientes sobre el mismo hecho: la GM detenida.

      · El OPERADOR declara en su pre-uso las horas que la GM estuvo parada (turno_perdida).
      · El JEFE declara en el CMMS los m³ que quedaron sin clasificar (avance_faena).

    Con el rendimiento del clasificado ya medido, una predice a la otra:
        horas de pana × rend ≈ m³ sin clasificar
    Si una dice pana y la otra no reporta madera pendiente (o al revés), es una
    incongruencia que hay que mirar: alguna de las dos capturas está fallando.

    Devuelve el HTML de la línea, o "" si no hay con qué cruzar (no se inventa nada).
    """
    if not av:
        return ""
    dia_av = int(av['fecha'][8:10])
    horas_pana = sum(t['horas'] for t in tp_faena
                     if t['proceso'] == 'CLASIFICADO' and t['dia'] == dia_av)
    declarado = av.get('sin_clasificar')
    if declarado is None:
        # El jefe todavía no informa el pendiente (campo nuevo). Solo se reclama cuando
        # SÍ hubo pana: si no la hubo, no hay nada que declarar y el aviso sería ruido.
        if horas_pana:
            return (f"<br><b>Cruce del clasificado:</b> el operador declaró "
                    f"<b>{horas_pana:g} h</b> de GM detenida, pero "
                    f"<b style='color:#B9770E'>el jefe no informó los m³ que quedaron sin "
                    f"clasificar</b> ese día.")
        return ""
    rend = (hp.get('CLASIFICADO') or {}).get('rend')

    if horas_pana and rend:
        esperado = horas_pana * rend
        # ±40%: es un estimado al ojo contra un cálculo de horas, no una medición.
        cuadra = declarado > 0 and abs(declarado - esperado) <= 0.4 * max(esperado, declarado)
        col = '#1E8449' if cuadra else '#B9770E'
        veredicto = 'cuadra' if cuadra else 'NO cuadra — revisar'
        return (f"<br><b>Cruce del clasificado:</b> el operador declaró "
                f"<b>{horas_pana:g} h</b> de GM detenida × {rend:.1f} m³/h ≈ "
                f"<b>{esperado:,.0f} m³</b> · el jefe informó <b>{declarado:,.0f} m³</b> "
                f"sin clasificar → <b style='color:{col}'>{veredicto}</b>.")
    if horas_pana and not rend:
        return (f"<br><b>Cruce del clasificado:</b> el operador declaró <b>{horas_pana:g} h</b> "
                f"de GM detenida y el jefe informó <b>{declarado:,.0f} m³</b> sin clasificar. "
                f"Falta el rendimiento del clasificado (pre-uso de la GM en días seguidos) "
                f"para contrastar las cifras.")
    if declarado > 0:
        return (f"<br><b>Cruce del clasificado:</b> el jefe informó <b>{declarado:,.0f} m³</b> "
                f"sin clasificar, pero <b style='color:#B9770E'>nadie declaró tiempo perdido "
                f"de la GM</b> ese día en el pre-uso.")
    return ("<br><b>Cruce del clasificado:</b> sin madera pendiente y sin tiempo perdido "
            "de la GM — las dos fuentes coinciden.")


def cobertura_preuso(hp, ult_dia):
    """Nota de procedencia bajo la tabla de Productividad: con cuántos turnos se midió el Uso/Rend
    real. Deja a la vista el gate de adopción — la celda solo se llena si hay pre-uso diario."""
    if not hp:
        return ("<div class=cob>Uso y Rend <b>real</b> quedan <i>por reportar</i>: esta faena no "
                "tiene pre-usos de días consecutivos este mes. Se llenan solos cuando el operador "
                "hace el <b>pre-uso diario</b> (un turno = Δ horómetro entre dos pre-usos seguidos).</div>")
    tot = sum(a['turnos'] for a in hp.values())
    dias = sorted({d for a in hp.values() for d in a['dias']})
    det = " · ".join(f"{p.title()} {a['turnos']}" for p, a in sorted(hp.items()))
    ultimo = f" · último día medido: {max(dias):02d}" if dias else ""
    return (f"<div class=cob>Uso y Rend <b>real</b>: horómetro del pre-uso — {tot} turno(s) en "
            f"{len(dias)} día(s) ({det}){ultimo}.</div>")



def nota_meta_procesos(metas_p, dias_con_flujo=0):
    """Nota bajo la tabla diaria: qué metas por proceso están cargadas y con cuántos días de
    producción declarada se está llenando volteo/madereo.

    Volteo y madereo dependen de dos cosas: la meta (Excel) y el FLUJO del día, que el jefe
    declara en /t/avance desde el 2026-07-31. Antes solo se guardaba el NIVEL del colchón, y un
    nivel no se descuenta de una meta (derivar el flujo restando dos niveles se probó con M7 y
    daba volteo NEGATIVO: son estimaciones al ojo)."""
    faltan = [p.title() for p in ('VOLTEO', 'MADEREO', 'CLASIFICADO')
              if not (metas_p or {}).get(p)]
    if faltan:
        return ("<div class=cob><b>Metas por proceso</b>: falta cargar "
                f"<b>{', '.join(faltan)}</b> en la hoja CONFIGURACIÓN del Excel maestro "
                "(columnas I, J y K). Mientras tanto esas columnas van en blanco — no se "
                "inventa una meta.</div>")
    if not dias_con_flujo:
        return ("<div class=cob><b>Volteo y madereo</b>: las metas están cargadas, pero el "
                "saldo no baja hasta que el jefe declare cuánto se <b>volteó y madereó cada "
                "día</b> en la app de terreno (Avance del día). El colchón que ya informa es un "
                "nivel, y un nivel no se descuenta de una meta.</div>")
    return (f"<div class=cob><b>Volteo y madereo</b>: saldo descontado con la producción "
            f"declarada por el jefe en {dias_con_flujo} día(s) del mes. Los días sin declarar "
            f"no descuentan — el saldo queda arriba de lo real hasta que se informen.</div>")


def shoveleo_mes(cmms, fid, mes_key):
    """Shoveleo del mes para una faena, desde `turno_shoveleo` (lo declara el operador de la
    shovel en su pre-uso, EN PROD 2026-07-31).

    El indicador que sirve NO son las horas sueltas: es **qué parte del turno de la shovel se
    va en acomodar madera** en vez de producir. El shoveleo es trabajo preparatorio — mueve y
    apila para que el skidder pueda cargar —, así que un porcentaje alto sostenido dice que el
    volteo está dejando la madera mal puesta, o que el terreno lo obliga.

    Base = horas REALES del turno (Δ horómetro), que es lo que se guardó junto a la
    declaración. Con la jornada nominal el porcentaje se pasaría del 100% la mitad de los días:
    la mediana de Δ horómetro de las shovels es 12,5 h contra 10,5 configuradas.

    Devuelve None si nadie declaró: la celda queda "rep.", nunca un 0 que se lea como medición.
    """
    filas = [x for x in ((cmms or {}).get('shoveleo', {}).get(fid) or [])
             if str(x.get('fecha', ''))[:7] == mes_key]
    if not filas:
        return None
    horas = sum(x['horas'] for x in filas)
    # Solo los días que traen la base pueden entrar al porcentaje; los otros suman horas y nada
    # más. Mezclarlos daría un porcentaje sobre un denominador incompleto.
    con_base = [x for x in filas if x.get('horas_turno')]
    base = sum(x['horas_turno'] for x in con_base)
    h_base = sum(x['horas'] for x in con_base)
    return {
        'dias': len(filas),
        'horas': horas,
        'h_dia': horas / len(filas),
        'pct': (h_base / base * 100) if base else None,
        'dias_pct': len(con_base),
        'equipos': sorted({x['equipo'] for x in filas if x.get('equipo')}),
    }


def nota_shoveleo(sh):
    """Nota de procedencia del shoveleo: con cuántos turnos se calculó. Deja el gate de
    adopción a la vista, igual que la cobertura del pre-uso."""
    if not sh:
        return ""
    eq = " · ".join(sh['equipos']) if sh['equipos'] else "la shovel"
    pct = (f" — <b>{sh['pct']:.0f}% del turno</b> ({sh['dias_pct']} día(s) con base de horas)"
           if sh['pct'] is not None else "")
    return (f"<div class=cob><b>Shoveleo</b>: {fmt(sh['horas'], 1)} h declaradas por {eq} en "
            f"{sh['dias']} día(s) del mes{pct}. Son horas <b>trabajadas</b>, no tiempo perdido: "
            f"el shoveleo acomoda la madera para que el madereo pueda cargar.</div>")


def nota_ref(ref, tec, esp):
    """Deja a la vista de dónde sale la columna Plan. Importa decirlo: hasta el 2026-07-30 el
    plan era el p75 de la propia faena, así que una faena podía figurar 'sobre el plan' por el
    solo hecho de repetir su promedio."""
    if not ref:
        return ("<div class=cob><b>Plan</b>: Arauco no publica referencia para "
                f"{TECN.get(tec, tec)} × {ESPN.get(esp, esp)}, así que la columna queda en "
                "'—'. <b>Habitual</b> es el p75 del historial de esta faena.</div>")
    ritmo, carga, rend = ref
    return (f"<div class=cob><b>Plan</b> = referencia de Arauco para {TECN.get(tec, tec)} × "
            f"{ESPN.get(esp, esp)}: <b>{ritmo:g} ciclo/hr × {carga:g} m³/ciclo = "
            f"{rend:.1f} m³/hr</b>. <b>Habitual</b> es el p75 del historial de esta faena — "
            "sirve para saber qué se logra hoy, no para fijar la meta.</div>")


def aviso_ciclos(pp):
    """Avisa si se dejaron días fuera del cálculo de Carga y Ritmo. Es visible a propósito:
    un día descartado es un día MAL DECLARADO al NOC, y el aviso es lo que empuja a corregirlo."""
    n = (pp or {}).get('dias_fuera', 0)
    if not n:
        return ""
    ok = pp.get('dias_ok', 0)
    return (f"<div class=cob><b>Carga y Ritmo</b> se calcularon con {ok} día(s): se dejaron "
            f"<b>{n}</b> fuera por tener ciclos fuera de rango físico (ciclos mal declarados al "
            f"NOC). El volumen y el rendimiento del mes NO se tocan: usan todos los días.</div>")


def tramo_de(vma):
    """En qué tramo de la Guía VMA cae un árbol de ese tamaño. None si no hay VMA."""
    if vma is None or vma != vma:
        return None
    for i in range(len(TR) - 1):
        if TR[i] < vma <= TR[i + 1]:
            return LB[i]
    return LB[-1] if vma > TR[-2] else LB[0]


def guia_tabla(tec, esp, cell, teo, vma=None):
    """Tabla Habitual/Meta/Teórico por tramo VMA, para la tecnología+especie de la faena.

    RESALTA el tramo en el que está la faena. Sin eso la guía es una tabla de 6 filas y el jefe
    no tiene cómo saber cuál es la suya — que era justo el problema: todo el informe se apoya en
    el VMA para elegir contra qué compararte, y el número no aparecía por ninguna parte."""
    mio = tramo_de(vma)
    filas = ""
    for tr in LB:
        c = cell.get((tec, esp, tr))
        t = teo.get((tec, esp, tr))
        hab = f"{c['p50']}" if c else "—"
        meta = f"{c['p75']}" if c else "—"
        teov = f"{t}" if t is not None else "—"
        if not c and t is None:
            continue
        marca = " style='background:#fff3cf;font-weight:700' title='Tramo de esta faena'" if tr == mio else ""
        flecha = " ←" if tr == mio else ""
        filas += (f"<tr{marca}><td class=l>{tr}{flecha}</td><td>{hab}</td><td class=nf>{meta}</td>"
                  f"<td class=gu>{teov}</td></tr>")
    if not filas:
        return "<div class=pr>Sin muestra suficiente de VMA para esta tecnología/especie.</div>"
    return ("<table><tr><th class=l>Tramo VMA [m³/árbol]</th><th>Habitual<br>[m³/hr]</th>"
            "<th>Meta<br>[m³/hr]</th><th>Teórico<br>[m³/hr]</th></tr>" + filas + "</table>")


def sheet(fa, g, cell, teo, meta_mes, cap, cmms=None, kpis=None, bn=None, metas_p=None):
    mes_key = g.dia.str[:7].max()
    anio, mes = int(mes_key[:4]), int(mes_key[5:7])
    jul = g[(g.faena == fa) & (g.dia.str[:7] == mes_key)].copy()
    if len(jul) == 0:
        return ""
    ult = jul.dia.max()
    ult_dia = int(ult[8:10])
    last = jul[jul.dia == ult].iloc[0]
    predio = last.predio
    especie_cod = last.especie
    especie = ESP.get(especie_cod, especie_cod)
    tec = jul.tec.mode().iat[0]

    kpi_fa = (kpis or {}).get('por_faena', {}).get(fa)
    pg = prod_general(jul, meta_mes, anio, mes, ult_dia, kpi_fa, (kpis or {}).get('dr'))
    pp = plan_productividad(fa, jul, cell)
    n_mes, ops = dias_operables(anio, mes)
    meta_dia = meta_mes / max(len(ops), 1)

    # trozado real por día del mes (índice = día del mes)
    real_por_dia = {int(k[8:10]): float(v) for k, v in jul.groupby('dia').m3.sum().items()}

    # ── Datos del CMMS de esta faena (jefe + preuso) ──
    fid = FAENA_ID.get(fa)
    av_dias = (cmms or {}).get('avance_dias', {}).get(fid, {})   # 'YYYY-MM-DD' -> {volteado, cancha}
    # SOLO el mes del informe. La RPC devuelve todo el histórico de turno_perdida y acá se
    # agrupa por DÍA DEL MES: sin este filtro, en agosto el "acumulado del mes" sumaría julio,
    # y el 15 de julio caería en la misma celda que el 15 de agosto. Hoy no se nota porque solo
    # hay datos de un mes — se rompería el 1 de agosto.
    tp_faena = [t for t in ((cmms or {}).get('tp', {}).get(fid, []))
                if str(t.get('fecha', ''))[:7] == mes_key]
    tp_dia_proc = {}                                             # (día, proceso) -> horas
    for t in tp_faena:
        tp_dia_proc[(t['dia'], t['proceso'])] = tp_dia_proc.get((t['dia'], t['proceso']), 0) + t['horas']

    hp = horas_preuso(fa, cmms, real_por_dia, mes_key)

    # Qué días quedaron CONFIRMADOS: el tiempo perdido de un día se declara en el pre-uso de la
    # mañana SIGUIENTE ("ayer tu equipo trabajó"), así que el día D lo confirma el pre-uso de
    # D+1. Se usa `informe_preuso_dias` y NO las horas de `informe_horas_faena`: las horas
    # necesitan dos pre-usos de días CONSECUTIVOS (Δ horómetro) y ahí se pierden los días en que
    # el operador sí declaró pero el par no se pudo formar. Caso real (M7, julio): la GM-07 hizo
    # pre-uso el 28 y el 30 — con el salto de un día el par se descarta, y el clasificado de M7
    # quedaba sin una sola marca en todo el mes pese a haber declarado dos veces.
    from datetime import date, timedelta
    _pd = (cmms or {}).get('preuso_dias', {}).get(fid, set())
    dias_con_preuso = set()
    for f_iso, proc in _pd:
        try:
            d_conf = date.fromisoformat(f_iso) - timedelta(days=1)   # el pre-uso de hoy confirma AYER
        except Exception:
            continue
        if d_conf.strftime('%Y-%m') == mes_key:
            dias_con_preuso.add((d_conf.day, proc))

    def tp_cell(dd, proc):
        """Celda de tiempo perdido. Tres estados, y la diferencia importa:

          · ROJO con las horas — se declaró pérdida.
          · VERDE con un ✓     — hubo pre-uso de ese proceso ese día y NO se declaró pérdida:
                                 es un turno limpio CONFIRMADO.
          · EN BLANCO          — nadie declaró nada. NO se pinta verde: dar por bueno un día
                                 que no revisó nadie es exactamente el error que el informe
                                 viene evitando en todas las demás celdas ("rep." en vez de 0).
        """
        h = tp_dia_proc.get((dd, proc))
        if h:
            return f"<td class=tp>{h:g}</td>"
        if (dd, proc) in dias_con_preuso:
            return ("<td class=sp title='El operador declaró este día en el pre-uso de la mañana "
                    "siguiente y no reportó tiempo perdido'>✓</td>")
        # "s/p" = SIN PRE-USO. La celda vacía se leía como "no pasó nada" cuando en realidad
        # significa "nadie miró", que es lo contrario. Va en gris tenue a propósito: son 962 de
        # las 992 celdas del mes, y con el peso visual del rojo la tabla quedaría ilegible.
        # Solo en días PASADOS con producción — el futuro ya va en blanco y ahí el blanco sí
        # quiere decir "todavía no".
        if dd <= ult_dia:
            return ("<td class=nd title='Sin pre-uso: el operador no lo hizo el día siguiente, "
                    "así que no se sabe si hubo tiempo perdido'>s/p</td>")
        return "<td class=bl></td>"

    # VMA del mes: m³ ÷ árboles ACUMULADOS, no el promedio de los VMA diarios (misma regla
    # que carga y ritmo: promedio de razones ≠ razón de totales).
    _arb = float(jul.arb.sum())
    vma_mes = (float(jul.m3.sum()) / _arb) if _arb > 0 else None

    # ── Información General ──
    ig = (f"<div class=ig>"
          f"<div><span>Fecha</span><b>{ult_dia:02d} / {mes:02d} / {anio}</b></div>"
          f"<div><span>Team</span><b>{fa}</b></div>"
          f"<div><span>Predio</span><b>{predio}</b></div>"
          f"<div><span>Jefe de Faena</span>{campo_jefes(fa, ult)}</div>"
          f"<div><span>Especie</span><b>{especie}</b></div>"
          f"<div><span>Tecnología</span><b>{TECN.get(tec, tec)}</b></div>"
          # El VMA va ACÁ, junto a especie y tecnología: los tres describen con qué te tocó
          # trabajar, no lo que produjiste. Y es el que elige contra qué tramo de la Guía te
          # compara todo el resto del informe — hasta hoy no aparecía en la hoja.
          # El ESTADO DEL RODAL no va acá: se abre por ACTA en Cumplimiento Acta. Una faena
          # trabaja 2-6 actas al mes con condiciones distintas (M1.1: 98% fresca en una,
          # 60% quemada en otra) y el promedio del mes escondía justamente eso.
          + (f"<div title='m³ ÷ árboles del mes · tramo {tramo_de(vma_mes)} de la Guía'>"
             f"<span>VMA</span><b>{vma_mes:.3f} m³/árbol</b></div>" if vma_mes else
             "<div><span>VMA</span><b class=fill>sin dato</b></div>") + "</div>")

    # ── Producción General ──
    pgen = (f"<div class=kpi>"
            f"<div><span>Meta mes</span><b>{fmt(pg['meta_mes'])}</b></div>"
            f"<div><span>Avance plan</span><b>{fmt(pg['avance_plan'])}</b></div>"
            f"<div><span>Avance real</span><b>{fmt(pg['avance_real'])}</b></div>"
            f"<div><span>Cumplimiento</span><b>{pg['cumpl']:.0f}%</b></div>"
            f"<div><span>Proyección mes</span><b>{fmt(pg['proy'])}</b></div>"
            f"<div><span>Meta día p/ llegar</span><b>{fmt(pg['meta_dia_req'])}</b></div>"
            f"<div><span>Real diario</span><b>{fmt(pg['real_diario'])}</b></div></div>")

    # ── Tiempo perdido ACUMULADO del mes por proceso (del preuso) ──
    acum_proc = {}   # proceso -> {horas, causas:{texto:horas}}
    for t in tp_faena:
        a = acum_proc.setdefault(t['proceso'], {'horas': 0.0, 'causas': {}})
        a['horas'] += t['horas']
        k = texto_tp(t)          # con "Otro" agrupa por la nota, no por la etiqueta vacía
        a['causas'][k] = a['causas'].get(k, 0) + t['horas']
    # La columna "Hizo perder" se RETIRÓ (gerencia 2026-07-26): CONTRADECÍA la fila de al lado.
    # Solo contaba las horas declaradas con la causa formal "sin frente", y en la práctica el
    # operador escribe el motivo en "Otro" — M7 mostraba "Procesado perdió 13,25 h por falta de
    # madereo" y a la vez "Madereo hizo perder 3,5". Un número que contradice a otro es peor que
    # no mostrarlo.
    # El dato sigue disponible (`proceso_causante` de informe_tp_faena): reponer la columna
    # cuando los operadores usen la causa del catálogo. Es más probable ahora que se llama
    # "Sin madera en cancha" / "Sin madera trozada" en vez del viejo "Sin frente / sin madera".
    if acum_proc:
        filas_ac = ""
        for proc in ['VOLTEO', 'MADEREO', 'PROCESADO', 'CLASIFICADO']:
            a = acum_proc.get(proc)
            if not a:
                continue
            causa_top = max(a['causas'], key=a['causas'].get)
            filas_ac += (f"<tr><td class=l>{proc.title()}</td><td class=tp>{a['horas']:g}</td>"
                         f"<td class=l>{causa_top}</td></tr>")
        tp_acum = ("<div class=tpaclab>Acumulado del mes por proceso (del preuso):</div>"
                   "<table class=tpac><tr><th class=l>Proceso</th>"
                   "<th>Perdió [hrs]</th><th class=l>Causa principal</th></tr>"
                   + filas_ac + "</table>")
    else:
        tp_acum = ""

    # ── Principales Tiempos Perdidos: pre-llenado del preuso (turno_perdida), resto en blanco ──
    tp_ord = sorted(tp_faena, key=lambda t: -t['horas'])[:6]     # las de mayor pérdida primero
    tp_rows = ""
    for i, t in enumerate(tp_ord, 1):
        desc = texto_tp(t)
        # Código del NOC de Arauco, para que el jefe lo transcriba sin traducir de memoria.
        # Vacío = sin equivalente confirmado con Arauco; nunca un código inventado.
        cod = t.get('codigo_noc')
        cod_td = f"<td class=gu>{cod}</td>" if cod else "<td class=bl></td>"
        tp_rows += (f"<tr><td>{i}</td><td class=l>{t['proceso'].title()}</td>{cod_td}"
                    f"<td class=l>{desc}</td><td class=tp>{t['horas']:g}</td>"
                    f"<td>{t['fecha'][8:10]}/{t['fecha'][5:7]}</td></tr>")
    tp = (tp_acum +
          "<table><tr><th>Nº</th><th class=l>Proceso</th><th title='Código de tiempo perdido "
          "del NOC de Arauco'>Cód.<br>NOC</th><th class=l>Descripción</th>"
          "<th>Tiempo [hrs]</th><th>Fecha</th></tr>" + tp_rows + "</table>")

    cumple = "Sí" if pg['cumple_plan'] else "No"
    cumpl_block = (
        f"<div class=q>¿Se cumple avance plan del día? <b>{cumple}</b> "
        f"(cumplimiento {pg['cumpl']:.0f}%)</div>"
        f"<div class=recu>Volumen a recuperar: <b>{fmt(pg['recuperar'])} m³/día</b> "
        f"(sobre {pg['dias_rest']} "
        f"{'día operable restante' if int(pg['dias_rest']) == 1 else 'días operables restantes'})"
        f" &nbsp;·&nbsp; "
        f"Volumen proyectado mes: <b>{fmt(pg['proy'])} m³</b>"
        f"</div>")

    # ── PRODUCCIÓN — tabla diaria por proceso ──
    procesos = ['VOLTEO', 'MADEREO', 'PROCESADO', 'CLASIFICADO']
    head1 = "<tr><th rowspan=2 class=l>Día</th>" + "".join(
        f"<th colspan=4 class='{'grp1' if i%2==0 else 'grp2'}'>{p}</th>"
        for i, p in enumerate(procesos)) + "</tr>"
    head2 = "<tr>" + "".join(
        f"<th class='{'grp1' if i%2==0 else 'grp2'}' title='Saldo por cumplir: parte en la meta "
        f"del mes y baja con lo producido'>Saldo<br>[m³]</th>"
        f"<th class='{'grp1' if i%2==0 else 'grp2'}'>Meta día<br>[m³]</th>"
        f"<th class='{'grp1' if i%2==0 else 'grp2'}'>Real<br>[m³]</th>"
        f"<th class='{'grp1' if i%2==0 else 'grp2'}'>T.P<br>[hrs]</th>"
        for i in range(4)) + "</tr>"
    # SALDO por cumplir, no plan acumulado (así lo pide Arauco): la columna **parte en la meta
    # del mes y baja hasta 0** a medida que se produce, en vez de subir de 0 a la meta. Es una
    # cuenta regresiva de lo que falta. Descuenta lo REALMENTE trozado (no el plan): si el mes
    # va atrasado el saldo no llega a 0, y eso es justamente la información que hay que ver.
    # MES COMPLETO (gerencia): el informe replica el tablero de Arauco que se llena en faena,
    # así que no puede perder días. Los totales van abajo.
    #
    # SALDO = el que se tiene al EMPEZAR el día, no al cerrarlo (así lo lleva Arauco en su
    # "meta acumulada"). Antes se mostraba el de cierre y la Meta día de al lado se calculaba
    # con el de apertura: en la fila del día 1 se leía "7.520 | 258", donde el 258 sale de
    # 8.000 — un número que no aparecía en ninguna parte de la fila. Ahora las dos columnas
    # hablan del mismo saldo.
    #
    # Cada proceso lleva SU meta y por lo tanto SU saldo (metas_p). El clasificado descuenta el
    # mismo m³ del NOC que el procesado (si la GM no está en pana clasifica todo lo trozado).
    # Volteo y madereo NO tienen saldo: el CMMS captura el NIVEL del colchón, no la producción
    # del día, y un nivel no se puede descontar de una meta. Ver nota_meta_procesos().
    meta_cla = (metas_p or {}).get('CLASIFICADO')
    # VOLTEO y MADEREO ya tienen saldo cuando el jefe declara el FLUJO del día (columnas
    # m3_volteado_dia / m3_madereado_dia, EN PROD desde el 2026-07-31). Con la meta cargada y
    # el flujo declarado, la cuenta regresiva es la misma que la del procesado.
    flujo = {'VOLTEO': 'vol_dia', 'MADEREO': 'mad_dia'}
    dias_con_flujo = sum(1 for k, v in av_dias.items()
                         if k[:7] == mes_key and v.get('vol_dia') is not None)
    saldos_p = {p: float(metas_p[p]) for p in flujo if (metas_p or {}).get(p)}
    filas = ""
    tot_real = tot_tp = 0.0
    saldo = float(meta_mes)
    saldo_cla = float(meta_cla) if meta_cla else None
    for d in range(1, n_mes+1):
        es_op = f"{mes:02d}-{d:02d}" not in FERIADOS_IRR   # se trabaja todos los días
        saldo_previo = saldo        # lo que faltaba al EMPEZAR el día → da la meta de ese día
        saldo_cla_previo = saldo_cla
        prev_p = dict(saldos_p)
        saldo = max(0.0, saldo - real_por_dia.get(d, 0.0))
        if saldo_cla is not None:
            saldo_cla = max(0.0, saldo_cla - real_por_dia.get(d, 0.0))
        _av_d = av_dias.get(f"{mes_key}-{d:02d}") or {}
        for p, k in flujo.items():
            if p in saldos_p and _av_d.get(k) is not None:
                saldos_p[p] = max(0.0, saldos_p[p] - float(_av_d[k]))
        # Días futuros: fila en blanco (la meta es día a día, no se proyecta hacia adelante).
        jnom, jlado = turno_de(fa, f"{mes_key}-{d:02d}")
        cl_dia = f"l j{jlado}" if jlado else "l"
        ttl = f" title='Turno de {jnom}'" if jnom else ""
        if d > ult_dia:
            # Día futuro: sin datos, pero CON su color de turno — la rotación se sabe de
            # antemano y así el jefe ve en la hoja impresa cuándo le toca volver.
            filas += (f"<tr><td class='{cl_dia}'{ttl}>{d:02d}</td>"
                      + "<td class=bl></td>" * 16 + "</tr>")
            continue
        tot_real += real_por_dia.get(d, 0.0)
        tot_tp += sum(v for (dd_, _p), v in tp_dia_proc.items() if dd_ == d)
        real = real_por_dia.get(d)
        # VOLTEO / MADEREO: Real = lo PRODUCIDO ese día (flujo declarado por el jefe), que es lo
        # que significa la columna "Volumen día" del tablero de Arauco. NO el colchón: ese es un
        # nivel y vive en el bloque de buffers, no acá.
        # Nivel del colchón, por si ese día no hay flujo. Ver el comentario de abajo.
        NIVEL = {'VOLTEO': 'volteado', 'MADEREO': 'cancha'}

        def celdas_proc(proc, clave):
            m = (metas_p or {}).get(proc)
            prod = _av_d.get(clave)
            if prod is not None:
                real = f"<td class=nf>{prod:,.0f}</td>"
            else:
                # Sin flujo declarado, se muestra el COLCHÓN si el jefe lo informó ese día,
                # en el estilo azul de "guía" y con asterisco: es otro número, no la
                # producción del día, y mezclarlos sin marcarlos es lo que no se puede hacer.
                #
                # Por qué existe esta rama: la columna de flujo nació el 31-07-2026 y TODO el
                # histórico llega en null. Al cambiar esta celda de nivel a flujo, las
                # declaraciones que el jefe ya había hecho —M7 declaró los días 24 y 26 al 29—
                # desaparecieron del PDF y quedaron como "rep.", o sea "nadie declaró". Borrar
                # de la vista un dato que alguien sí entregó es peor que mostrarlo etiquetado.
                niv = _av_d.get(NIVEL.get(proc, ''))
                if niv is not None:
                    real = (f"<td class=gu title='Colchón declarado por el jefe: {niv:,.0f} m³ "
                            f"esperando el proceso siguiente. NO es la producción del día — esa "
                            f"se declara en la app de terreno desde el 31-07.'>{niv:,.0f}*</td>")
                else:
                    real = "<td class=pr>rep.</td>"
            if not m:
                # Sin meta cargada no hay ni saldo ni meta día: no se inventa el denominador.
                return f"<td class=bl></td><td class=bl></td>{real}"
            sal = fmt(prev_p.get(proc, m)) if dias_con_flujo else "<span class=pr>—</span>"
            if not es_op:
                mdia = "—"
            elif dias_con_flujo and d == ult_dia:
                # Mismo criterio que procesado y clasificado en el día de hoy: días restantes
                # del kpis.json, para que las cuatro columnas cuenten la misma historia.
                mdia = fmt(max(0.0, saldos_p.get(proc, m)) / max(int(pg['dias_rest']), 1))
            elif not dias_con_flujo:
                # SIN NINGUNA declaración en el mes, la meta dinámica se dispara sola: el saldo
                # nunca baja, los días restantes sí, y para fin de mes pide un número absurdo
                # (M1.1 llegaba a 4.721 m³/día). Con cero datos lo honesto es el PLAN LINEAL, y
                # el saldo va en "—" en vez de repetir la meta entera 31 veces como si nada se
                # hubiera hecho. Al primer día declarado, las dos columnas pasan a ser dinámicas.
                mdia = fmt(m / max(len(ops), 1))
            else:
                # Misma meta día dinámica que el procesado: lo que falta, repartido en los días
                # que quedan desde hoy.
                dias_desde_d = len([x for x in ops if x >= d]) or 1
                mdia = fmt(max(0.0, prev_p.get(proc, m)) / dias_desde_d)
            return f"<td>{sal}</td><td>{mdia}</td>{real}"

        vol = f"{celdas_proc('VOLTEO', 'vol_dia')}{tp_cell(d,'VOLTEO')}"
        mad = f"{celdas_proc('MADEREO', 'mad_dia')}{tp_cell(d,'MADEREO')}"
        # PROCESADO: pre-llenado del NOC.
        # Meta día es DINÁMICA TODOS los días, no solo hoy: es lo que ese día había que trozar para
        # llegar a la meta, dado lo que se llevaba trozado hasta el día ANTERIOR, repartido en
        # los días que quedaban desde ahí. Si un día se produce poco, la meta de los siguientes
        # SUBE; si se produce de más, baja. El plan lineal fijo (meta ÷ días del mes) no servía:
        # no sabe cómo viene el mes.
        pm_ac = fmt(saldo_previo)
        if not es_op:
            pm_di = "—"
        elif d == ult_dia:
            # HOY muestra el MISMO número que "Meta día p/ llegar" de Producción General: lo que
            # exige de aquí en adelante, ya descontado lo trozado hoy. Es lo que el jefe necesita
            # al mirar el informe, y evita dos "metas del día" distintas en la misma hoja.
            pm_di = fmt(pg['meta_dia_req'])
        else:
            # Días pasados: lo que ESE día había que trozar, con lo que se llevaba hasta el
            # anterior repartido en los días que quedaban desde ahí.
            dias_desde_d = len([x for x in ops if x >= d]) or 1
            pm_di = fmt(max(0.0, saldo_previo) / dias_desde_d)
        rr = f"<td class=nf>{fmt(real)}</td>" if real is not None else "<td class=bl></td>"
        pro = f"<td>{pm_ac}</td><td>{pm_di}</td>{rr}{tp_cell(d,'PROCESADO')}"
        # CLASIFICADO: lo hace la GM (excavadora) sobre lo que trozó el PM. Si NO está en pana
        # clasifica TODA la madera del procesador → su Real Día es el mismo m³ del NOC.
        # Cuando la GM se empana el PM sigue trozando y queda madera sin clasificar: ese
        # pendiente lo informa el jefe en el CMMS (aún por capturar) y ahí habrá que restarlo.
        # Con meta propia cargada, el clasificado lleva SU saldo y SU meta día; si no, repite las
        # del procesado (es el comportamiento que había, y sigue siendo cierto: clasifica lo mismo
        # que se trozó — lo único que cambia es contra qué meta se mide).
        if saldo_cla_previo is not None:
            c_ac = fmt(saldo_cla_previo)
            if not es_op:
                c_di = "—"
            elif d == ult_dia:
                # MISMO criterio que el procesado en el día de hoy: lo que falta DESPUÉS de hoy,
                # repartido en los días restantes de `kpis.json`. Sin esto los dos procesos
                # mostraban metas distintas en columnas pegadas (30-07: procesado 2.065 contra
                # clasificado 1.112) — el procesado contaba los días desde HOY y el clasificado
                # desde la fila. Dos "meta del día" en la misma hoja es el error que este
                # informe viene corrigiendo desde el principio.
                c_di = fmt(max(0.0, saldo_cla) / max(int(pg['dias_rest']), 1))
            else:
                dias_desde_d = len([x for x in ops if x >= d]) or 1
                c_di = fmt(max(0.0, saldo_cla_previo) / dias_desde_d)
        else:
            c_ac, c_di = pm_ac, pm_di
        cla = f"<td>{c_ac}</td><td>{c_di}</td>{rr}{tp_cell(d,'CLASIFICADO')}"
        cl_hoy = " class=hoy" if d == ult_dia else ""
        # Bloque de turno en la celda del DÍA (no en la fila): la fila lleva celdas verdes,
        # rojas, azules y la barra amarilla de HOY, y un fondo encima las arruinaría todas.
        # Pintando solo el día queda una banda a la izquierda que muestra los bloques de 7 días
        # de un jefe y 7 del otro — y con eso se puede leer el mes por turno, no por faena.
        filas += f"<tr{cl_hoy}><td class='{cl_dia}'{ttl}>{d:02d}</td>{vol}{mad}{pro}{cla}</tr>"
    # Fila de TOTALES del mes al pie (lo que el tablero de Arauco cierra abajo).
    v4 = "<td class=bl></td>" * 4
    tot_row = (f"<tr class=tot><td class=l><b>TOTAL</b></td>{v4}{v4}"
               f"<td></td><td></td><td class=nf>{fmt(tot_real)}</td><td class=tp>{tot_tp:g}</td>"
               f"<td></td><td></td><td class=nf>{fmt(tot_real)}</td><td></td></tr>")
    # Leyenda de los dos turnos. Con los NOMBRES: un color sin nombre obliga a adivinar, y el
    # sentido de esto es poder decir "este bloque es de fulano".
    jefes = []
    for lado in ('A', 'B'):
        nom = next((turno_de(fa, f"{mes_key}-{dd:02d}")[0] for dd in range(1, n_mes + 1)
                    if turno_de(fa, f"{mes_key}-{dd:02d}")[1] == lado), None)
        if nom:
            jefes.append(f"<i style='background:{'#e3ecf5' if lado=='A' else '#f7efe2'};"
                         f"border-left:3px solid {'#4a7ba7' if lado=='A' else '#b9863e'}'></i>{nom}")
    ley_turnos = (f"<div class=leyj>Turnos 7×7 en la columna del día:{''.join(jefes)}</div>"
                  if len(jefes) == 2 else "")
    diaria = (f"<table class=diaria>{head1}{head2}{filas}{tot_row}</table>" + ley_turnos
              + nota_meta_procesos(metas_p, dias_con_flujo))

    # ── PRODUCTIVIDAD por proceso (Plan guía vs Real NOC + horómetro del pre-uso) ──
    # Uso Real y Rend Real salen del HORÓMETRO DEL PRE-USO (ver horas_preuso). Donde no hay
    # pre-uso con tramo de un día la celda sigue diciendo "rep." — no se reconstruye nada.
    def uso_cell(proc):
        """HORAS trabajadas por equipo-día, no el porcentaje: es lo que el jefe copia a la
        pizarra de Arauco (ahí el casillero se llena con horas, ej. 11 plan / 10,4 real).
        El semáforo mantiene el criterio del 90% de la jornada, pero el número que se ve
        son las horas."""
        a = hp.get(proc)
        if not a or a['uso'] is None or not a['eq_dia']:
            return "<td class=pr>rep.</td>"
        hdia = a['horas'] / a['eq_dia']          # horas por equipo-día (comparable con la jornada)
        u = a['uso']
        col = '#1E8449' if u >= USO*100 else ('#B9770E' if u >= 60 else '#943126')
        extra = ' · sobre la jornada: el equipo trabajó doble turno' if hdia > HDISP * 1.3 else ''
        return (f"<td class=nf style='color:{col}' title='{a['horas']:g} h en {a['turnos']} "
                f"turno(s) de pre-uso · {u:.0f}% de la jornada de {HDISP} h{extra}'>{hdia:.1f}</td>")

    def rend_cell(proc, fallback=None):
        a = hp.get(proc)
        if a and a['rend'] is not None:
            return (f"<td class=nf title='m³ del NOC ÷ horas de pre-uso · "
                    f"{a['rend_dias']} día(s) con el proceso completo'>{a['rend']:.1f}</td>")
        if fallback is not None:
            return f"<td class=nf>{fallback:.1f}</td>"
        return "<td class=pr>rep.</td>"

    # Formato de la PIZARRA de Arauco (gerencia 2026-07-26): un bloque por proceso con SOLO los
    # factores que le aplican, en vez de una matriz con guiones. Carga y Ritmo existen únicamente
    # en madereo — el NOC solo entrega ciclos del equipo que reporta el folio —, así que en la
    # pizarra esas filas no están en los otros procesos y acá tampoco.
    def bloque(titulo, filas):
        fs = "".join(f"<tr><td class=l>{lab}</td>{hab}{plan}{real}{cum}</tr>"
                     for lab, hab, plan, real, cum in filas)
        return (f"<div><table class=prod><tr><th class=proc colspan=5>{titulo}</th></tr>"
                f"<tr><th class=l>Factores</th>"
                f"<th title='Lo que esta faena logra habitualmente (p75 de su propio historial)'>"
                f"Habitual</th>"
                f"<th title='Referencia de Arauco para esta tecnología y especie'>Plan</th>"
                f"<th>Real</th><th>Cumpl.</th></tr>{fs}</table></div>")

    gu  = lambda v: f"<td class=gu>{v}</td>"
    vac = "<td>—</td>"
    nada = "<td></td>"
    hplan = gu(f"{HDISP:g}")

    def cumpl(real, plan):
        """Columna '% cumplimiento' del tablero de Arauco (Real ÷ Plan). Sin plan o sin real
        queda vacía: el porcentaje contra un número que no existe no informa nada."""
        try:
            if not plan or real is None or real != real:
                return nada
            p = real / plan * 100
        except Exception:
            return nada
        col = '#1E8449' if p >= 90 else ('#B9770E' if p >= 60 else '#943126')
        return f"<td style='color:{col};font-weight:600'>{p:.0f}%</td>"

    # Referencia de Arauco para ESTA faena. Solo aplica a MADEREO: el NOC únicamente entrega
    # ciclos del equipo que reporta el folio, así que carga y ritmo no están medidos en los
    # otros procesos — y en la pizarra de Arauco tampoco aparecen ahí.
    sh = shoveleo_mes(cmms, fid, mes_key)
    ref = ref_arauco(tec, especie_cod)
    r_ritmo, r_carga, r_rend = ref if ref else (None, None, None)
    ra = lambda v, d=2: gu(f"{v:.{d}f}") if v else vac

    prodv = (
        "<div class=two>"
        + bloque("VOLTEO", [
            ("Horas [hrs]", vac, hplan, uso_cell('VOLTEO'), nada),
            ("Rendimiento [m³/hr]", "<td class=pr>guía</td>", vac, "<td class=pr>rep.</td>", nada),
            # Shoveleo: va en VOLTEO porque es donde lo lleva Arauco en su planilla, y porque
            # la shovel trabaja para el volteo. Plan queda en "—": Arauco NO publica una
            # referencia de shoveleo (lo verifiqué en las 4 hojas de su libro), y poner una
            # inventada sería peor que no tenerla.
            ("Shoveleo [hrs/día]", vac, vac,
             (f"<td class=nf title='{sh['horas']:g} h en {sh['dias']} día(s) declarados'>"
              f"{sh['h_dia']:.1f}</td>") if sh else "<td class=pr>rep.</td>", nada),
            ("Shoveleo [% turno]", vac, vac,
             (f"<td class=nf>{sh['pct']:.0f}%</td>" if (sh and sh['pct'] is not None)
              else "<td class=pr>rep.</td>"), nada)])
        + bloque("MADEREO", [
            ("Horas [hrs]", vac, hplan, uso_cell('MADEREO'), nada),
            ("Rendimiento [m³/hr]", f"<td>{pp['plan_rend']}</td>", ra(r_rend, 1),
             f"<td class=nf>{pp['real_rend']:.1f}</td>", cumpl(pp['real_rend'], r_rend)),
            ("Carga [m³/ciclo]", f"<td>{pp['plan_carga']}</td>", ra(r_carga),
             f"<td class=nf>{pp['real_carga']:.2f}</td>", cumpl(pp['real_carga'], r_carga)),
            ("Ritmo [ciclo/hr]", f"<td>{pp['plan_ritmo']}</td>", ra(r_ritmo),
             f"<td class=nf>{pp['real_ritmo']:.2f}</td>", cumpl(pp['real_ritmo'], r_ritmo))])
        + "</div><div class=two>"
        + bloque("PROCESADO", [
            ("Horas [hrs]", vac, hplan, uso_cell('PROCESADO'), nada),
            ("Rendimiento [m³/hr]", "<td class=pr>guía</td>", vac, rend_cell('PROCESADO'), nada)])
        + bloque("CLASIFICADO", [
            ("Horas [hrs]", vac, hplan, uso_cell('CLASIFICADO'), nada),
            ("Rendimiento [m³/hr]", "<td class=pr>guía</td>", vac, rend_cell('CLASIFICADO'), nada)])
        + "</div>" + nota_ref(ref, tec, especie_cod) + nota_shoveleo(sh)
        + cobertura_preuso(hp, ult_dia) + aviso_ciclos(pp))

    # ── GUÍA DE PRODUCTIVIDAD integrada ──
    ritmo = cap
    if ritmo:
        objetivos = (f"Ritmo del procesador de esta faena: <b>{ritmo:.0f} m³/día</b>. "
                     f"Objetivos de buffer: volteo ≥ 3 días × ritmo = <b>{ritmo*3:,.0f} m³</b> · "
                     f"madereo ≥ 2 días × ritmo = <b>{ritmo*2:,.0f} m³</b>.")
        # Colchón real reportado por el jefe (avance_faena del CMMS), si está disponible.
        av = (cmms or {}).get('avance', {}).get(FAENA_ID.get(fa))
        if av:
            def colchon(m3, obj_dias):
                dias = m3 / ritmo if ritmo else 0
                col = '#1E8449' if dias >= obj_dias else ('#B9770E' if dias >= obj_dias * .6 else '#943126')
                return (f"<b style='color:{col}'>{m3:,.0f} m³ = {dias:.1f} días</b>"
                        f" (obj. ≥ {obj_dias})")
            # El pendiente de clasificado va al REVÉS: es deuda, no colchón (menos es
            # mejor). Mismos umbrales que el semáforo del CMMS: verde ≤ medio día de
            # producción acumulada, rojo pasado un día entero sin clasificar.
            def deuda(m3):
                dias = m3 / ritmo if ritmo else 0
                col = '#1E8449' if dias <= .5 else ('#B9770E' if dias <= 1 else '#943126')
                return (f"<b style='color:{col}'>{m3:,.0f} m³ = {dias:.1f} días</b>"
                        f" (obj. ≤ 0,5)")
            sc = av.get('sin_clasificar')
            extra = f" · <b>sin clasificar</b> {deuda(sc)}" if sc is not None else ""
            objetivos += (f"<br><b>Reportado por el jefe</b> ({av['fecha']}): "
                          f"volteado adelantado {colchon(av['volteado'], 3)} · "
                          f"en cancha {colchon(av['cancha'], 2)}{extra}.")
            objetivos += cruce_clasificado(av, tp_faena, hp)
    else:
        objetivos = "Ritmo del procesador: sin capacidad cargada (sin dato de trozado del mes)."
    guia = guia_tabla(tec, especie_cod, cell, teo, vma_mes)
    guia_block = f"<div class=guia>{objetivos}</div>"

    # ── Stock en Bosque + Cumplimiento Acta (ambos del BN del NOC, ver cargar_bn) ──
    # Control Calidad se retiró: era una tabla vacía para llenar a mano y el informe ya no se
    # completa en papel (gerencia 2026-07-25).
    bn_fa = (bn or {}).get(fa, [])
    # Tres columnas: guía VMA + stock + acta. A ancho completo la guía empujaba el informe a
    # una página extra que quedaba casi vacía.
    otros = (
        "<div class=two>"
        f"<div><h2>Guía VMA · {ESPN.get(especie_cod, especie_cod)}</h2>{guia}</div>"
        f"<div><h2>Stock en Bosque</h2>{tabla_stock(bn_fa, ult)}</div>"
        f"<div><h2>Cumplimiento Acta</h2>{tabla_acta(bn_fa)}</div>"
        "</div>")

    return f"""<div class="sheet faena" data-faena="{fa}">
<header>{'<img src="'+LOGO+'">' if LOGO else ''}<div>
<h1>Tablero de Gestión Diaria de Faena · {NOMBRE.get(fa, fa)}</h1>
<div class=sub>{MESES[mes]} {anio} · Predio {predio} · {especie} · NOC + CMMS · <b>informe de gestión diaria</b></div>
</div></header>
<h2>Información General</h2>{ig}
<h2>Producción General</h2>{pgen}
<h2>Productividad según el VMA del bosque</h2>{guia_block}{prodv}
<h2>Principales Tiempos Perdidos</h2>{tp}{cumpl_block}
<div class=foot>Hoja 1 de 2 · GESTIÓN. El registro del mes día por día va en la hoja siguiente.</div>
</div>
<div class="sheet faena registro" data-faena="{fa}">
<header>{'<img src="'+LOGO+'">' if LOGO else ''}<div>
<h1>Registro del mes · {NOMBRE.get(fa, fa)}</h1>
<div class=sub>{MESES[mes]} {anio} · Predio {predio} · {especie} · hoja 2 de 2</div>
</div></header>
<h2>Producción — tabla diaria por proceso</h2>{diaria}
<div class=foot>Verde = del NOC · <i>rep.</i> = lo declara el jefe en el CMMS · <b>✓</b> en T.P = hubo pre-uso y NO se declaró tiempo perdido (turno limpio); <b>s/p</b> = sin pre-uso, no se sabe · <b>*</b> = colchón declarado por el jefe, no producción del día · <b>fila amarilla = HOY</b>. <b>Saldo</b> = lo que falta para la meta. <b>Meta día de hoy</b> = lo que exige por día para llegar.</div>
{otros}
</div>"""


# Faena (código NOC) → faena_id del CMMS (nodos_activos.parentId / avance_faena.faena_id).
FAENA_ID = {'M1.1':'faena-m1-1','M1.2':'faena-m1-2','M1.3':'faena-m1-3','M1.4':'faena-m1-4',
            'M5':'faena-m5','M7':'faena-m7','M9':'faena-m9','M11':'faena-m11'}

def datos_cmms():
    """Trae del CMMS (Supabase) lo que el NOC no sabe y el terreno declara, vía RPC de solo
    lectura con la CLAVE PÚBLICA (anon) — no necesita service-role ni expone las tablas:
      · informe_avance_dias() → volteado/cancha por faena y DÍA (jefe) → columnas VOLTEO/MADEREO
        y el último día → BUFFERS (colchón).
      · informe_tp_faena()    → tiempos perdidos del preuso por faena/día/proceso/causa → T.P.
      · informe_horas_faena() → horas trabajadas (Δ horómetro entre pre-usos de días seguidos)
        por faena/día/proceso → Factor Uso real y Rend real de la tabla de Productividad.
    Requiere SUPABASE_URL + SUPABASE_KEY en env; sin ellas devuelve vacío y el informe muestra
    "por reportar". Best-effort: cualquier error → vacío, no rompe el pipeline."""
    import os, json, urllib.request, urllib.error
    # avance: último por faena (buffers). avance_dias: faena -> {día_int -> {volteado,cancha}}.
    # tp: faena -> [{fecha, dia, proceso, causa, detalle, horas}]
    # horas: faena -> {(día_int, proceso) -> {horas, equipos, dotacion}}
    # _auth: RPCs que rechazaron la credencial → main() aborta con EXIT_CMMS_AUTH.
    out = {'avance': {}, 'avance_dias': {}, 'tp': {}, 'horas': {}, 'preuso_dias': {},
           'shoveleo': {}, '_auth': []}
    url = os.environ.get('SUPABASE_URL'); key = os.environ.get('SUPABASE_KEY')
    if not url or not key:
        return out

    def rpc(nombre):
        req = urllib.request.Request(
            url.rstrip('/') + '/rest/v1/rpc/' + nombre, data=b'{}', method='POST',
            headers={'apikey': key, 'Authorization': 'Bearer ' + key,
                     'Content-Type': 'application/json'})
        try:
            return json.loads(urllib.request.urlopen(req, timeout=20).read())
        except urllib.error.HTTPError as e:
            # 401 = la clave no sirve · 403 = la clave sirve pero al rol le falta el GRANT.
            # Las dos son configuración, no red: se anotan para abortar al volver a main().
            if e.code in (401, 403):
                out['_auth'].append(nombre)
            raise

    try:
        for r in rpc('informe_avance_dias'):     # viene ordenado por faena, fecha desc
            fid = r['faena_id']; dia = str(r['fecha'])[:10]   # clave = fecha ISO, no día suelto
            # sin_clasificar: None = ese día no se informó (es campo nuevo), ≠ 0 declarado.
            sc = r.get('m3_sin_clasificar')
            # NIVEL del colchón (volteado/cancha) vs FLUJO del día (vol_dia/mad_dia). Son cosas
            # distintas: el nivel alimenta el semáforo de colchón, el flujo descuenta la meta
            # del proceso. None = no informado ≠ 0 declarado.
            vd, md = r.get('m3_volteado_dia'), r.get('m3_madereado_dia')
            v = {'volteado': float(r['m3_volteado']), 'cancha': float(r['m3_cancha']),
                 'sin_clasificar': None if sc is None else float(sc),
                 'vol_dia': None if vd is None else float(vd),
                 'mad_dia': None if md is None else float(md)}
            out['avance_dias'].setdefault(fid, {})[dia] = v
            if fid not in out['avance']:         # primer registro = el más reciente = buffer
                out['avance'][fid] = {**v, 'fecha': r['fecha']}
    except Exception as e:
        print(f"  ⚠️  CMMS avance no disponible ({e}); volteo/madereo quedan 'por reportar'")
    try:
        for r in rpc('informe_tp_faena'):
            out['tp'].setdefault(r['faena_id'], []).append(
                {'fecha': r['fecha'], 'dia': int(r['fecha'][8:10]), 'proceso': r['proceso'],
                 'causa': r['causa'], 'detalle': r.get('detalle'), 'horas': float(r['horas']),
                 'codigo_noc': r.get('codigo_noc'),
                 # proceso que CAUSÓ la detención (eslabón anterior); None si no aplica
                 'causante': r.get('proceso_causante')})
    except Exception as e:
        print(f"  ⚠️  CMMS tiempos perdidos no disponibles ({e}); T.P queda para llenar")
    try:
        for r in rpc('informe_horas_faena'):
            out['horas'].setdefault(r['faena_id'], {})[(str(r['fecha'])[:10], r['proceso'])] = {
                'horas': float(r['horas']), 'equipos': int(r['equipos']),
                'dotacion': int(r['dotacion'])}
    except Exception as e:
        print(f"  ⚠️  CMMS horas de preuso no disponibles ({e}); Uso/Rend real quedan 'por reportar'")
    # Días con pre-uso (≠ días con horas calculadas): marca el turno limpio en la columna T.P.
    # Ver el comentario en sheet() sobre por qué NO sirven las horas para esto.
    try:
        for r in rpc('informe_preuso_dias'):
            out['preuso_dias'].setdefault(r['faena_id'], set()).add(
                (str(r['fecha'])[:10], r['proceso']))
    except Exception as e:
        print(f"  ⚠️  CMMS días de preuso no disponibles ({e}); la columna T.P no marca turnos limpios")
    # Horas de shoveleo declaradas por la shovel (tabla turno_shoveleo, EN PROD 2026-07-31).
    # Son horas TRABAJADAS, no tiempo perdido: por eso vienen de su propia RPC y NO de
    # informe_tp_faena. Ver la migración 20260731c del CMMS.
    try:
        for r in rpc('informe_shoveleo_faena'):
            out['shoveleo'].setdefault(r['faena_id'], []).append(
                {'fecha': str(r['fecha'])[:10], 'equipo': r.get('equipo'),
                 'horas': float(r['horas']),
                 'horas_turno': None if r.get('horas_turno') is None else float(r['horas_turno'])})
    except Exception as e:
        print(f"  ⚠️  CMMS shoveleo no disponible ({e}); la fila de shoveleo queda 'por reportar'")
    return out

def datos_tm():
    """Top causas de tiempo perdido por faena, del NOC (lo genera compute_kpis en el paso 2.7).
    Sin archivo → {} y la hoja de análisis omite ese bloque."""
    import json
    p = BASE / "tm_por_faena.json"
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding='utf-8'))
    except Exception:
        return {}


def datos_kpis():
    """Lee kpis.json (lo genera compute_kpis en el paso 2.7, antes del informe) para que la
    proyección y el avance del informe CALCEN con la pestaña KPIs. Devuelve {faena→registro}
    más los conteos de días del mes. Sin el archivo → vacío (el informe recalcula solo)."""
    import json
    p = BASE / "kpis.json"
    if not p.exists():
        return {'por_faena': {}, 'dr': None}
    try:
        d = json.loads(p.read_text(encoding='utf-8'))
        return {'por_faena': {f['team']: f for f in d.get('faenas', [])},
                'dr': d.get('dias_restantes')}
    except Exception as e:
        print(f"  ⚠️  kpis.json no legible ({e}); la proyección se recalcula en el informe")
        return {'por_faena': {}, 'dr': None}


def main():
    df = cargar_pg()
    if df is None or len(df) == 0:
        print("❌ sin datos PG"); return 1
    g = base_diaria(df)
    d, cell = tabla_p75(g)
    teo = teorico(d)
    metas = metas_excel()
    metas_proc = metas_procesos()   # las 4 metas por proceso (E + I/J/K del Excel)
    mes_key = g.dia.str[:7].max()
    gm = g[g.dia.str[:7] == mes_key]
    cap = {fa: round(x.m3.quantile(.90)) for fa, x in gm.groupby('faena')}
    faenas = [f for f in FAENA_ORDER if f in set(gm.faena)]
    cmms = datos_cmms()
    # Se aborta ANTES de escribir el HTML: más vale dejar en el hosting el informe completo de
    # ayer que pisarlo con uno al que le falta la mitad. Ver EXIT_CMMS_AUTH.
    if cmms.get('_auth'):
        print("")
        print(f"❌ CMMS: credencial rechazada por {', '.join(cmms['_auth'])}")
        print("   El informe saldría SIN volteo, SIN madereo, SIN tiempos perdidos y SIN horas")
        print("   de pre-uso. NO se escribe el HTML y NO se sube: queda el del día anterior.")
        print("   Revisar el secret SUPABASE_KEY y el GRANT EXECUTE de esas RPC en Supabase.")
        return EXIT_CMMS_AUTH
    kpis = datos_kpis()
    bn = cargar_bn()      # acta + stock del NOC (reporte BN)
    tm = datos_tm()       # top causas de tiempo perdido del NOC

    logo_img = ('<img src="' + LOGO + '">') if LOGO else ''
    idx = (f"<div class=\"sheet cover\" style=\"page-break-after:auto\">"
           f"<header>{logo_img}<div>"
           f"<h1>Informe de Faena — {MESES[int(mes_key[5:7])]} {mes_key[:4]}</h1>"
           f"<div class=sub>Tablero de Gestión Diaria (formato Arauco) · mitad productividad · "
           f"{len(faenas)} faenas · una hoja A4 por faena</div></div></header>"
           f"<div class=foot>Guía impresa que el jefe de faena completa a mano en el predio. "
           f"Pre-llena lo que el NOC ya sabe; el resto es \"por reportar\". "
           f"SSO (IAP, madurez, riesgos, tarea crítica, mapa de riesgo) va fuera de este informe.</div></div>")

    # Cada faena aporta DOS hojas: la 1 es el tablero (NOC + CMMS) y la 2 el análisis de KPIs.
    # BENCHMARK por tecnología con el MISMO criterio de la hoja 1 (Carga y Ritmo ya sin los
    # días de ciclos mal declarados). No se usa el de kpis.json: venía del número sin filtrar.
    pps, tecs = {}, {}
    for fa in faenas:
        jul_fa = g[(g.faena == fa) & (g.dia.str[:7] == mes_key)]
        if not len(jul_fa):
            continue
        pps[fa] = plan_productividad(fa, jul_fa, cell)
        tecs[fa] = str(jul_fa.tec.mode().iat[0]) if len(jul_fa.tec.mode()) else ''
    bench = {}
    for fa in faenas:
        pares = [f for f in faenas if tecs.get(f) == tecs.get(fa) and f in pps]
        if len(pares) < 2:          # sin par de su tecnología no hay con quién compararse
            continue
        kfa = (kpis or {}).get('por_faena', {})
        usos = [kfa.get(f, {}).get('uso_pct') for f in pares if kfa.get(f, {}).get('uso_pct')]
        bench[fa] = {
            'Uso':   min(max(usos), 105) if usos else None,
            'Ritmo': max((pps[f]['real_ritmo'] for f in pares if pps[f]['real_ritmo'] == pps[f]['real_ritmo']), default=None),
            'Carga': max((pps[f]['real_carga'] for f in pares if pps[f]['real_carga'] == pps[f]['real_carga']), default=None),
        }

    hojas = {fa: sheet(fa, g, cell, teo, metas.get(fa, METAS_DEFAULT.get(fa, 0)), cap.get(fa),
                       cmms, kpis, bn, metas_proc.get(fa))
             for fa in faenas}
    sheets = "".join(hojas[fa] for fa in faenas)

    # Un HTML por ZONA (no por faena): son los 2 PDFs que se mandan por Telegram. Mismas zonas
    # que las tablas y las imágenes del pipeline (GENERAR_IMAGEN), así el chat queda con 2
    # archivos en vez de 8 y es manejable para reenviar. Cada zona trae sus 4 hojas A4, una por
    # faena, con page-break entre ellas. Sin la barra de selección: es para imprimir.
    for zona, fzona in ZONAS.items():
        cuales = [fa for fa in faenas if fa in fzona]
        if not cuales:
            continue
        doc = (f"<!doctype html><html lang=es><head><meta charset=utf-8>"
               f"<title>Informe de Faena {zona} {mes_key}</title>"
               f"<style>{CSS}{CSS_INFORME}</style></head><body>"
               + "".join(hojas[fa] for fa in cuales) + "</body></html>")
        (BASE / f"Informe_Zona_{zona}.html").write_text(doc, encoding="utf-8")

    opciones = "".join(f"<option value=\"{fa}\">{NOMBRE.get(fa, fa)}</option>" for fa in faenas)
    barra = (
        "<div class=barra>"
        "<label>Faena:&nbsp;"
        "<select id=selFaena onchange=filtrar()>"
        "<option value=''>Todas las faenas</option>" + opciones +
        "</select></label>"
        "<button onclick=\"window.print()\">🖨️ Descargar PDF</button>"
        "<span class=hint id=hint>Elige una faena y descarga solo ese informe.</span>"
        "</div>")
    js = ("<script>function filtrar(){"
          "var v=document.getElementById('selFaena').value;"
          "var h=document.getElementById('hint');"
          "document.querySelectorAll('.sheet').forEach(function(s){"
          "if(!v){s.classList.remove('oculto');return;}"
          "if(s.classList.contains('cover')){s.classList.add('oculto');return;}"
          "s.classList.toggle('oculto',s.getAttribute('data-faena')!==v);});"
          "h.textContent=v?('Mostrando 1 faena — \"Descargar PDF\" baja solo este informe.')"
          ":'Elige una faena y descarga solo ese informe.';}</script>")
    html = (f"<!doctype html><html lang=es><head><meta charset=utf-8>"
            f"<title>Informe de Faena {mes_key}</title>"
            f"<style>{CSS}{CSS_INFORME}</style></head><body>{barra}{idx}{sheets}{js}</body></html>")
    out = BASE / "Informe_Faena.html"
    out.write_text(html, encoding="utf-8")
    # COPIA CON EL MES EN EL NOMBRE = registro histórico. Cada corrida sobrescribe la del mes en
    # curso y, al cambiar de mes, la del mes cerrado queda congelada sola — sin lógica de cierre
    # ni tarea aparte. Mismo patrón que los snapshots del dashboard (Dashboard_Cosecha_AAAA-MM).
    (BASE / f"Informe_Faena_{mes_key}.html").write_text(html, encoding="utf-8")
    print(f"✅ Informe_Faena.html — {len(faenas)} faenas, mes {mes_key}, {len(html):,} bytes")
    print(f"   + registro histórico: Informe_Faena_{mes_key}.html")
    print(f"   + {len(ZONAS)} HTML por zona (Informe_Zona_<zona>.html) para los PDF de Telegram")
    print(f"   faenas: {faenas}")
    print(f"   capacidades (trozado p90): {cap}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
