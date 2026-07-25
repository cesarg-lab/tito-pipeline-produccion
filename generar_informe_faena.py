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
    base_diaria, tabla_p75, teorico, metas_excel, cargar_pg,
    CSS, TEC_NORM, NOMBRE, ESP, ESPN, TECN, LB, TR, TRAMO_MID,
    USO, HDISP, METAS_DEFAULT, LOGO,
)

BASE = Path(__file__).parent
FAENA_ORDER = ['M1.1','M1.2','M1.3','M1.4','M5','M7','M9','M11']

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
tr.hoy td{background:#fff3cf!important;box-shadow:inset 0 0 0 1px #e0a800}
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
@media print{button.noprint{display:none!important}.barra{display:none!important}
  .diaria{font-size:7px}.sheet{padding:8px 10px}}
"""

MESES = ['','Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto',
         'Septiembre','Octubre','Noviembre','Diciembre']


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
    real_rend = jul.rend.median(); real_carga = jul.carga.median(); real_ritmo = jul.ritmo.median()
    return dict(plan_rend=plan_rend, plan_carga=plan_carga, plan_ritmo=plan_ritmo,
                real_rend=real_rend, real_carga=real_carga, real_ritmo=real_ritmo)


def horas_preuso(fa, cmms, real_por_dia):
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
    for (dia, proc), v in h.items():
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
    return (f"<div class=cob>Uso y Rend <b>real</b> medidos con el <b>horómetro del pre-uso</b>: "
            f"{tot} turno(s) en {len(dias)} día(s) del mes ({det}){ultimo}. Un turno = Δ horómetro "
            f"entre dos pre-usos de días seguidos, contra jornada de {HDISP} h. Sin pre-uso diario "
            f"la celda dice <i>rep.</i> — no se estima.</div>")


def guia_tabla(tec, esp, cell, teo):
    """Tabla Habitual/Meta/Teórico por tramo VMA, para la tecnología+especie de la faena."""
    filas = ""
    for tr in LB:
        c = cell.get((tec, esp, tr))
        t = teo.get((tec, esp, tr))
        hab = f"{c['p50']}" if c else "—"
        meta = f"{c['p75']}" if c else "—"
        teov = f"{t}" if t is not None else "—"
        if not c and t is None:
            continue
        filas += (f"<tr><td class=l>{tr}</td><td>{hab}</td><td class=nf>{meta}</td>"
                  f"<td class=gu>{teov}</td></tr>")
    if not filas:
        return "<div class=pr>Sin muestra suficiente de VMA para esta tecnología/especie.</div>"
    return ("<table><tr><th class=l>Tramo VMA (m³/árbol)</th><th>Habitual</th>"
            "<th>Meta</th><th>Teórico</th></tr>" + filas + "</table>")


def sheet(fa, g, cell, teo, meta_mes, cap, cmms=None, kpis=None):
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
    av_dias = (cmms or {}).get('avance_dias', {}).get(fid, {})   # día -> {volteado, cancha}
    tp_faena = (cmms or {}).get('tp', {}).get(fid, [])
    tp_dia_proc = {}                                             # (día, proceso) -> horas
    for t in tp_faena:
        tp_dia_proc[(t['dia'], t['proceso'])] = tp_dia_proc.get((t['dia'], t['proceso']), 0) + t['horas']

    def tp_cell(dd, proc):
        h = tp_dia_proc.get((dd, proc))
        return f"<td class=tp>{h:g}</td>" if h else "<td class=bl></td>"

    # ── Información General ──
    ig = (f"<div class=ig>"
          f"<div><span>Fecha</span><b>{ult_dia:02d} / {mes:02d} / {anio}</b></div>"
          f"<div><span>Team / Turno</span><b>{fa}</b> <span class=fill>/ ____</span></div>"
          f"<div><span>Predio</span><b>{predio}</b></div>"
          f"<div><span>Jefe de Faena</span>{campo_jefes(fa, ult)}</div>"
          f"<div><span>Especie</span><b>{especie}</b></div>"
          f"<div><span>Tecnología</span><b>{TECN.get(tec, tec)}</b></div></div>")

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
    acum_proc = {}   # proceso -> {horas, causas:{causa:horas}}
    for t in tp_faena:
        a = acum_proc.setdefault(t['proceso'], {'horas': 0.0, 'causas': {}})
        a['horas'] += t['horas']
        a['causas'][t['causa']] = a['causas'].get(t['causa'], 0) + t['horas']
    # DOS VISTAS del mismo hecho: lo que cada proceso SUFRIÓ (impacto: producción perdida) y
    # lo que CAUSÓ aguas abajo (acción: qué hay que arreglar). "Sin frente" significa que el
    # eslabón anterior no dejó qué hacer, así que la hora la sufre uno y la provoca otro:
    # anotarla solo contra el que se detuvo castiga al equipo que no tuvo la culpa.
    causado_proc = {}      # proceso -> horas que le hizo perder a otros
    for t in tp_faena:
        if t.get('causante'):
            causado_proc[t['causante']] = causado_proc.get(t['causante'], 0) + t['horas']
    if acum_proc or causado_proc:
        filas_ac = ""
        for proc in ['VOLTEO', 'MADEREO', 'PROCESADO', 'CLASIFICADO']:
            a = acum_proc.get(proc)
            causado = causado_proc.get(proc, 0)
            if not a and not causado:
                continue
            sufrido = f"{a['horas']:g}" if a else "—"
            causa_top = max(a['causas'], key=a['causas'].get) if a else "—"
            caus_td = (f"<td class=tp>{causado:g}</td>" if causado else "<td>—</td>")
            filas_ac += (f"<tr><td class=l>{proc.title()}</td><td class=tp>{sufrido}</td>"
                         f"<td class=l>{causa_top}</td>{caus_td}</tr>")
        tp_acum = ("<div class=tpaclab>Acumulado del mes por proceso (del preuso):</div>"
                   "<table class=tpac><tr><th class=l>Proceso</th>"
                   "<th title='Horas que este proceso perdió'>Perdió [hrs]</th>"
                   "<th class=l>Causa principal</th>"
                   "<th title='Horas que este proceso le hizo perder al siguiente de la cadena "
                   "por dejarlo sin frente'>Hizo perder [hrs]</th></tr>"
                   + filas_ac + "</table>"
                   "<div class=cob>«Hizo perder» = horas que ese proceso le costó al SIGUIENTE de "
                   "la cadena (volteo → madereo → procesado → clasificado) por dejarlo sin frente. "
                   "La misma hora se cuenta una vez en cada columna: una mide el impacto, la otra "
                   "de dónde viene. El volteo no tiene proceso anterior: sin frente ahí es rodal, "
                   "camino o permisos.</div>")
    else:
        tp_acum = ""

    # ── Principales Tiempos Perdidos: pre-llenado del preuso (turno_perdida), resto en blanco ──
    tp_ord = sorted(tp_faena, key=lambda t: -t['horas'])[:6]     # las de mayor pérdida primero
    tp_rows = ""
    for i, t in enumerate(tp_ord, 1):
        desc = t['causa'] + (f" — {t['detalle']}" if t.get('detalle') else "")
        # Código del NOC de Arauco, para que el jefe lo transcriba sin traducir de memoria.
        # Vacío = sin equivalente confirmado con Arauco; nunca un código inventado.
        cod = t.get('codigo_noc')
        cod_td = f"<td class=gu>{cod}</td>" if cod else "<td class=bl></td>"
        tp_rows += (f"<tr><td>{i}</td><td class=l>{t['proceso'].title()}</td>{cod_td}"
                    f"<td class=l>{desc}</td><td class=tp>{t['horas']:g}</td>"
                    f"<td class=bl></td><td class=bl></td><td>{t['fecha'][8:10]}/{t['fecha'][5:7]}</td></tr>")
    for _ in range(max(0, 4 - len(tp_ord))):                     # filas vacías para llenar a mano
        tp_rows += ("<tr><td class=bl>&nbsp;</td><td class=bl></td><td class=bl></td><td class=bl></td>"
                    "<td class=bl></td><td class=bl></td><td class=bl></td><td class=bl></td></tr>")
    tp = (tp_acum +
          "<table><tr><th>Nº</th><th class=l>Proceso</th><th title='Código de tiempo perdido "
          "del NOC de Arauco'>Cód.<br>NOC</th><th class=l>Descripción</th>"
          "<th>Tiempo [hrs]</th><th class=l>Acción</th><th class=l>Responsable</th>"
          "<th>Fecha</th></tr>" + tp_rows + "</table>")

    cumple = "Sí" if pg['cumple_plan'] else "No"
    cumpl_block = (
        f"<div class=q>¿Se cumple avance plan del día? <b>{cumple}</b> "
        f"(cumplimiento {pg['cumpl']:.0f}%) &nbsp;·&nbsp; ¿Por qué? <span class=fill>________________</span></div>"
        f"<div class=recu>Volumen a recuperar: <b>{fmt(pg['recuperar'])} m³/día</b> "
        f"(sobre {pg['dias_rest']} días operables restantes) &nbsp;·&nbsp; "
        f"Volumen proyectado mes: <b>{fmt(pg['proy'])} m³</b> &nbsp;·&nbsp; "
        f"Palanca de repunte: <span class=fill>________________</span></div>")

    # ── PRODUCCIÓN — tabla diaria por proceso ──
    procesos = ['VOLTEO', 'MADEREO', 'PROCESADO', 'CLASIFICADO']
    head1 = "<tr><th rowspan=2 class=l>Día</th>" + "".join(
        f"<th colspan=4 class='{'grp1' if i%2==0 else 'grp2'}'>{p}</th>"
        for i, p in enumerate(procesos)) + "</tr>"
    head2 = "<tr>" + "".join(
        f"<th class='{'grp1' if i%2==0 else 'grp2'}' title='Saldo por cumplir: parte en la meta "
        f"del mes y baja con lo producido'>Saldo</th>"
        f"<th class='{'grp1' if i%2==0 else 'grp2'}'>M.Día</th>"
        f"<th class='{'grp1' if i%2==0 else 'grp2'}'>Real</th>"
        f"<th class='{'grp1' if i%2==0 else 'grp2'}'>T.P</th>"
        for i in range(4)) + "</tr>"
    # SALDO por cumplir, no plan acumulado (así lo pide Arauco): la columna **parte en la meta
    # del mes y baja hasta 0** a medida que se produce, en vez de subir de 0 a la meta. Es una
    # cuenta regresiva de lo que falta. Descuenta lo REALMENTE trozado (no el plan): si el mes
    # va atrasado el saldo no llega a 0, y eso es justamente la información que hay que ver.
    filas = ""
    saldo = float(meta_mes)
    for d in range(1, n_mes+1):
        es_op = f"{mes:02d}-{d:02d}" not in FERIADOS_IRR   # se trabaja todos los días
        saldo = max(0.0, saldo - real_por_dia.get(d, 0.0))
        # Días FUTUROS (después de hoy): fila en blanco — el tablero se llena solo hasta hoy.
        if d > ult_dia:
            filas += f"<tr><td class=l>{d:02d}</td>" + "<td class=bl></td>" * 16 + "</tr>"
            continue
        real = real_por_dia.get(d)
        av = av_dias.get(d)
        # VOLTEO / MADEREO: Real = lo que el jefe declaró ese día (avance_faena); T.P del preuso.
        v_real = f"<td class=nf>{av['volteado']:,.0f}</td>" if av else "<td class=pr>rep.</td>"
        m_real = f"<td class=nf>{av['cancha']:,.0f}</td>" if av else "<td class=pr>rep.</td>"
        vol = f"<td class=bl></td><td class=bl></td>{v_real}{tp_cell(d,'VOLTEO')}"
        mad = f"<td class=bl></td><td class=bl></td>{m_real}{tp_cell(d,'MADEREO')}"
        # PROCESADO: pre-llenado NOC. M.Ac = plan lineal (dónde deberías ir). M.Día:
        # días pasados = plan del día; HOY = meta día DINÁMICA (lo que exige llegar a la meta
        # con lo ya procesado). Cambia día a día con el real.
        pm_ac = fmt(saldo)
        if not es_op:
            pm_di = "—"
        elif d == ult_dia:
            pm_di = fmt(pg['meta_dia_req'])            # dinámica: lo que exige de hoy a fin de mes
        else:
            pm_di = fmt(meta_dia)                      # plan del día ya transcurrido
        rr = f"<td class=nf>{fmt(real)}</td>" if real is not None else "<td class=bl></td>"
        pro = f"<td>{pm_ac}</td><td>{pm_di}</td>{rr}{tp_cell(d,'PROCESADO')}"
        # CLASIFICADO: lo hace la GM (excavadora) sobre lo que trozó el PM. Si NO está en pana
        # clasifica TODA la madera del procesador → su Real Día es el mismo m³ del NOC.
        # Cuando la GM se empana el PM sigue trozando y queda madera sin clasificar: ese
        # pendiente lo informa el jefe en el CMMS (aún por capturar) y ahí habrá que restarlo.
        cla = f"<td>{pm_ac}</td><td>{pm_di}</td>{rr}{tp_cell(d,'CLASIFICADO')}"
        cl_hoy = " class=hoy" if d == ult_dia else ""
        filas += f"<tr{cl_hoy}><td class=l>{d:02d}</td>{vol}{mad}{pro}{cla}</tr>"
    diaria = f"<table class=diaria>{head1}{head2}{filas}</table>"

    # ── PRODUCTIVIDAD por proceso (Plan guía vs Real NOC + horómetro del pre-uso) ──
    # Uso Real y Rend Real salen del HORÓMETRO DEL PRE-USO (ver horas_preuso). Donde no hay
    # pre-uso con tramo de un día la celda sigue diciendo "rep." — no se reconstruye nada.
    hp = horas_preuso(fa, cmms, real_por_dia)

    def uso_cell(proc):
        a = hp.get(proc)
        if not a or a['uso'] is None:
            return "<td class=pr>rep.</td>"
        u = a['uso']
        col = '#1E8449' if u >= USO*100 else ('#B9770E' if u >= 60 else '#943126')
        return (f"<td class=nf style='color:{col}' title='{a['horas']:g} h en {a['turnos']} "
                f"turno(s) de pre-uso · jornada {HDISP} h'>{u:.0f}</td>")

    def rend_cell(proc, fallback=None):
        a = hp.get(proc)
        if a and a['rend'] is not None:
            return (f"<td class=nf title='m³ del NOC ÷ horas de pre-uso · "
                    f"{a['rend_dias']} día(s) con el proceso completo'>{a['rend']:.1f}</td>")
        if fallback is not None:
            return f"<td class=nf>{fallback:.1f}</td>"
        return "<td class=pr>rep.</td>"

    prodv = (
        "<table><tr><th class=l>Proceso</th><th>Factor Uso [%]<br>Plan</th><th>Uso<br>Real</th>"
        "<th>Rend [m³/hr]<br>Plan</th><th>Rend<br>Real</th>"
        "<th>Carga [m³/ciclo]<br>Plan</th><th>Carga<br>Real</th>"
        "<th>Ritmo [ciclo/hr]<br>Plan</th><th>Ritmo<br>Real</th></tr>"
        f"<tr><td class=l>Volteo</td><td class=gu>{USO*100:.0f}</td>{uso_cell('VOLTEO')}"
        f"<td class=pr>guía</td><td class=pr>rep.</td><td>—</td><td>—</td><td>—</td><td>—</td></tr>"
        f"<tr><td class=l>Madereo</td><td class=gu>{USO*100:.0f}</td>{uso_cell('MADEREO')}"
        f"<td class=gu>{pp['plan_rend']}</td><td class=nf>{pp['real_rend']:.1f}</td>"
        f"<td class=gu>{pp['plan_carga']}</td><td class=nf>{pp['real_carga']:.2f}</td>"
        f"<td class=gu>{pp['plan_ritmo']}</td><td class=nf>{pp['real_ritmo']:.2f}</td></tr>"
        f"<tr><td class=l>Procesado</td><td class=gu>{USO*100:.0f}</td>{uso_cell('PROCESADO')}"
        f"<td class=pr>guía</td>{rend_cell('PROCESADO')}<td>—</td><td>—</td><td>—</td><td>—</td></tr>"
        f"<tr><td class=l>Clasificado</td><td class=gu>{USO*100:.0f}</td>{uso_cell('CLASIFICADO')}"
        f"<td class=pr>guía</td>{rend_cell('CLASIFICADO')}<td>—</td><td>—</td><td>—</td><td>—</td></tr>"
        "</table>" + cobertura_preuso(hp, ult_dia))

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
    guia = guia_tabla(tec, especie_cod, cell, teo)
    guia_block = (f"<div class=guia>{objetivos}</div>"
                  f"<div class=two><div><b style='font-size:10.5px;color:#1A5276'>"
                  f"Guía VMA×especie · {TECN.get(tec, tec)} · {ESPN.get(especie_cod, especie_cod)}</b>"
                  f"{guia}</div></div>")

    # ── Estados Línea Madereo (solo torre / M11) ──
    if tec == 'TORRE' or fa == 'M11':
        estados = ("<h2>Estados Línea Madereo (torre)</h2>"
                   "<table><tr><th class=l>Estado</th><th>Nº líneas</th><th>Largo [mts]</th>"
                   "<th>Volumen aprox [m³]</th><th>Avance [m³]</th></tr>"
                   "<tr><td class=l>En producción</td><td class=bl></td><td class=bl></td><td class=bl></td><td class=bl></td></tr>"
                   "<tr><td class=l>Líneas volteadas</td><td class=bl></td><td class=bl></td><td class=bl></td><td class=bl></td></tr>"
                   "<tr><td class=l>Líneas preinstaladas</td><td class=bl></td><td class=bl></td><td class=bl></td><td class=bl></td></tr>"
                   "</table>")
    else:
        estados = ""

    # ── Stock en Bosque / Cumplimiento Acta / Control Calidad (estructura en blanco) ──
    otros = (
        "<div class=two>"
        "<div><h2>Stock en Bosque</h2><table>"
        "<tr><th class=l>Producto</th><th class=l>Destino</th><th>Stock [m³]</th><th>Antig. [días]</th></tr>"
        "<tr><td class=bl>&nbsp;</td><td class=bl></td><td class=bl></td><td class=bl></td></tr>"
        "<tr><td class=bl>&nbsp;</td><td class=bl></td><td class=bl></td><td class=bl></td></tr></table></div>"
        "<div><h2>Cumplimiento Acta</h2><table>"
        "<tr><th class=l>Tipo</th><th>Plan [%]</th><th>Real Ac. [%]</th></tr>"
        "<tr><td class=l>Podado</td><td class=bl></td><td class=bl></td></tr>"
        "<tr><td class=l>Aserrable</td><td class=bl></td><td class=bl></td></tr>"
        "<tr><td class=l>Pulpable</td><td class=bl></td><td class=bl></td></tr></table></div>"
        "<div><h2>Control Calidad</h2><table>"
        "<tr><th class=l>Fecha</th><th>Pérdida [USD/m³]</th></tr>"
        "<tr><td class=bl>&nbsp;</td><td class=bl></td></tr>"
        "<tr><td class=bl>&nbsp;</td><td class=bl></td></tr></table></div>"
        "</div>")

    return f"""<div class="sheet faena" data-faena="{fa}">
<header>{'<img src="'+LOGO+'">' if LOGO else ''}<div>
<h1>Tablero de Gestión Diaria de Faena · {NOMBRE.get(fa, fa)}</h1>
<div class=sub>{MESES[mes]} {anio} · Predio {predio} · {especie} · pre-llenado con el NOC · <b>mitad productividad — guía para llenar en terreno</b></div>
</div></header>
<h2>Información General</h2>{ig}
<h2>Producción General</h2>{pgen}
<h2>Principales Tiempos Perdidos</h2>{tp}{cumpl_block}
<h2>Producción — tabla diaria por proceso</h2>{diaria}
<div class=foot>Verde = pre-llenado del NOC (trozado real y meta día del procesado). "rep." / celda amarilla = <b>por reportar</b> por el jefe (volteo y madereo en m³, tiempos perdidos, acta, stock). <b>Fila amarilla = HOY</b>. <b>Saldo</b> = lo que falta para la meta del mes: parte en la meta y baja con lo producido (llega a 0 solo si se cumple). <b>M.Día de hoy en adelante es dinámica</b>: lo que la meta exige por día con lo ya procesado (se recalcula cada día). Día = por hora de inicio del turno.</div>
<h2>Productividad — Plan (guía VMA+especie) vs Real (NOC)</h2>{prodv}
<h2>Guía de Productividad</h2>{guia_block}
{estados}
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
    import os, json, urllib.request
    # avance: último por faena (buffers). avance_dias: faena -> {día_int -> {volteado,cancha}}.
    # tp: faena -> [{fecha, dia, proceso, causa, detalle, horas}]
    # horas: faena -> {(día_int, proceso) -> {horas, equipos, dotacion}}
    out = {'avance': {}, 'avance_dias': {}, 'tp': {}, 'horas': {}}
    url = os.environ.get('SUPABASE_URL'); key = os.environ.get('SUPABASE_KEY')
    if not url or not key:
        return out

    def rpc(nombre):
        req = urllib.request.Request(
            url.rstrip('/') + '/rest/v1/rpc/' + nombre, data=b'{}', method='POST',
            headers={'apikey': key, 'Authorization': 'Bearer ' + key,
                     'Content-Type': 'application/json'})
        return json.loads(urllib.request.urlopen(req, timeout=20).read())

    try:
        for r in rpc('informe_avance_dias'):     # viene ordenado por faena, fecha desc
            fid = r['faena_id']; dia = int(r['fecha'][8:10])
            # sin_clasificar: None = ese día no se informó (es campo nuevo), ≠ 0 declarado.
            sc = r.get('m3_sin_clasificar')
            v = {'volteado': float(r['m3_volteado']), 'cancha': float(r['m3_cancha']),
                 'sin_clasificar': None if sc is None else float(sc)}
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
            out['horas'].setdefault(r['faena_id'], {})[(int(r['fecha'][8:10]), r['proceso'])] = {
                'horas': float(r['horas']), 'equipos': int(r['equipos']),
                'dotacion': int(r['dotacion'])}
    except Exception as e:
        print(f"  ⚠️  CMMS horas de preuso no disponibles ({e}); Uso/Rend real quedan 'por reportar'")
    return out

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
    mes_key = g.dia.str[:7].max()
    gm = g[g.dia.str[:7] == mes_key]
    cap = {fa: round(x.m3.quantile(.90)) for fa, x in gm.groupby('faena')}
    faenas = [f for f in FAENA_ORDER if f in set(gm.faena)]
    cmms = datos_cmms()
    kpis = datos_kpis()

    logo_img = ('<img src="' + LOGO + '">') if LOGO else ''
    idx = (f"<div class=\"sheet cover\" style=\"page-break-after:auto\">"
           f"<header>{logo_img}<div>"
           f"<h1>Informe de Faena — {MESES[int(mes_key[5:7])]} {mes_key[:4]}</h1>"
           f"<div class=sub>Tablero de Gestión Diaria (formato Arauco) · mitad productividad · "
           f"{len(faenas)} faenas · una hoja A4 por faena</div></div></header>"
           f"<div class=foot>Guía impresa que el jefe de faena completa a mano en el predio. "
           f"Pre-llena lo que el NOC ya sabe; el resto es \"por reportar\". "
           f"SSO (IAP, madurez, riesgos, tarea crítica, mapa de riesgo) va fuera de este informe.</div></div>")

    sheets = "".join(sheet(fa, g, cell, teo, metas.get(fa, METAS_DEFAULT.get(fa, 0)), cap.get(fa), cmms, kpis)
                     for fa in faenas)

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
    print(f"✅ Informe_Faena.html — {len(faenas)} faenas, mes {mes_key}, {len(html):,} bytes")
    print(f"   faenas: {faenas}")
    print(f"   capacidades (trozado p90): {cap}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
