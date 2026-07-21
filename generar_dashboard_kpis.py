#!/usr/bin/env python3
"""
generar_dashboard_kpis.py — Dashboard HTML de KPIs Uso/Ritmo/Carga/VMA por faena.
Data-driven: lee un JSON de KPIs (de compute) + JSON de tiempos perdidos por faena,
emite HTML autocontenido (sin dependencias externas). Sirve para el pipeline diario.

Uso: python3 generar_dashboard_kpis.py <kpis.json> <tm_por_faena.json> <salida.html>
"""
import json, sys
from pathlib import Path

def miles(n):
    try: return f"{float(n):,.0f}".replace(',', '.')
    except: return str(n)
def dec(n, d=1):
    try: return f"{float(n):,.{d}f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except: return str(n)

def uso_fill(u):
    if u > 100.5: return 'fill-anom'
    if u >= 90: return 'fill-good'   # exigencia del cliente: Uso >= 90%
    if u >= 75: return 'fill-warn'
    return 'fill-crit'
def chip_class(p):
    if p >= 95: return 'c-good'
    if p >= 70: return 'c-warn'
    return 'c-crit'

def trozado_min(causas):
    return sum(v for k, v in causas.items() if 'rozado' in k.lower())

def diagnostico_faena(f, causas, vma):
    """Diagnóstico por reglas: palanca, causa raíz, recomendación. Grounded en los datos."""
    pal = f.get('palanca_limitante', '—')
    b = f.get('benchmark_tipo', {})
    uso = f['uso_pct']; troz = trozado_min(causas)
    top = list(causas.items())[:2]
    top_txt = ', '.join(f"{k.strip()} ({miles(v)} min)" for k, v in top) if top else 'sin registros de tiempo perdido relevantes'
    # causa raíz
    if pal == 'Uso':
        causa = f"El tiempo disponible se pierde en detenciones: {top_txt}. La palanca es tiempo perdido, atacable con mantención/coordinación."
    else:
        estr = "árboles chicos" if vma < 0.3 else "producto/terreno"
        causa = (f"La brecha manda por {pal}, de raíz estructural ({estr}, VMA {dec(vma,3)}), no por tiempo perdido. "
                 f"En paralelo el Uso ({dec(uso)}%) arrastra fallas: {top_txt}.")
    # recomendación
    if uso < 80 and troz > 0:
        rec = f"Prioridad terreno: confiabilidad del cabezal de trozado ({miles(troz)} min de falla/falta) para recuperar Uso."
    elif pal == 'Ritmo':
        rec = ("Optimizar el ciclo: distancia de arrastre y enganche/desenganche (aéreo) o vías de saca y velocidad del procesador (terrestre)."
               if f['tipo'] == 'Aéreo' else "Estudio de ciclo del skidder/procesador: distancia de arrastre, trazado de vías, velocidad de trozado.")
    elif pal == 'Carga':
        rec = "Palanca de Planificación: asignar rodal de mayor diámetro y consolidar enganche por vuelta; la máquina ya trabaja a buen ritmo/uso."
    else:
        rec = f"Recuperar Uso atacando las detenciones principales: {top_txt}."
    return dict(palanca=pal, causa=causa, recomendacion=rec, benchmark=b)

import base64 as _b64
_lp = Path(__file__).parent / "millalemu-logo.png"
LOGO_DATAURI = ("data:image/png;base64," + _b64.b64encode(_lp.read_bytes()).decode()) if _lp.exists() else ""


def build(kpis, tmf):
    F = kpis['faenas']
    parcial = kpis.get('parcial', False)
    T = kpis.get('totales', {})
    # totales (compat con formato junio que no tiene 'totales')
    acum = T.get('acum_m3') or sum(f['vol_m3'] for f in F)
    meta = T.get('meta_m3') or sum(f['meta_m3'] for f in F)
    cumpl = T.get('cumpl_pct') or round(acum/meta*100, 1)
    proy = T.get('proy_cierre_m3'); proy_pct = T.get('proy_cumpl_pct')
    vma_g = T.get('vma_ponderado') or kpis.get('vma_global_m3_arbol', 0)
    arb_g = T.get('arboles', 0)
    n_cumplen = sum(1 for f in F if (f.get('proy_cumpl_pct') or f['cumpl_pct']) >= 95)
    # palanca dominante
    from collections import Counter
    pal = Counter(f.get('palanca_limitante', '?') for f in F)
    pal_dom, pal_n = pal.most_common(1)[0]

    prod_max = max(f['prod_m3_h_disp'] if 'prod_m3_h_disp' in f else f.get('prod_m3_h_turno', 1) for f in F) or 1
    def prodval(f): return f.get('prod_m3_h_disp', f.get('prod_m3_h_turno', 0))
    def vmaval(f): return f.get('vma_m3_arbol', f.get('vma_ponderado_m3_arbol', 0))

    # ---- filas por faena (se agrupan por tipo más abajo) ----
    rowmap = {}
    details = {}
    for f in F:
        u = f['uso_pct']; uw = min(u, 100)
        anom = u > 100.5
        usoval = f'<span class="usoval {"warnflag" if anom else ""}">{dec(u)}%</span>'
        vma = vmaval(f); vma_dim = 'dim' if vma < 0.3 else ''
        pv = prodval(f); pw = round(pv/prod_max*100)
        # chip = proyeccion si parcial, sino cumplimiento
        cp = f.get('proy_cumpl_pct') if parcial else f['cumpl_pct']
        chip = f'<span class="chip {chip_class(cp)}"><span class="dot"></span>{dec(cp)}%</span>'
        acum_col = f'<td class="num dim">{miles(f["vol_m3"])}</td>'
        # ---- detalle por faena (para el modal) ----
        causas = tmf.get(f['nombre'], {}).get('top_causas', {})
        tm_tot = tmf.get(f['nombre'], {}).get('total_min', 0)
        details[f['nombre']] = dict(
            nombre=f['nombre'], tipo=f['tipo'],
            uso=u, ritmo=f['ritmo_ciclos_h'], carga=f['carga_m3_ciclo'], vma=round(vma,3),
            prod=pv, prod_ef=f.get('prod_m3_h_efec',0),
            acum=f['vol_m3'], meta=f['meta_m3'], cumpl=f['cumpl_pct'],
            proy=f.get('proy_cierre_m3'), proy_pct=f.get('proy_cumpl_pct'), prom_dia=f.get('prom_diario_m3'),
            dias=f.get('dias'), dd=f.get('dias_detalle', []), hrs_ef=f['hrs_efectivas'], hrs_disp=f['hrs_disponible'],
            ciclos=f['ciclos'], arboles=f.get('arboles'), arb_ciclo=f.get('arboles_por_ciclo'),
            tm=dict(total=tm_tot, mant=f.get('tm_mant_min',0), oper=f.get('tm_oper_min',0), proc=f.get('tm_proc_min',0),
                    causas=[{'n':k.strip(),'min':v} for k,v in list(causas.items())[:6]]),
            diag=diagnostico_faena(f, causas, vma), parcial=parcial,
            reto=dict(palanca=f.get('palanca_limitante'), pot=f.get('benchmark_tipo', {}),
                      gaps=f.get('gaps_palanca', {}), proy_pot=f.get('proy_potencial_pct'),
                      opp=f.get('oportunidad_palanca_m3'), status=f.get('meta_status'),
                      tec=f.get('grupo_tec', ''), solo=(f.get('brecha_palanca_pct', 0) == 0),
                      predio=f.get('predio',''), especie=f.get('especie','')),
        )
        rowmap[f['nombre']] = f'''<tr class="clickrow" tabindex="0" role="button" data-faena="{f['nombre']}" aria-label="Ver detalle de {f['nombre']}">
  <td class="l"><span class="fa">{f['nombre']}</span><button class="rowpdf" title="Informe PDF de {f['nombre']}" onclick="event.stopPropagation();openM('{f['nombre']}')">📄 PDF</button><span class="chev">›</span></td>
  <td><div class="usocell"><div class="usobar"><i class="{uso_fill(u)}" style="width:{uw}%"></i><span class="mark" style="left:90%"></span></div>{usoval}</div></td>
  <td class="num">{dec(f['ritmo_ciclos_h'],2)}</td>
  <td class="num">{dec(f['carga_m3_ciclo'],2)}</td>
  <td class="num {vma_dim}">{dec(vma,3)}</td>
  <td><div class="prodcell"><div class="prodbar"><i style="width:{pw}%"></i></div><span>{dec(pv,1)}</span></div></td>
  {acum_col}
  <td>{chip}</td>
  <td class="l">{f'<span class="pal">{f.get("palanca_limitante","—")}</span><span class="opp">↑ +{miles(f.get("oportunidad_palanca_m3") or 0)} m³ a potencial</span>' if f.get('oportunidad_palanca_m3') else '<span class="pal" style="color:var(--ink-3);font-weight:500;font-size:11.5px">propio récord</span>'}</td>
</tr>'''

    # ---- VMA reframe ----
    lo = min(F, key=vmaval); hi = max(F, key=vmaval)
    fastest = max(F, key=lambda x: x['ritmo_ciclos_h'])
    vma_ul = (f'<li><b>{fastest["nombre"]}</b> corre al Ritmo más alto ({dec(fastest["ritmo_ciclos_h"],2)} c/h) '
              f'en parte <b>porque sus árboles son chicos</b> (VMA {dec(vmaval(fastest),3)}): con árboles pequeños necesita muchos ciclos para juntar volumen.</li>'
              f'<li><b>{lo["nombre"]}</b> parte con la peor mano del bosque — <b>VMA {dec(vmaval(lo),3)}, el más bajo</b> '
              f'(≈{dec(lo.get("arboles_por_ciclo") or 0,0)} árboles/ciclo). Su bajo rendimiento es en parte estructural, no solo fallas.</li>'
              f'<li><b>{hi["nombre"]}</b> (VMA {dec(vmaval(hi),3)}, el más grande) hace volumen con pocos ciclos.</li>')

    # ---- prioridades ----
    troz = sorted(((n, trozado_min(v['top_causas'])) for n, v in tmf.items()), key=lambda x: -x[1])
    troz_txt = ' · '.join(f'{n.replace("Millalemu ","M")} {miles(m)}' for n, m in troz if m > 0)
    aereo_ritmo = [f['nombre'] for f in F if f['tipo'] == 'Aéreo' and f.get('palanca_limitante') == 'Ritmo']
    carga_f = [f['nombre'] for f in F if f.get('palanca_limitante') == 'Carga']
    p2 = ', '.join(f.replace('Millalemu ','M') for f in aereo_ritmo) or 'faenas aéreas'
    p3 = ', '.join(f.replace('Millalemu ','M') for f in carga_f) or 'faenas de baja carga'

    # ---- alertas ----
    peor = min(F, key=lambda x: x['cumpl_pct'])
    peor_c = tmf.get(peor['nombre'], {}).get('top_causas', {})
    peor_causas = ', '.join(f'{k.strip()} ({miles(v)} min)' for k, v in list(peor_c.items())[:2])
    peor_cp = peor.get('proy_cumpl_pct') if parcial else peor['cumpl_pct']
    ocultos = [f for f in F if (f.get('proy_cumpl_pct') or f['cumpl_pct']) >= 95 and f['uso_pct'] < 75]
    oc_txt = '; '.join(f"{f['nombre']} (Uso {dec(f['uso_pct'])}%)" for f in ocultos) or '—'
    anoms = [f for f in F if f['uso_pct'] > 100.5]
    an_txt = ' y '.join(f['nombre'] for f in anoms) or '—'

    alert_ocultos = ''
    if ocultos:
        alert_ocultos = f'''<div class="alert warn"><div class="ic">🟠</div><div class="body">
      <h4>Buen ritmo de cierre que esconde turno botado — {oc_txt}</h4>
      <p>Proyectan cumplir pero con Uso bajo: buena parte del turno pagado está detenida. Cumplen por tamaño de árbol o por acumulado, no por disciplina de tiempo — hay margen escondido en el Uso.</p>
    </div></div>'''
    alert_anom = ''
    if anoms:
        alert_anom = f'''<div class="alert info"><div class="ic">🟣</div><div class="body">
      <h4>Dato a revisar (registro de horas) — {an_txt}</h4>
      <p>Uso &gt;100% significa horas efectivas por sobre la jornada disponible nominal. Antes de fijar metas o premiar, <b>auditar los folios</b> para descartar sobre-jornada o error de captura — no es un logro.</p>
    </div></div>'''

    # ---- header stats ----
    if parcial:
        head_stat_main = f'''<div class="stat lead">
      <p class="k">Proyección de cierre</p>
      <div class="v mono">{dec(proy_pct,0)}<small>%</small></div>
      <div class="meter"><span style="width:{min(proy_pct,100)}%"></span></div>
      <p class="note">proy. {miles(proy)} de {miles(meta)} m³ · al día {kpis.get('ultimo_dia','')}/{kpis.get('mes_num_str','')}</p>
    </div>'''
        stat2 = f'''<div class="stat"><p class="k">Acumulado al día {kpis.get('ultimo_dia','')}</p>
      <div class="v mono">{dec(cumpl,1)}<small>%</small></div>
      <p class="note">{miles(acum)} m³ · {kpis.get('dias_restantes','')} días hábiles restantes</p></div>'''
    else:
        head_stat_main = f'''<div class="stat lead">
      <p class="k">Cumplimiento del mes</p>
      <div class="v mono">{dec(cumpl,1)}<small>%</small></div>
      <div class="meter"><span style="width:{min(cumpl,100)}%"></span></div>
      <p class="note">{miles(acum)} de {miles(meta)} m³ · brecha {miles(acum-meta)} m³</p>
    </div>'''
        stat2 = f'''<div class="stat"><p class="k">Producción</p>
      <div class="v mono">{miles(acum)}<small> m³</small></div>
      <p class="note">{n_cumplen} de {len(F)} faenas cumplen</p></div>'''

    sub = f"{kpis['mes']} · {len(F)} faenas · " + ("mes en curso — proyección de cierre" if parcial else "mes cerrado")
    col_acum = "Acum. m³" if parcial else "Volumen m³"
    col_chip = "Proy. cierre" if parcial else "Cumpl."
    # ---- tablas agrupadas por tipo (Terrestre / Aéreo) ----
    thead = (f'<thead><tr><th class="l">Faena</th><th>Uso</th><th>Ritmo<br><span>c/h</span></th>'
             f'<th>Carga<br><span>m³/ciclo</span></th><th>VMA<br><span>m³/árbol</span></th>'
             f'<th>Product.<br><span>m³/h disp</span></th><th>{col_acum}</th><th>{col_chip}</th><th class="l">Desafío<br><span>vs tu potencial</span></th></tr></thead>')
    def tabla(subset, titulo, subt):
        body = ''.join(rowmap[f['nombre']] for f in sorted(subset, key=lambda x: x['cumpl_pct']))
        return (f'<div class="grpwrap"><h3 class="grp">{titulo}<span>{subt}</span></h3>'
                f'<div class="tablecard"><div class="scroll"><table class="mono">{thead}<tbody>{body}</tbody></table></div></div></div>')
    terr = [f for f in F if f['tipo'] == 'Terrestre']
    aereo = [f for f in F if f['tipo'] == 'Aéreo']
    tables_html = ''
    if terr: tables_html += tabla(terr, 'Terrestre', f'{len(terr)} faenas · skidder / madereo–trozado')
    if aereo: tables_html += tabla(aereo, 'Aéreo', f'{len(aereo)} faenas · torre / cable')
    detail_json = json.dumps(details, ensure_ascii=False)
    _des = T.get('desafio', {})
    _ng = _des.get('por_gestion', 0); _nm = _des.get('multipalanca', 0); _nr = _des.get('requiere_rodal', 0); _sp = _des.get('sin_par', 0)
    _sp_txt = f' · {_sp} sin par de su tecnología (se mide contra su propio récord)' if _sp else ''
    if _nr == 0 and _nm == 0 and _ng:
        reto_global = f'<b>Meta global alcanzable por gestión:</b> los {_ng} equipos con par pueden cumplir su meta cerrando su palanca-desafío — hacia los {miles(meta)} m³ del sistema. Cada uno se mide contra el mejor de su MISMA tecnología (skidder / torre){_sp_txt}.'
    else:
        reto_global = f'<b>Hacia la meta global de {miles(meta)} m³:</b> {_ng} cumplen por gestión · {_nm} requieren varias palancas · {_nr} requieren rodal/plan{_sp_txt}.'

    return HTML.format(
        detail_json=detail_json, modal_js=MODAL_JS, logo=LOGO_DATAURI, reto_global=reto_global,
        mes=kpis['mes'], sub=sub, head_main=head_stat_main, stat2=stat2,
        vma_g=dec(vma_g,3), arb_g=miles(arb_g),
        pal_dom=pal_dom, pal_n=pal_n, nfaenas=len(F),
        tables_html=tables_html,
        vma_ul=vma_ul, troz_txt=troz_txt, p2=p2, p3=p3,
        peor=peor['nombre'], peor_cp=dec(peor_cp,1), peor_vol=miles(peor['vol_m3']),
        peor_ritmo=dec(peor['ritmo_ciclos_h'],2), peor_vma=dec(vmaval(peor),3),
        peor_uso=dec(peor['uso_pct'],1), peor_causas=peor_causas,
        alert_ocultos=alert_ocultos, alert_anom=alert_anom,
        parcial_note=("Uso medido sobre tiempo disponible (10,5 h/día). Mes en curso: las cifras de cierre son proyección al ritmo actual. " if parcial else "Uso medido sobre tiempo disponible (10,5 h/día). "),
    )

# CSS + esqueleto (idéntico al diseño validado)
HTML = r'''<title>Cosecha Millalemu · Uso · Ritmo · Carga · VMA — {mes}</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<script>window.KPI_LOGO="{logo}";window.KPI_MES="{mes}";</script>
<style>
  :root{{--bg:#eef1f6;--surface:#ffffff;--surface-2:#f6f8fa;--surface-3:#e2e6ea;--ink:#1c2530;--ink-2:#55606c;--ink-3:#8a949f;--line:#e2e6ea;--line-2:#cdd4da;--accent:#417505;--accent-2:#2a4d07;--accent-soft:#41750514;--good:#2e9b3f;--good-bg:#2e9b3f16;--warn:#8a6200;--warn-bg:#e8a20020;--crit:#d8392b;--crit-bg:#d8392b16;--anom:#7a5bd0;--anom-bg:#7a5bd018;--bar-track:#e2e6ea;--shadow:0 1px 2px rgba(15,23,42,.04),0 3px 12px rgba(15,23,42,.06);}}
  @media (prefers-color-scheme:dark){{:root{{--bg:#0F172A;--surface:#1c2530;--surface-2:#25334A;--surface-3:#55606c;--ink:#e2e6ea;--ink-2:#cdd4da;--ink-3:#94A3B8;--line:#55606c;--line-2:#475569;--accent:#8bbf4f;--accent-2:#6aab2e;--accent-soft:#8bbf4f22;--good:#43b054;--good-bg:#2e9b3f22;--warn:#e8a200;--warn-bg:#e8a20020;--crit:#e46a5f;--crit-bg:#d8392b1c;--anom:#9a86d6;--anom-bg:#7a5bd022;--bar-track:#55606c;--shadow:0 1px 2px rgba(0,0,0,.3),0 6px 20px rgba(0,0,0,.3);}}}}
  :root[data-theme="light"]{{--bg:#eef1f6;--surface:#ffffff;--surface-2:#f6f8fa;--surface-3:#e2e6ea;--ink:#1c2530;--ink-2:#55606c;--ink-3:#8a949f;--line:#e2e6ea;--line-2:#cdd4da;--accent:#417505;--accent-2:#2a4d07;--accent-soft:#41750514;--good:#2e9b3f;--good-bg:#2e9b3f16;--warn:#8a6200;--warn-bg:#e8a20020;--crit:#d8392b;--crit-bg:#d8392b16;--anom:#7a5bd0;--anom-bg:#7a5bd018;--bar-track:#e2e6ea;--shadow:0 1px 2px rgba(15,23,42,.04),0 3px 12px rgba(15,23,42,.06);}}
  :root[data-theme="dark"]{{--bg:#0F172A;--surface:#1c2530;--surface-2:#25334A;--surface-3:#55606c;--ink:#e2e6ea;--ink-2:#cdd4da;--ink-3:#94A3B8;--line:#55606c;--line-2:#475569;--accent:#8bbf4f;--accent-2:#6aab2e;--accent-soft:#8bbf4f22;--good:#43b054;--good-bg:#2e9b3f22;--warn:#e8a200;--warn-bg:#e8a20020;--crit:#e46a5f;--crit-bg:#d8392b1c;--anom:#9a86d6;--anom-bg:#7a5bd022;--bar-track:#55606c;--shadow:0 1px 2px rgba(0,0,0,.3),0 6px 20px rgba(0,0,0,.3);}}
  *{{box-sizing:border-box}}html{{-webkit-text-size-adjust:100%}}
  body{{margin:0;background:var(--bg);color:var(--ink);font-family:'IBM Plex Sans',system-ui,-apple-system,'Segoe UI',sans-serif;font-size:15px;line-height:1.55;letter-spacing:-.002em;-webkit-font-smoothing:antialiased}}
  .mono{{font-variant-numeric:tabular-nums}}
  .wrap{{max-width:1120px;margin:0 auto;padding:0 24px 80px}}
  header{{padding:28px 0 4px}}
  .hdr{{background:linear-gradient(135deg,var(--accent),var(--accent-2));border-radius:16px;padding:24px 30px;color:#fff;box-shadow:var(--shadow)}}
  .heyebrow{{font-size:12px;letter-spacing:.12em;text-transform:uppercase;color:#fff;opacity:.82;font-weight:600;margin:0 0 8px}}
  h1{{font-size:clamp(22px,3.6vw,32px);line-height:1.1;margin:0 0 6px;font-weight:800;text-wrap:balance;letter-spacing:-.5px;color:#fff}}
  .hsub{{color:#fff;opacity:.9;font-size:14.5px;margin:0}}
  .hdr{{position:relative}}
  .hlogo{{position:absolute;top:22px;right:26px;width:46px;height:46px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.22)}}
  .mprint{{background:var(--accent);border:1px solid var(--accent);color:#fff;font-size:12.5px;font-weight:600;padding:6px 12px;border-radius:8px;cursor:pointer;display:inline-flex;align-items:center;gap:6px}}
  .mprint:hover{{filter:brightness(1.08)}}
  .print-head{{display:none}}
  @media print {{
    body>*{{display:none !important}}
    #mback{{display:block !important;position:static;background:#fff;padding:0;inset:auto;overflow:visible}}
    .modal{{box-shadow:none !important;max-width:100% !important;border:0 !important;border-radius:0 !important;animation:none !important}}
    .mclose,.mprint{{display:none !important}}
    .print-head{{display:flex !important;align-items:center;gap:12px;padding:2px 0 14px;border-bottom:2px solid var(--accent);margin:0 0 6px}}
    .print-head img{{width:46px;height:46px;border-radius:8px;flex:0 0 auto}}
    .print-head b{{font-size:16px;color:#1c2530}} .print-head span{{display:block;color:#55606c;font-size:11.5px;margin-top:2px}}
    @page{{margin:15mm}}
  }}
  .hero{{display:grid;grid-template-columns:1.4fr 1fr 1fr 1fr;gap:14px;margin:28px 0 0}}
  @media(max-width:820px){{.hero{{grid-template-columns:1fr 1fr}}}}
  .stat{{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:18px 18px 16px;box-shadow:var(--shadow)}}
  .stat .k{{font-size:11.5px;letter-spacing:.09em;text-transform:uppercase;color:var(--ink-3);font-weight:650;margin:0 0 8px}}
  .stat .v{{font-size:34px;font-weight:720;letter-spacing:-.02em;line-height:1}}
  .stat .v small{{font-size:16px;font-weight:600;color:var(--ink-2)}}
  .stat .note{{font-size:12.5px;color:var(--ink-3);margin:8px 0 0}}
  .stat.lead{{background:linear-gradient(160deg,var(--accent-soft),transparent 70%),var(--surface)}}
  .meter{{height:7px;background:var(--bar-track);border-radius:5px;margin:12px 0 0;overflow:hidden}}
  .meter>span{{display:block;height:100%;border-radius:5px;background:var(--accent)}}
  section{{margin:40px 0 0}}
  .sec-h{{display:flex;align-items:baseline;gap:12px;margin:0 0 4px}}
  h2{{font-size:19px;font-weight:700;letter-spacing:-.01em;margin:0}}
  .sec-sub{{color:var(--ink-3);font-size:13.5px;margin:2px 0 18px}}
  .identity{{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:20px 22px;box-shadow:var(--shadow);display:flex;flex-wrap:wrap;align-items:center;gap:14px 10px}}
  .formula{{font-size:15px;display:flex;flex-wrap:wrap;align-items:center;gap:8px}}
  .term{{padding:6px 11px;border-radius:9px;font-weight:650;font-size:13.5px;border:1px solid var(--line-2);background:var(--surface-2)}}
  .term b{{display:block;font-size:11px;font-weight:600;color:var(--ink-3);letter-spacing:.02em}}
  .op{{color:var(--ink-3);font-weight:700}}.eq{{color:var(--accent);font-weight:800}}
  .prod{{background:var(--accent-soft);border-color:var(--accent)}}
  .identity .aside{{color:var(--ink-2);font-size:13px;flex:1 1 240px;min-width:220px;border-left:2px solid var(--line-2);padding-left:14px}}
  .grpwrap{{margin:0 0 18px}}
  .grp{{display:flex;align-items:baseline;gap:10px;font-size:14px;font-weight:750;color:var(--accent);margin:0 0 9px;letter-spacing:-.01em}}
  .grp span{{font-size:12px;font-weight:500;color:var(--ink-3);letter-spacing:0}}
  .legend-solo{{border:1px solid var(--line);border-radius:12px;background:var(--surface)}}
  .tablecard{{background:var(--surface);border:1px solid var(--line);border-radius:14px;box-shadow:var(--shadow);overflow:hidden}}
  .scroll{{overflow-x:auto}}
  table{{border-collapse:collapse;width:100%;min-width:880px}}
  thead th{{text-align:right;font-size:11px;letter-spacing:.05em;text-transform:uppercase;color:var(--ink-3);font-weight:650;padding:14px 12px 11px;border-bottom:1px solid var(--line-2);white-space:nowrap;background:var(--surface-2)}}
  thead th.l{{text-align:left}}
  thead th span{{font-weight:500;text-transform:none;letter-spacing:0}}
  tbody td{{padding:13px 12px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}}
  tbody tr:last-child td{{border-bottom:0}}tbody tr:hover td{{background:var(--surface-2)}}
  td.l{{text-align:left}}
  .fa{{font-weight:680;font-size:14.5px}}
  .tp{{display:inline-block;font-size:10.5px;letter-spacing:.04em;text-transform:uppercase;color:var(--ink-3);border:1px solid var(--line-2);border-radius:20px;padding:2px 8px;margin-left:2px;font-weight:600}}
  .num{{font-size:14.5px;font-weight:600}}.num.dim{{color:var(--ink-2);font-weight:550}}
  .usocell{{display:flex;align-items:center;gap:10px;justify-content:flex-end}}
  .usobar{{width:78px;height:9px;border-radius:5px;background:var(--bar-track);position:relative;overflow:hidden;flex:0 0 auto}}
  .usobar>i{{position:absolute;left:0;top:0;bottom:0;border-radius:5px}}
  .usobar>.mark{{position:absolute;top:-2px;bottom:-2px;width:2px;background:var(--ink-3);opacity:.5}}
  .usoval{{width:52px;text-align:right;font-weight:650;font-size:14px}}
  .fill-good{{background:var(--good)}}.fill-warn{{background:var(--warn)}}.fill-crit{{background:var(--crit)}}
  .fill-anom{{background:repeating-linear-gradient(45deg,var(--anom),var(--anom) 4px,transparent 4px,transparent 7px);background-color:var(--anom-bg)}}
  .prodcell{{display:flex;align-items:center;gap:9px;justify-content:flex-end}}
  .prodbar{{width:64px;height:9px;border-radius:5px;background:var(--bar-track);overflow:hidden;flex:0 0 auto}}
  .prodbar>i{{display:block;height:100%;border-radius:5px;background:var(--accent)}}
  .chip{{display:inline-flex;align-items:center;gap:5px;font-size:12.5px;font-weight:650;padding:3px 9px;border-radius:20px;font-variant-numeric:tabular-nums}}
  .c-good{{background:var(--good-bg);color:var(--good)}}.c-warn{{background:var(--warn-bg);color:var(--warn)}}.c-crit{{background:var(--crit-bg);color:var(--crit)}}
  .dot{{width:6px;height:6px;border-radius:50%;background:currentColor;flex:0 0 auto}}
  .pal{{font-size:12.5px;font-weight:700;color:var(--ink)}}
  .opp{{display:block;font-size:11px;font-weight:600;color:var(--accent);margin-top:1px;white-space:nowrap}}
  .warnflag{{color:var(--anom);font-weight:700}}
  .legend{{display:flex;flex-wrap:wrap;gap:16px;padding:12px 16px;border-top:1px solid var(--line);font-size:12px;color:var(--ink-3);background:var(--surface-2)}}
  .legend span{{display:inline-flex;align-items:center;gap:6px}}
  .sw{{width:11px;height:11px;border-radius:3px;flex:0 0 auto}}
  .callout{{background:var(--surface);border:1px solid var(--line);border-left:4px solid var(--accent);border-radius:12px;padding:20px 22px;box-shadow:var(--shadow)}}
  .callout p{{margin:0 0 10px;color:var(--ink-2);font-size:14px}}.callout ul{{margin:6px 0 0;padding-left:18px;color:var(--ink-2);font-size:14px}}
  .callout li{{margin:5px 0}}.callout b{{color:var(--ink)}}
  .prios{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px}}
  @media(max-width:860px){{.prios{{grid-template-columns:1fr}}}}
  .prio{{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:18px;box-shadow:var(--shadow);display:flex;flex-direction:column}}
  .prio .rank{{display:flex;align-items:center;gap:9px;margin:0 0 10px}}
  .prio .rnum{{width:26px;height:26px;border-radius:8px;background:var(--accent);color:#fff;display:grid;place-items:center;font-weight:750;font-size:14px;flex:0 0 auto}}
  .prio .when{{font-size:11px;text-transform:uppercase;letter-spacing:.06em;font-weight:700}}
  .when.now{{color:var(--good)}}.when.mid{{color:var(--warn)}}.when.plan{{color:var(--ink-3)}}
  .prio h4{{margin:0 0 8px;font-size:15px;font-weight:700;line-height:1.25}}
  .prio p{{margin:0;color:var(--ink-2);font-size:13.5px}}
  .prio .lever{{margin-top:12px;font-size:12px;color:var(--ink-3)}}.prio .lever b{{color:var(--accent)}}
  .alerts{{display:flex;flex-direction:column;gap:12px}}
  .alert{{display:flex;gap:14px;background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:16px 18px;box-shadow:var(--shadow)}}
  .alert .ic{{font-size:18px;line-height:1;flex:0 0 auto;margin-top:1px}}.alert .body{{flex:1}}
  .alert h4{{margin:0 0 4px;font-size:14.5px;font-weight:700}}.alert p{{margin:0;color:var(--ink-2);font-size:13.5px}}
  .alert.crit{{border-left:4px solid var(--crit)}}.alert.warn{{border-left:4px solid var(--warn)}}.alert.info{{border-left:4px solid var(--anom)}}
  .clickrow{{cursor:pointer}}
  .clickrow:focus-visible{{outline:2px solid var(--accent);outline-offset:-2px}}
  .chev{{color:var(--ink-3);font-weight:700;margin-left:8px;opacity:0;transition:opacity .15s}}
  .clickrow:hover .chev,.clickrow:focus-visible .chev{{opacity:1;color:var(--accent)}}
  .rowpdf{{margin-left:10px;font-size:11px;font-weight:600;color:var(--accent);background:var(--accent-soft);border:1px solid var(--accent);border-radius:6px;padding:2px 8px;cursor:pointer;vertical-align:middle;transition:all .15s}}
  .rowpdf:hover{{background:var(--accent);color:#fff}}
  .mback{{position:fixed;inset:0;background:rgba(16,20,15,.55);display:none;align-items:flex-start;justify-content:center;padding:40px 20px;z-index:50;overflow-y:auto}}
  @supports (backdrop-filter:blur(2px)){{.mback{{backdrop-filter:blur(3px)}}}}
  .mback.open{{display:flex}}
  .modal{{background:var(--surface);border:1px solid var(--line-2);border-radius:18px;max-width:640px;width:100%;box-shadow:0 20px 60px rgba(0,0,0,.35);animation:pop .18s ease}}
  @keyframes pop{{from{{transform:translateY(8px);opacity:0}}to{{transform:none;opacity:1}}}}
  @media(prefers-reduced-motion:reduce){{.modal{{animation:none}}}}
  .mhead{{display:flex;align-items:flex-start;gap:12px;padding:22px 24px 16px;border-bottom:1px solid var(--line)}}
  .mhead h3{{margin:0;font-size:20px;font-weight:750;letter-spacing:-.01em}}
  .mhead .tp{{margin:6px 0 0}}
  .mclose{{margin-left:auto;background:var(--surface-2);border:1px solid var(--line);color:var(--ink-2);width:32px;height:32px;border-radius:9px;font-size:15px;cursor:pointer;line-height:1;flex:0 0 auto}}
  .mclose:hover{{background:var(--surface-3);color:var(--ink)}}
  .mbody{{padding:20px 24px 24px;display:flex;flex-direction:column;gap:20px}}
  .mrow{{display:flex;gap:12px;flex-wrap:wrap}}
  .mstat{{flex:1 1 110px;background:var(--surface-2);border:1px solid var(--line);border-radius:11px;padding:12px 14px}}
  .mstat .k{{font-size:10.5px;text-transform:uppercase;letter-spacing:.06em;color:var(--ink-3);font-weight:650;margin:0 0 5px}}
  .mstat .v{{font-size:20px;font-weight:720;letter-spacing:-.01em;font-variant-numeric:tabular-nums}}
  .mstat .v small{{font-size:11.5px;color:var(--ink-3);font-weight:600}}
  .msec-t{{font-size:12px;text-transform:uppercase;letter-spacing:.07em;color:var(--accent);font-weight:700;margin:0 0 10px}}
  .decomp{{display:flex;flex-direction:column;gap:10px}}
  .dterm{{display:grid;grid-template-columns:74px 1fr;align-items:center;gap:12px}}
  .dterm .lab{{font-size:12.5px;color:var(--ink-2);font-weight:600}}
  .dtrack{{height:8px;background:var(--bar-track);border-radius:5px;position:relative;overflow:visible}}
  .dtrack>i{{position:absolute;left:0;top:0;bottom:0;border-radius:5px}}
  .dtrack>.bm{{position:absolute;top:-3px;bottom:-3px;width:2px;background:var(--ink-3);opacity:.55;border-radius:2px}}
  .dval{{font-size:12.5px;font-weight:650;font-variant-numeric:tabular-nums;color:var(--ink);margin-top:5px;grid-column:2}}
  .dval b{{color:var(--ink-3);font-weight:500}}
  .causa-list{{display:flex;flex-direction:column;gap:8px}}
  .cz{{display:grid;grid-template-columns:1fr 70px 44px;align-items:center;gap:10px;font-size:12.5px}}
  .cz .cn{{color:var(--ink-2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
  .cz .cb{{height:7px;background:var(--bar-track);border-radius:4px;overflow:hidden}}
  .cz .cb>i{{display:block;height:100%;background:var(--crit);border-radius:4px}}
  .cz .cm{{text-align:right;font-variant-numeric:tabular-nums;color:var(--ink-2);font-weight:600}}
  .reto{{margin-top:12px;padding:11px 14px;border-radius:9px;font-size:13px;line-height:1.5;color:var(--ink-2)}}
  .reto b{{color:var(--ink)}}
  .reto-gestion{{background:var(--good-bg);border:1px solid var(--good)}}
  .reto-multipalanca{{background:var(--warn-bg);border:1px solid var(--warn)}}
  .reto-rodal{{background:var(--crit-bg);border:1px solid var(--crit)}}
  .retobanner{{background:var(--accent-soft);border:1px solid var(--line-2);border-left:4px solid var(--accent);border-radius:12px;padding:14px 18px;margin:20px 0 0;font-size:14px;color:var(--ink-2)}}
  .retobanner b{{color:var(--ink)}}
  .daytab{{width:100%;min-width:0;border-collapse:collapse;font-size:11.5px;font-variant-numeric:tabular-nums;table-layout:fixed}}
  .daytab th{{text-align:right;padding:5px 6px;color:var(--accent);font-size:9.5px;text-transform:uppercase;letter-spacing:.02em;font-weight:700;border-bottom:1.5px solid var(--accent);white-space:nowrap}}
  .daytab th:first-child{{text-align:left;width:15%}}
  .daytab td{{text-align:right;padding:3.5px 6px;border-bottom:1px solid var(--line);white-space:nowrap}}
  .daytab td:first-child{{text-align:left;font-weight:700;color:var(--accent)}}
  .daytab tbody tr:nth-child(even){{background:var(--surface-2)}}
  .daytab tbody tr.tot td{{border-top:2px solid var(--accent);border-bottom:0;font-weight:800;color:var(--accent);background:var(--surface);padding-top:6px}}
  .trend{{margin-top:2px}}
  .trend svg{{display:block;width:100%;height:122px;overflow:visible}}
  .diagbox{{background:var(--accent-soft);border:1px solid var(--line);border-left:3px solid var(--accent);border-radius:10px;padding:14px 16px}}
  .diagbox p{{margin:0 0 8px;font-size:13.5px;color:var(--ink-2)}}.diagbox p:last-child{{margin:0}}.diagbox b{{color:var(--ink)}}
  .pill-pal{{display:inline-block;font-size:11px;text-transform:uppercase;letter-spacing:.05em;font-weight:700;color:var(--accent);background:var(--surface);border:1px solid var(--accent);border-radius:20px;padding:2px 10px;margin-bottom:10px}}
  footer{{margin:44px 0 0;padding:22px 0 0;border-top:1px solid var(--line);color:var(--ink-3);font-size:12.5px}}
  footer p{{margin:0 0 6px}}footer b{{color:var(--ink-2)}}
</style>
<div class="wrap">
<header>
  <div class="hdr">
    <img src="{logo}" alt="Millalemu" class="hlogo">
    <p class="heyebrow">Forestal Millalemu · Control de Cosecha</p>
    <h1>Uso · Ritmo · Carga · VMA por faena</h1>
    <p class="hsub">{sub}</p>
  </div>
  <div class="retobanner">🎯 {reto_global}</div>
  <div class="hero">
    {head_main}
    {stat2}
    <div class="stat"><p class="k">VMA ponderado</p><div class="v mono">{vma_g}<small> m³/árb</small></div><p class="note">{arb_g} árboles madereados</p></div>
    <div class="stat"><p class="k">Palanca que más limita</p><div class="v" style="font-size:26px">{pal_dom}</div><p class="note">limita {pal_n} de {nfaenas} faenas · Uso bajo transversal</p></div>
  </div>
</header>
<section>
  <div class="sec-h"><h2>Cómo se lee</h2></div>
  <p class="sec-sub">Cada faena descompone su productividad en tres palancas multiplicativas. El VMA (tamaño de árbol) es el telón de fondo estructural.</p>
  <div class="identity">
    <div class="formula">
      <span class="term"><b>Utilización</b>Uso %</span><span class="op">×</span>
      <span class="term"><b>Velocidad</b>Ritmo (ciclos/h)</span><span class="op">×</span>
      <span class="term"><b>Volumen/ciclo</b>Carga (m³/ciclo)</span><span class="op eq">=</span>
      <span class="term prod"><b>Productividad</b>m³/h disponible</span>
    </div>
    <div class="aside"><b>Uso</b> = horas efectivas ÷ horas disponibles (turno 12 h − 1 h almuerzo − 0,5 h pausa activa = <b>10,5 h/día</b>). <b>VMA</b> = m³ ÷ árboles: es el rodal asignado, no lo controla el operador.</div>
  </div>
</section>
<section>
  <div class="sec-h"><h2>Las {nfaenas} faenas</h2></div>
  <p class="sec-sub">Separadas por sistema — Ritmo y Carga solo se comparan dentro del mismo tipo. Barra de Uso: marca gris = objetivo cliente 90%. <b style="color:var(--accent)">Clic en una faena (o en 📄 PDF) para ver su detalle y descargar/imprimir su informe.</b></p>
  {tables_html}
  <div class="legend legend-solo">
    <span><span class="sw" style="background:var(--good)"></span>Uso ≥ 90% (cumple cliente)</span>
    <span><span class="sw" style="background:var(--warn)"></span>Uso 75–90%</span>
    <span><span class="sw" style="background:var(--crit)"></span>Uso &lt; 70%</span>
    <span><span class="sw" style="background:repeating-linear-gradient(45deg,var(--anom),var(--anom) 3px,transparent 3px,transparent 5px)"></span>Uso &gt;100% — revisar registro</span>
  </div>
</section>
<section>
  <div class="sec-h"><h2>El VMA reordena las prioridades</h2></div>
  <p class="sec-sub">La corrección más importante del análisis.</p>
  <div class="callout">
    <p>El Ritmo aparece como “palanca limitante” en varias faenas, pero el <b>VMA demuestra que gran parte de esa brecha es estructural, no gestionable</b>: es el tamaño del árbol que le tocó a cada faena.</p>
    <ul>{vma_ul}</ul>
    <p style="margin-top:12px">La <b>Carga</b> la fija el rodal (VMA + largo de troza) → es palanca de <b>Planificación forestal</b>, no de terreno. La única palanca realmente accionable a corto plazo es el <b>Uso</b>.</p>
  </div>
</section>
<section>
  <div class="sec-h"><h2>Dónde está la plata accionable</h2></div>
  <p class="sec-sub">Sacando lo estructural, las tres acciones de mayor impacto en m³.</p>
  <div class="prios">
    <div class="prio"><div class="rank"><span class="rnum">1</span><span class="when now">Terreno · hoy</span></div>
      <h4>Confiabilidad del cabezal de trozado</h4>
      <p>Es la causa transversal del Uso bajo. Minutos de falla/falta de trozado que hoy inmovilizan el frente (min): {troz_txt}. Plan preventivo del cabezal (sierra, espada/cadena, rodillos, cuchillos, hidráulica) + repuestos críticos.</p>
      <div class="lever">Palanca: <b>Uso</b> — cierra la brecha en varias faenas a la vez</div></div>
    <div class="prio"><div class="rank"><span class="rnum">2</span><span class="when mid">Terreno · medio plazo</span></div>
      <h4>Optimizar el ciclo del sistema aéreo</h4>
      <p>Acortar distancia media de arrastre (reposicionar torre, más cambios de línea) y reducir tiempo muerto de enganche/desenganche. Enfoque en {p2}.</p>
      <div class="lever">Palanca: <b>Ritmo</b> — ganancia real es una fracción del techo teórico</div></div>
    <div class="prio"><div class="rank"><span class="rnum">3</span><span class="when plan">Planificación · lento</span></div>
      <h4>Asignar rodal de mayor diámetro</h4>
      <p>{p3} trabajan a ritmo/uso alto de su tipo pero mueven poco volumen por ciclo. La palanca la fija el bosque: destinar rodal más grueso cuando el plan lo permita.</p>
      <div class="lever">Palanca: <b>Carga / VMA</b> — mayor volumen potencial, decisión de Planificación</div></div>
  </div>
</section>
<section>
  <div class="sec-h"><h2>Alertas</h2></div>
  <div class="alerts">
    <div class="alert crit"><div class="ic">🔴</div><div class="body">
      <h4>{peor} — la más atrasada del mes ({peor_cp}% proyectado, {peor_vol} m³ acumulados)</h4>
      <p>Combina Ritmo lento ({peor_ritmo} c/h) y el VMA más bajo ({peor_vma}) con Uso deprimido ({peor_uso}%) por {peor_causas}. Cuando opera va lento <i>y</i> se detiene mucho: la faena que más gestión exige.</p>
    </div></div>
    {alert_ocultos}
    {alert_anom}
  </div>
</section>
<footer>
  <p><b>Fuentes.</b> KPIs Uso/Ritmo/Carga, VMA, volumen y tiempos perdidos: NOC · reporte Productividad Genérico (API), una sola fuente. Misma lógica de folio-dedup y clasificación de tiempos perdidos del pipeline oficial.</p>
  <p><b>Notas.</b> {parcial_note}Las “oportunidades” del análisis son techos teóricos (mueven una palanca al líder de su tipo con las otras dos fijas) y sirven para priorizar, no como metas.</p>
</footer>
</div>
<div class="mback" id="mback" role="dialog" aria-modal="true" aria-labelledby="mtitle"><div class="modal" id="modal"></div></div>
<script id="fdata" type="application/json">{detail_json}</script>
<script>{modal_js}</script>'''

MODAL_JS = r'''
const FD = JSON.parse(document.getElementById('fdata').textContent);
const back = document.getElementById('mback'), modal = document.getElementById('modal');
let lastFocus = null;
function nf(n,d){ return (n==null||n===''||isNaN(Number(n))) ? '—' : Number(n).toLocaleString('es-CL',{minimumFractionDigits:d,maximumFractionDigits:d}); }
function term(lab, val, unit, bm, decs, isLim){
  const pct = bm ? Math.min(val/bm*100,100) : 100;
  const col = isLim ? 'var(--crit)' : 'var(--accent)';
  const bmtxt = bm ? ` <b>· líder ${nf(bm,decs)}</b>` : '';
  return `<div class="dterm"><span class="lab">${lab}${isLim?' ◄':''}</span>`+
    `<div class="dtrack"><i style="width:${pct}%;background:${col}"></i><span class="bm" style="left:100%"></span></div>`+
    `<span class="dval">${nf(val,decs)} ${unit}${bmtxt}</span></div>`;
}
function render(name){
  const d = FD[name]; if(!d) return;
  const b = d.diag.benchmark||{}, lim = d.diag.palanca;
  const proyTxt = d.parcial ? `${nf(d.proy_pct,0)}%` : `${nf(d.cumpl,1)}%`;
  const causas = d.tm.causas||[]; const cmax = causas.length?causas[0].min:1;
  const causasHtml = causas.map(c=>`<div class="cz"><span class="cn">${c.n}</span><span class="cb"><i style="width:${Math.round(c.min/cmax*100)}%"></i></span><span class="cm">${nf(c.min,0)}</span></div>`).join('');
  const tmBlock = causas.length ? `<div><p class="msec-t">Tiempos perdidos — ${nf(d.tm.total,0)} min (mant ${nf(d.tm.mant,0)} · oper ${nf(d.tm.oper,0)} · proc ${nf(d.tm.proc,0)})</p><div class="causa-list">${causasHtml}</div></div>` : '';
  modal.innerHTML = `
    <div class="print-head">${window.KPI_LOGO?`<img src="${window.KPI_LOGO}" alt="Millalemu">`:''}<div><b>Informe de Faena — ${d.nombre}</b><span>Forestal Millalemu · Control de Cosecha · ${window.KPI_MES||''}</span></div></div>
    <div class="mhead"><div><h3 id="mtitle">${d.nombre}</h3><span class="tp">${d.tipo}</span></div>
      <button class="mprint" style="margin-left:auto" onclick="printFaena()" title="Imprimir o guardar como PDF">🖨 Imprimir / PDF</button>
      <button class="mclose" aria-label="Cerrar" onclick="closeM()">✕</button></div>`;
  return renderRest(d, name);
}
function renderRest(d, name){
  const b = d.diag.benchmark||{}, lim = d.diag.palanca;
  const proyTxt = d.parcial ? `${nf(d.proy_pct,0)}%` : `${nf(d.cumpl,1)}%`;
  const causas = d.tm.causas||[]; const cmax = causas.length?causas[0].min:1;
  const causasHtml = causas.map(c=>`<div class="cz"><span class="cn">${c.n}</span><span class="cb"><i style="width:${Math.round(c.min/cmax*100)}%"></i></span><span class="cm">${nf(c.min,0)}</span></div>`).join('');
  const tmBlock = causas.length ? `<div><p class="msec-t">Tiempos perdidos — ${nf(d.tm.total,0)} min (mant ${nf(d.tm.mant,0)} · oper ${nf(d.tm.oper,0)} · proc ${nf(d.tm.proc,0)})</p><div class="causa-list">${causasHtml}</div></div>` : '';
  modal.insertAdjacentHTML('beforeend', `
    <div class="mbody">
      <div class="mrow">
        <div class="mstat"><p class="k">Acumulado</p><div class="v">${nf(d.acum,0)}<small> / ${nf(d.meta,0)} m³</small></div></div>
        <div class="mstat"><p class="k">${d.parcial?'Proyección cierre':'Cumplimiento'}</p><div class="v">${proyTxt}</div></div>
        <div class="mstat"><p class="k">Productividad</p><div class="v">${nf(d.prod,1)}<small> m³/h</small></div></div>
        <div class="mstat"><p class="k">VMA</p><div class="v">${nf(d.vma,3)}<small> m³/árb</small></div></div>
      </div>
      <div><p class="msec-t">Tu desafío — actual vs tu potencial (mejor ${(d.reto&&d.reto.tec)||''} de tu tecnología)</p><div class="decomp">
        ${term('Uso', d.uso, '%', b.uso_pct, 1, lim==='Uso')}
        ${term('Ritmo', d.ritmo, 'c/h', b.ritmo_ciclos_h, 2, lim==='Ritmo')}
        ${term('Carga', d.carga, 'm³/c', b.carga_m3_ciclo, 2, lim==='Carga')}
      </div>
      <div class="reto reto-${(d.reto&&d.reto.solo)?'multipalanca':((d.reto&&d.reto.status)||'gestion')}">${(function(){
        const r=d.reto||{}, pal=r.palanca||'—';
        if(r.solo) return `<b>Sin par de tu tecnología (${r.tec||'—'}).</b> Te medís contra tu propio récord — se va poblando mes a mes. Por ahora tu potencial es tu mejor desempeño.`;
        const pp = r.proy_pot==null?'':(r.proy_pot>=130?'supera tu meta':nf(r.proy_pot,0)+'% de tu meta');
        if(r.status==='rodal') return `<b>Tu desafío: ${pal}.</b> Ni con el potencial de tu tecnología llegás a la meta → palanca de <b>Planificación</b> (más rodal/árbol) o revisar la meta.`;
        if(r.status==='multipalanca') return `<b>Tu desafío: ${pal}.</b> Cerrarla te acerca (+${nf(r.opp,0)} m³), pero para tu meta necesitás mejorar más de una palanca.`;
        return `<b>Tu desafío: ${pal}.</b> Cerrándola a tu potencial sumás <b>+${nf(r.opp,0)} m³</b> → tu proyección pasa a <b>${pp}</b>. Cumplís tu meta y aportás a la meta global.`;
      })()}</div></div>
      <div class="mrow">
        <div class="mstat"><p class="k">Hrs efectivas</p><div class="v">${nf(d.hrs_ef,0)}<small> h</small></div></div>
        <div class="mstat"><p class="k">Hrs disponibles</p><div class="v">${nf(d.hrs_disp,0)}<small> h</small></div></div>
        <div class="mstat"><p class="k">Ciclos</p><div class="v">${nf(d.ciclos,0)}</div></div>
        <div class="mstat"><p class="k">Árboles/ciclo</p><div class="v">${nf(d.arb_ciclo,1)}</div></div>
      </div>
      ${tmBlock}
      <div><p class="msec-t">Diagnóstico</p><div class="diagbox">
        <span class="pill-pal">Limita: ${d.diag.palanca}</span>
        <p>${d.diag.causa}</p><p><b>Acción:</b> ${d.diag.recomendacion}</p>
      </div></div>
      ${(d.dd&&d.dd.length)?(function(){
        const dd=d.dd, n=dd.length;
        const sv=dd.reduce((a,x)=>a+x.vol,0), sh=dd.reduce((a,x)=>a+x.hrs,0), sc=dd.reduce((a,x)=>a+x.cic,0);
        const mx=Math.max(...dd.map(x=>x.vol),1), avg=sv/n;
        const W=560,H=118,pad=16,bw=(W-pad*2)/n, ay=H-pad-(avg/mx)*(H-pad*2);
        const bars=dd.map((x,i)=>{const bh=(x.vol/mx)*(H-pad*2),bx=pad+i*bw;return `<rect x="${(bx+1).toFixed(1)}" y="${(H-pad-bh).toFixed(1)}" width="${Math.max(bw-2,1).toFixed(1)}" height="${bh.toFixed(1)}" rx="1.5" fill="var(--accent)" opacity="0.85"/>`+((i%3===0||i===n-1)?`<text x="${(bx+bw/2).toFixed(1)}" y="${H-3}" text-anchor="middle" font-size="8" fill="var(--ink-3)">${x.d}</text>`:'');}).join('');
        const rows=dd.map(x=>`<tr><td>${x.d}</td><td>${nf(x.vol,0)}</td><td>${nf(x.hrs,1)}</td><td>${nf(x.cic,0)}</td><td>${nf(x.rend,1)}</td></tr>`).join('');
        return `<div class="trend"><p class="msec-t">Tendencia del mes — m³ por día</p>
          <svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none">
            <line x1="${pad}" y1="${ay.toFixed(1)}" x2="${W-pad}" y2="${ay.toFixed(1)}" stroke="var(--ink-3)" stroke-dasharray="3 3" stroke-width="1"/>
            <text x="${W-pad}" y="${(ay-3).toFixed(1)}" text-anchor="end" font-size="8" fill="var(--ink-3)">prom ${nf(avg,0)} m³/día</text>
            ${bars}
          </svg></div>
          <div style="margin-top:16px"><p class="msec-t">Detalle diario — ${n} jornadas</p>
          <div class="scroll"><table class="daytab"><thead><tr><th>Día</th><th>m³</th><th>Hrs efec.</th><th>Ciclos</th><th>m³/hr</th></tr></thead><tbody>${rows}
          <tr class="tot"><td>Total</td><td>${nf(sv,0)}</td><td>${nf(sh,1)}</td><td>${nf(sc,0)}</td><td>${sh?nf(sv/sh,1):'—'}</td></tr>
          </tbody></table></div></div>`;
      })():''}
    </div>`);
}
function printFaena(){ window.print(); }
function openM(name){ lastFocus=document.activeElement; render(name); back.classList.add('open'); document.body.style.overflow='hidden'; const c=modal.querySelector('.mclose'); if(c)c.focus(); }
function closeM(){ back.classList.remove('open'); document.body.style.overflow=''; if(lastFocus)lastFocus.focus(); }
back.addEventListener('click', e=>{ if(e.target===back) closeM(); });
document.addEventListener('keydown', e=>{ if(e.key==='Escape' && back.classList.contains('open')) closeM(); });
document.querySelectorAll('.clickrow').forEach(r=>{
  const n = r.getAttribute('data-faena');
  r.addEventListener('click', ()=>openM(n));
  r.addEventListener('keydown', e=>{ if(e.key==='Enter'||e.key===' '){ e.preventDefault(); openM(n); }});
});
'''

if __name__ == '__main__':
    kp = Path(sys.argv[1]) if len(sys.argv) > 1 else Path('kpis_julio.json')
    tmp = Path(sys.argv[2]) if len(sys.argv) > 2 else Path('tm_por_faena_julio.json')
    out = Path(sys.argv[3]) if len(sys.argv) > 3 else Path('dashboard_kpis.html')
    kpis = json.loads(kp.read_text(encoding='utf-8'))
    # mes_num_str para el pie de fecha
    kpis['mes_num_str'] = f"{kpis.get('mes','').split()[0][:3]}"
    tmf = json.loads(tmp.read_text(encoding='utf-8'))
    out.write_text(build(kpis, tmf), encoding='utf-8')
    print('OK ->', out, f'({out.stat().st_size//1024} KB)')
