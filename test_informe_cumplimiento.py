"""Pruebas de lo que se agregó el 2026-09-09: rend real de volteo + horas por día del pre-uso."""
import sys; sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent))
from generar_informe_faena import (rend_declarado, horas_preuso, nota_ref,
                                   plan_arboles, nota_shoveleo, nota_km, FAENA_ID)

MK = "2026-09"
ok = fail = 0
def eq(nom, got, esp):
    global ok, fail
    if got == esp or (isinstance(got, float) and isinstance(esp, float) and abs(got-esp) < 1e-6):
        ok += 1; print(f"  ✅ {nom}: {got}")
    else:
        fail += 1; print(f"  ❌ {nom}: obtuve {got!r}, esperaba {esp!r}")

# ── horas_preuso guarda las horas POR DÍA de los días medidos completos ──
print("horas_preuso · hfull_dia")
fa = 'M7'; fid = FAENA_ID[fa]
cmms = {'horas': {fid: {
    ('2026-09-01','VOLTEO'): {'horas': 8.0, 'equipos': 2, 'dotacion': 2},   # completo
    ('2026-09-02','VOLTEO'): {'horas': 5.0, 'equipos': 1, 'dotacion': 2},   # incompleto
    ('2026-09-03','VOLTEO'): {'horas': 9.0, 'equipos': 2, 'dotacion': 2},   # completo
    ('2026-08-31','VOLTEO'): {'horas': 99.,'equipos': 2, 'dotacion': 2},    # otro mes
}}}
hp = horas_preuso(fa, cmms, {}, MK)
eq("días completos", sorted(hp['VOLTEO']['dias_full']), [1, 3])
eq("horas por día", hp['VOLTEO']['hfull_dia'], {1: 8.0, 3: 9.0})
eq("no cruza de mes", hp['VOLTEO']['horas'], 22.0)

# ── rend_declarado ──
print("\nrend_declarado · volteo")
av = {'2026-09-01': {'vol_dia': 240.0},
      '2026-09-03': {'arb_vol_dia': 800},          # m³ nulo → se deriva con el VMA del NOC
      '2026-09-02': {'vol_dia': 999.0}}            # día sin pre-uso completo: NO debe entrar
vma = {3: 0.30}
r, h, nd = rend_declarado(hp, av, vma, MK, 'VOLTEO', 'vol_dia', 'arb_vol_dia')
eq("días usados", nd, 2)
eq("horas usadas (solo los 2 días declarados)", h, 17.0)
eq("rend = (240 + 800×0,30) ÷ 17", round(r, 3), round(480.0/17.0, 3))

print("\nrend_declarado · bordes")
eq("sin declaración → None", rend_declarado(hp, {}, vma, MK, 'VOLTEO','vol_dia','arb_vol_dia')[0], None)
eq("sin pre-uso → None", rend_declarado({}, av, vma, MK, 'VOLTEO','vol_dia','arb_vol_dia')[0], None)
eq("conteo sin VMA no inventa", rend_declarado(hp, {'2026-09-03': {'arb_vol_dia': 800}}, {},
                                               MK, 'VOLTEO','vol_dia','arb_vol_dia')[2], 0)
# un día declarado que NO tiene pre-uso completo no debe aportar horas ni m³
solo2 = rend_declarado(hp, {'2026-09-02': {'vol_dia': 999.0}}, vma, MK,'VOLTEO','vol_dia','arb_vol_dia')
eq("día sin pre-uso completo se ignora", solo2, (None, 0.0, 0))

# ── nota_plan_meta nombra los procesos sin meta ──
print("\nnota_ref · explica los DOS planes en un solo recuadro")
n = nota_ref((8, 5.2, 41.6), 'SKIDDER 6X6 GRAPPLE', 'PIRA',
             {'VOLTEO': 8855, 'PROCESADO': 8060, 'CLASIFICADO': None}, 28)
eq("un solo recuadro", n.count('class=cob'), 1)
eq("dice el plan de madereo", '41.6 m³/hr' in n, True)
eq("dice el plan de los otros tres", '28 días operables' in n, True)
eq("delata la meta faltante", 'Clasificado' in n.split('CONFIGURACIÓN')[-1], True)
eq("no delata las que sí están", 'Volteo' not in n.split('CONFIGURACIÓN')[-1], True)
sin = nota_ref(None, 'TORRE', 'EUNI', {}, 28)
eq("sin referencia de Arauco no revienta", "no publica referencia" in sin, True)

# ── el plan de Arauco: meta ÷ días operables ÷ jornada ──
print("\nplan_rend_meta (fórmula de la planilla de Arauco, M7 septiembre)")
from generar_tablero_faena import HDISP
for proc, meta, esp in (('VOLTEO', 8855, 30.1), ('PROCESADO', 8060, 27.4), ('CLASIFICADO', 8060, 27.4)):
    eq(f"{proc}", round(meta/28/HDISP, 1), esp)

print("\nplan_arboles · M7 septiembre (meta volteo 8.855, VMA 0,342, 7 días declarados)")
eq("plan de los 7 días declarados", round(plan_arboles(8855, 28, 7, 0.342)), 6473)
eq("un día = plan diario ÷ VMA", round(plan_arboles(8855, 28, 1, 0.342)), 925)
eq("sin meta cargada", plan_arboles(None, 28, 7, 0.342), None)
eq("sin VMA no inventa", plan_arboles(8855, 28, 7, None), None)
eq("sin días declarados", plan_arboles(8855, 28, 0, 0.342), None)
# el período tiene que calzar: doble de días declarados, doble de plan
eq("escala con los días", plan_arboles(8855, 28, 14, 0.342), plan_arboles(8855, 28, 7, 0.342)*2)

print("\nnota_shoveleo · dice por qué no tiene Plan")
n = nota_shoveleo({'horas': 23.5, 'dias': 5, 'equipos': ['HM-05'], 'pct': 53.0, 'dias_pct': 5})
eq("declara que va sin plan", "sin Plan" in n, True)
eq("dice que ya está en la fila Horas", "Horas" in n and "10.5 h" in n, True)

print("\nnota_km · dice por qué el desplazamiento no lleva Plan")
eq("con GPS, explica", "sin Plan" in nota_km({'km_dia': 5.7}, None), True)
eq("sin GPS no ensucia la hoja", nota_km(None, None), "")

print("\nbloque de productividad · 4 columnas como la planilla de Arauco")
src = open(__import__("pathlib").Path(__file__).parent / "generar_informe_faena.py",
           encoding="utf-8").read()
eq("cabecera sin Habitual", "Habitual</th>" in src, False)
eq("colspan del título", "colspan=4>{titulo}" in src, True)
eq("filas de 4 celdas", "for lab, plan, real, cum in filas" in src, True)
eq("el pie ya no la nombra", "<b>Habitual</b>" in src, False)

# ── BARRERA: nombres usados antes de existir ─────────────────────────────────────────
# Por qué está: el 2026-09-09 se usó `plan_arb` UNA LÍNEA antes de definirlo. `ast.parse`
# no lo ve (es sintaxis válida) y las pruebas de arriba tampoco (no ejecutan sheet()), así
# que llegó a producción. El run salió VERDE —el paso del informe es "no crítico": reintenta,
# avisa por Telegram y NO pisa el HTML publicado— y dejó el informe del día anterior en el
# hosting, que parece al día. Es exactamente el modo de falla que ya había mordido el 31-07.
#
# Se filtra a las DOS clases que revientan en runtime; el resto de pyflakes (imports sin usar,
# variables muertas) es ruido preexistente y una barrera que grita siempre no la mira nadie.
# Probada contra un caso bueno y uno malo antes de dejarla acá.
print("\nnombres usados antes de existir (pyflakes)")
import subprocess, pathlib
MORTALES = ('undefined name', 'referenced before assignment', 'local variable defined in enclosing scope')
_here = pathlib.Path(__file__).parent
try:
    for f in ('generar_informe_faena.py', 'generar_tablero_faena.py'):
        r = subprocess.run([sys.executable, '-m', 'pyflakes', str(_here / f)],
                           capture_output=True, text=True)
        graves = [l for l in r.stdout.splitlines() if any(m in l for m in MORTALES)]
        eq(f"{f} sin nombres indefinidos", graves, [])
except FileNotFoundError:
    print("  ⚠️  pyflakes no instalado (pip install pyflakes) — barrera omitida")

print(f"\n{ok} ok · {fail} fallidas")
sys.exit(1 if fail else 0)
