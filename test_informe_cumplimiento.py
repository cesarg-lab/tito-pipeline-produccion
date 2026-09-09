"""Pruebas de lo que se agregó el 2026-09-09: rend real de volteo + horas por día del pre-uso."""
import sys; sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent))
from generar_informe_faena import rend_declarado, horas_preuso, nota_ref, FAENA_ID

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

print("\nbloque de productividad · 4 columnas como la planilla de Arauco")
src = open(__import__("pathlib").Path(__file__).parent / "generar_informe_faena.py",
           encoding="utf-8").read()
eq("cabecera sin Habitual", "Habitual</th>" in src, False)
eq("colspan del título", "colspan=4>{titulo}" in src, True)
eq("filas de 4 celdas", "for lab, plan, real, cum in filas" in src, True)
eq("el pie ya no la nombra", "<b>Habitual</b>" in src, False)

print(f"\n{ok} ok · {fail} fallidas")
sys.exit(1 if fail else 0)
