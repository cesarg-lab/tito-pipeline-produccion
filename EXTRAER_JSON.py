#!/usr/bin/env python3
"""
EXTRAER_JSON.py — Extrae el objeto `D` del Dashboard_Cosecha.html y lo guarda como data.json
══════════════════════════════════════════════════════════════════════════════════════════════
Razon de existir:
  El Dashboard_Cosecha.html contiene un objeto JS `const D = {...}` con TODA la data del mes
  (kpis, tiempos perdidos, proyecciones, Pareto, MTBF/MTTR, historico, etc.). Este script
  lo extrae y exporta como data.json para que Tito JARVIS y otros sistemas lo consuman
  via produccion.millalemu.com/data.json sin necesidad de parsear el HTML completo.

Uso:
  python3 EXTRAER_JSON.py

Sale:
  data.json con resumen optimizado (subset relevante, no toda la data cruda)
"""

import json
import re
import sys
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent
HTML_FILE = BASE_DIR / "Dashboard_Cosecha.html"
OUT_FILE = BASE_DIR / "data.json"


def extraer_objeto_D(html_content: str) -> dict:
    """Extrae el objeto JS `const D = {...};` del HTML y lo parsea como dict."""
    # Patron: const D = {...};
    # Hay que matchear hasta el cierre balanceado de llaves
    match = re.search(r'const\s+D\s*=\s*({)', html_content)
    if not match:
        raise ValueError("No se encontro 'const D = {' en el HTML")

    start = match.start(1)
    # Balanceo de llaves para encontrar el cierre
    depth = 0
    in_string = False
    escape = False
    end = -1

    for i in range(start, len(html_content)):
        c = html_content[i]
        if escape:
            escape = False
            continue
        if c == '\\':
            escape = True
            continue
        if c == '"' and not escape:
            in_string = not in_string
            continue
        if in_string:
            continue
        if c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                end = i + 1
                break

    if end == -1:
        raise ValueError("No se pudo encontrar el cierre del objeto D")

    json_str = html_content[start:end]
    return json.loads(json_str)


def construir_resumen(D: dict) -> dict:
    """Construye un dict optimizado para Tito - solo los campos relevantes."""
    cfg = D.get("cfg", {})
    resumen_ej = D.get("resumenEj", {})

    # Resumen ejecutivo del mes
    out = {
        "actualizado": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "fuente": "Dashboard Cosecha Forestal Millalemu",
        "url_dashboard": "http://produccion.millalemu.com",

        "mes": {
            "nombre": cfg.get("mesNombre"),
            "anio": cfg.get("anio"),
            "dias_mes": cfg.get("dm"),
            "dias_con_datos": cfg.get("dd"),
            "dias_restantes": cfg.get("dr"),
        },

        "totales": {
            "acumulado_m3": round(cfg.get("ta", 0), 1),
            "meta_m3": round(cfg.get("tm", 0), 1),
            "cumplimiento_pct": round((cfg.get("ta", 0) / cfg.get("tm", 1)) * 100, 1),
            "proyeccion_fin_mes_m3": round(resumen_ej.get("proyM3", 0), 1),
            "proyeccion_pct": round(resumen_ej.get("proyPct", 0), 1),
            "brecha_proyectada_m3": round(resumen_ej.get("proyM3", 0) - cfg.get("tm", 0), 1),
        },

        "mejor_faena": {
            "nombre": resumen_ej.get("mejor", {}).get("t"),
            "acumulado_m3": resumen_ej.get("mejor", {}).get("a"),
            "meta_m3": resumen_ej.get("mejor", {}).get("m"),
            "cumplimiento_pct": resumen_ej.get("mejor", {}).get("c"),
            "rendimiento_m3_hora": resumen_ej.get("mejor", {}).get("r"),
        },

        "peor_faena": {
            "nombre": resumen_ej.get("peor", {}).get("t"),
            "acumulado_m3": resumen_ej.get("peor", {}).get("a"),
            "meta_m3": resumen_ej.get("peor", {}).get("m"),
            "cumplimiento_pct": resumen_ej.get("peor", {}).get("c"),
            "tiempo_perdido_min": resumen_ej.get("peor", {}).get("tm"),
        },

        "faenas": [
            {
                "nombre": k.get("t"),
                "acumulado_m3": k.get("a"),
                "meta_m3": k.get("m"),
                "cumplimiento_pct": k.get("c"),
                "promedio_diario_m3": k.get("p"),
                "proyeccion_m3": k.get("pr"),
                "brecha_m3": k.get("b"),
                "rendimiento_m3_hora": k.get("r"),
                "horas_trabajadas": k.get("h"),
                "tiempo_perdido_min": k.get("tm"),
                "tiempo_total_turno_min": k.get("tt"),
                "especie": k.get("e"),
                "predio": k.get("pr2"),
                "dias_con_datos": k.get("d"),
            }
            for k in D.get("kpis", [])
        ],

        "tiempos_perdidos_top10_global": [
            {"causa": t.get("n"), "minutos": t.get("m")}
            for t in D.get("tmTop", [])
        ],

        "tiempos_perdidos_por_categoria": [
            {"categoria": c.get("n"), "minutos": c.get("v")}
            for c in D.get("tmCat", [])
        ],

        "tiempos_perdidos_por_faena": D.get("tmTeamCauses", {}),

        "pareto_80_20_global": [
            {
                "causa": p.get("n"),
                "horas": p.get("h"),
                "eventos": p.get("ev"),
                "dias_con_evento": p.get("d"),
                "porcentaje": p.get("pct"),
                "porcentaje_acumulado": p.get("pctAcum"),
                "horas_por_evento": p.get("hPorEv"),
            }
            for p in D.get("tmParetoGlobal", [])
            if p.get("pctAcum", 0) <= 85  # Solo el 80/20 (hasta 85%)
        ],

        "mtbf_mttr_por_equipo": [
            {
                "faena": m.get("t"),
                "fallas": m.get("fallas"),
                "horas_productivas": m.get("hrsProd"),
                "min_reparacion": m.get("minReparacion"),
                "mtbf_horas_entre_fallas": m.get("mtbf"),
                "mttr_min_por_reparacion": m.get("mttr"),
            }
            for m in D.get("mtbfMttr", [])
        ],

        "mix_especies": [
            {
                "codigo": e.get("cod"),
                "nombre": e.get("nom"),
                "volumen_m3": e.get("vol"),
                "porcentaje": e.get("pct"),
                "rendimiento_m3_hora": e.get("rend"),
            }
            for e in D.get("espMix", [])
        ],

        "ranking_eficiencia": [
            {
                "faena": r.get("t"),
                "volumen_m3": r.get("vol"),
                "horas": r.get("hrs"),
                "rendimiento_m3_hora": r.get("rend"),
            }
            for r in D.get("rankingEf", [])
        ],

        "predios_top": [
            {
                "predio": p.get("pr"),
                "volumen_m3": p.get("vol"),
                "horas": p.get("hrs"),
                "rendimiento_m3_hora": p.get("rend"),
                "equipos": p.get("eq"),
                "especies": p.get("esp"),
            }
            for p in D.get("predios", [])
        ],

        "tendencia_diaria_global": [
            {"dia": t.get("d"), "produccion_m3": t.get("v")}
            for t in D.get("trend", [])
        ],

        "avance_acumulado_vs_plan": [
            {"dia": a.get("d"), "real_m3": a.get("real"), "plan_m3": a.get("plan")}
            for a in D.get("avance", [])
        ],

        "dias_sin_produccion": D.get("diasSinProd", []),

        "historico_meses_anteriores": [
            {
                "mes": h.get("mesNombre"),
                "anio": h.get("anio"),
                "faena": h.get("equipo"),
                "meta_m3": h.get("meta"),
                "volumen_m3": h.get("vol"),
                "cumplimiento_pct": h.get("cumpl"),
                "promedio_diario_m3": h.get("promDia"),
                "dias_trabajados": h.get("diasTrab"),
            }
            for h in D.get("historico", [])
        ],

        "observaciones_analisis": {
            "total_comentarios": D.get("obsAnalisis", {}).get("totalComentarios"),
            "por_categoria": D.get("obsAnalisis", {}).get("porCategoria", []),
            "top_comentarios": [
                {"texto": t.get("txt"), "minutos": t.get("minutos"), "categoria": t.get("cat")}
                for t in D.get("obsAnalisis", {}).get("topComentarios", [])[:10]
            ],
            "palabras_clave": D.get("obsAnalisis", {}).get("palabrasClave", [])[:15],
        },
    }

    return out


def main():
    if not HTML_FILE.exists():
        print(f"ERROR: No existe {HTML_FILE}")
        print("Ejecuta primero: python3 GENERAR_HTML.py")
        sys.exit(1)

    print(f"Leyendo {HTML_FILE.name}...")
    html_content = HTML_FILE.read_text(encoding="utf-8")

    print("Extrayendo objeto D del HTML...")
    D = extraer_objeto_D(html_content)
    print(f"  -> {len(D)} claves de nivel superior")

    print("Construyendo resumen optimizado para Tito...")
    resumen = construir_resumen(D)

    print(f"Guardando {OUT_FILE.name}...")
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(resumen, f, ensure_ascii=False, indent=2)

    size_kb = OUT_FILE.stat().st_size / 1024
    print(f"OK - {OUT_FILE.name}: {size_kb:.1f} KB")

    # Resumen rapido para verificar
    totales = resumen.get("totales", {})
    print()
    print(f"Mes: {resumen['mes']['nombre']} {resumen['mes']['anio']}")
    print(f"Acumulado: {totales.get('acumulado_m3')} / {totales.get('meta_m3')} m3 ({totales.get('cumplimiento_pct')}%)")
    print(f"Proyeccion: {totales.get('proyeccion_fin_mes_m3')} m3 ({totales.get('proyeccion_pct')}%)")
    print(f"Faenas: {len(resumen.get('faenas', []))}")
    print(f"Top 10 TM: {len(resumen.get('tiempos_perdidos_top10_global', []))}")
    print(f"Pareto 80/20: {len(resumen.get('pareto_80_20_global', []))} causas")


if __name__ == "__main__":
    main()
