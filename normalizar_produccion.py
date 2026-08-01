"""
Normalizador de CSV de producción.

A partir del 2026-04-28 el pipeline usa **Base 2 NOC** en lugar de
Productividad Genérico (PG dejó de funcionar de forma confiable vía
Selenium en GeoNOC). Base 2 NOC trae las mismas columnas + extras,
pero con nombres y formatos distintos.

Este módulo se importa desde ACTUALIZAR_DASHBOARD.py, GENERAR_HTML.py
y GENERAR_IMAGEN.py para mantener una sola lógica de mapeo.
"""

import pandas as pd
from datetime import datetime, timedelta

# Hora a la que se da por terminada la jornada en terreno (Chile). El turno es de 10,5 h y
# cierra entre 19:00 y 19:30; pasada esta hora el día ya no tiene producción por delante.
FIN_JORNADA_H = 20


def dia_en_curso(anio: int, mes: int, ult_dia: int, fin_jornada_h: int = FIN_JORNADA_H) -> bool:
    """¿El último día con datos es HOY y la jornada todavía NO termina?

    Una sola definición para toda la salida (grilla, resumen, KPIs, informes). Si el día está
    en curso hay que proyectar lo que le falta; si ya cerró, no hay nada que proyectar y la
    proyección del mes tiene que dar exactamente el acumulado.

    Sin esto, la corrida de las 22:47 del último día del mes mandaba TRES cifras distintas del
    mismo mes cerrado (31-07-2026: dashboard 48.993, kpis 49.6k, Telegram 50.573).
    """
    hoy = datetime.utcnow() - timedelta(hours=4)        # Chile = UTC-4
    return (hoy.year == anio and hoy.month == mes
            and hoy.day == ult_dia and hoy.hour < fin_jornada_h)


def _hhmm_a_segundos(valor):
    """Convierte 'HH:MM' o 'H:MM' o 'HH:MM:SS' a segundos enteros."""
    if pd.isna(valor) or valor in ("", None):
        return 0
    try:
        partes = str(valor).strip().split(":")
        h = int(partes[0]) if partes[0] else 0
        m = int(partes[1]) if len(partes) > 1 and partes[1] else 0
        s = int(partes[2]) if len(partes) > 2 and partes[2] else 0
        return h * 3600 + m * 60 + s
    except (ValueError, AttributeError):
        return 0


def normalizar(prod: pd.DataFrame) -> pd.DataFrame:
    """
    Recibe un DataFrame leído de CSV de producción y devuelve uno
    con las columnas que esperan los scripts del pipeline:

    - Equipo, Fecha NOC, Desc Especie, Predio
    - Volumen SSC PU, Volumen SSC AS
    - Tiempo Efectivo (segundos), Hora Inicio (segundos), Hora Fin (segundos)
    - Número Ciclos, Árboles Madereados

    Detecta automáticamente si viene en formato Base 2 NOC (mayúsculas)
    o Productividad Genérico (legacy). Si ya está en formato PG, lo
    devuelve tal cual.
    """
    # Si ya está en formato PG (compatibilidad con CSVs viejos), no toca nada
    if "Equipo" in prod.columns and "Fecha NOC" in prod.columns:
        return prod

    # Formato Base 2 NOC → mapear
    if "EQUIPO" in prod.columns and "FECHA" in prod.columns:
        prod = prod.rename(columns={
            "EQUIPO": "Equipo",
            "FECHA": "Fecha NOC",
            "ESPECIE": "Desc Especie",
            "NOMBRE_PREDIO": "Origen",      # PG usaba 'Origen' para predio de cosecha
            "NUMERO_CICLOS": "Número Ciclos",
            "M3SSC": "Volumen SSC PU",
            "TIPO_EQUIPO": "Tipo Equipo",
            "FOLIO": "Número Noc",
            "CODIGO_PREDIO": "Código Predio",
        })

        # Volumen — Base 2 NOC trae M3SSC único; PG separaba PU/AS
        if "Volumen SSC PU" not in prod.columns:
            prod["Volumen SSC PU"] = 0
        prod["Volumen SSC AS"] = 0

        # Árboles — Base 2 NOC no lo trae; KPI de harvester quedará en 0
        if "Árboles Madereados" not in prod.columns:
            prod["Árboles Madereados"] = 0

        # Horómetro — Base 2 NOC no lo trae
        if "Horómetro" not in prod.columns:
            prod["Horómetro"] = 0

        # Tiempos — Base 2 NOC trae HH:MM, los scripts esperan segundos
        if "TIEMPO_EFECTIVO" in prod.columns:
            prod["Tiempo Efectivo"] = prod["TIEMPO_EFECTIVO"].apply(_hhmm_a_segundos)
        if "HORA_INICIO" in prod.columns:
            prod["Hora Inicio"] = prod["HORA_INICIO"].apply(_hhmm_a_segundos)
        if "HORA_TERMINO" in prod.columns:
            prod["Hora Fin"] = prod["HORA_TERMINO"].apply(_hhmm_a_segundos)

        # Asegurar que Número Ciclos sea numérico
        if "Número Ciclos" in prod.columns:
            prod["Número Ciclos"] = pd.to_numeric(
                prod["Número Ciclos"], errors="coerce"
            ).fillna(0).astype(int)

        return prod

    # Formato desconocido — devolver tal cual y dejar que el script falle
    # con un error legible
    return prod


if __name__ == "__main__":
    # Prueba rápida desde CLI
    import sys
    from pathlib import Path

    if len(sys.argv) < 2:
        print("Uso: python3 normalizar_produccion.py <ruta_csv>")
        sys.exit(1)

    df = pd.read_csv(sys.argv[1], sep=";", encoding="utf-8-sig")
    print(f"Original: {len(df)} filas, {len(df.columns)} columnas")
    print(f"Columnas: {list(df.columns)[:8]}...")
    df = normalizar(df)
    print(f"\nNormalizado: {len(df)} filas, {len(df.columns)} columnas")
    print(f"Columnas clave presentes:")
    for c in ["Equipo", "Fecha NOC", "Desc Especie", "Volumen SSC PU",
              "Volumen SSC AS", "Tiempo Efectivo", "Hora Inicio",
              "Hora Fin", "Número Ciclos", "Árboles Madereados",
              "Origen", "Tipo Equipo", "Número Noc", "Horómetro"]:
        marca = "✅" if c in df.columns else "❌"
        print(f"  {marca} {c}")
    print(f"\nMuestra:\n{df[['Equipo','Fecha NOC','Desc Especie','Volumen SSC PU','Tiempo Efectivo','Hora Inicio','Hora Fin']].head(3)}")
