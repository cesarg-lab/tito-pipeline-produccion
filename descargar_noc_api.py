#!/usr/bin/env python3
"""
descargar_noc_api.py — Descarga datos de GeoNOC vía API REST (sin Selenium)
═══════════════════════════════════════════════════════════════════════════════
Reemplaza a descargar_noc.py eliminando la dependencia de Selenium y Chrome.
Usa directamente la API REST de GeoNOC + autenticación ArcGIS Enterprise.

Requisitos:
  pip install requests

Uso:
  python3 descargar_noc_api.py
"""

import os
import sys
import json
import csv
import subprocess
import logging
from datetime import datetime, timedelta
from pathlib import Path

import requests
# Suprimir warnings de SSL (la red corporativa puede tener certificados internos)
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── Configuración ──────────────────────────────────────────────────────────
BASE_DIR    = Path(__file__).parent
CONFIG_FILE = BASE_DIR / ".noc_config.json"
LOG_FILE    = BASE_DIR / "descarga_log.txt"

# Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_FILE, encoding="utf-8")
    ]
)
log = logging.getLogger(__name__)

# ── Cargar configuración ───────────────────────────────────────────────────
if not CONFIG_FILE.exists():
    log.error("No se encontró .noc_config.json")
    sys.exit(1)

with open(CONFIG_FILE, encoding="utf-8") as f:
    cfg = json.load(f)

USUARIO  = cfg["username"]     # 79662560-0
PASSWORD = cfg["password"]     # Arauco2020
EMSEFOR  = cfg.get("emsefor", USUARIO)  # Normalmente es el mismo RUT

# URLs de la API
PORTAL_URL    = "https://araucaria.arauco.com/portal"
TOKEN_URL     = f"{PORTAL_URL}/sharing/rest/generateToken"
GEONOC_API    = "https://geonoc.arauco.com/geonocAPI/informevistanoc"
GEONOC_BASE   = "https://geonoc.arauco.com"

# ── Funciones de autenticación ─────────────────────────────────────────────
def obtener_token_arcgis(session: requests.Session) -> str:
    """Obtiene un token de ArcGIS Enterprise vía REST API."""
    log.info("Obteniendo token de ArcGIS Enterprise...")

    data = {
        "username": USUARIO,
        "password": PASSWORD,
        "client": "referer",
        "referer": GEONOC_BASE,
        "f": "json",
        "expiration": 20160,  # 14 días en minutos
    }

    resp = session.post(TOKEN_URL, data=data, verify=False, timeout=30)
    resp.raise_for_status()
    result = resp.json()

    if "token" in result:
        token = result["token"]
        expires = datetime.fromtimestamp(result.get("expires", 0) / 1000)
        log.info(f"Token obtenido — expira: {expires.strftime('%d/%m/%Y %H:%M')}")
        return token
    elif "error" in result:
        err = result["error"]
        log.error(f"Error de autenticación: {err.get('message', err)}")
        sys.exit(1)
    else:
        log.error(f"Respuesta inesperada: {result}")
        sys.exit(1)


def establecer_sesion(session: requests.Session, token: str):
    """Establece sesión con geonoc.arauco.com usando el token de ArcGIS."""
    # Método 1: Acceder al portal con el token para obtener cookies de sesión
    log.info("Estableciendo sesión con GeoNOC...")

    # Registrar el token como cookie de ArcGIS
    session.cookies.set("esri_aopc", token, domain="geonoc.arauco.com")

    # También intentar acceder a la app de planificación con el token
    # para que el servidor valide y establezca cookies de sesión
    try:
        resp = session.get(
            f"{GEONOC_BASE}/planificacion/",
            params={"token": token},
            verify=False,
            timeout=30,
            allow_redirects=True
        )
        log.info(f"Sesión establecida — status: {resp.status_code}")
    except Exception as e:
        log.warning(f"Advertencia al establecer sesión: {e}")


# ── Funciones de descarga ──────────────────────────────────────────────────
def descargar_reporte(session: requests.Session, token: str,
                      reporte: str, fecha_ini: str, fecha_fin: str) -> list:
    """
    Descarga un reporte de GeoNOC vía API REST.

    Args:
        reporte: "PG" (Productividad Genérico) o "TP" (Tiempos Perdidos)
        fecha_ini: "2026-04-01" (formato YYYY-MM-DD)
        fecha_fin: "2026-04-16" (formato YYYY-MM-DD)

    Returns:
        Lista de diccionarios con los registros
    """
    nombre = "Productividad Genérico" if reporte == "PG" else "Tiempos Perdidos"
    log.info(f"Descargando {nombre} ({fecha_ini} → {fecha_fin})...")

    payload = [{
        "EMSEFOR": EMSEFOR,
        "FECHA_INI": fecha_ini,
        "FECHA_FIN": fecha_fin,
        "ZONA": None,
        "REPORTE": reporte
    }]

    headers = {
        "Content-Type": "application/json",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": f"{GEONOC_BASE}/planificacion/",
    }

    # Intentar con token en cookie, en header, y en URL
    for intento, auth_method in enumerate([
        lambda: None,  # Solo cookies de sesión
        lambda: headers.update({"Authorization": f"Bearer {token}"}),
        lambda: None,  # Se maneja abajo con params
    ], 1):
        auth_method()

        params = {"token": token} if intento == 3 else {}

        try:
            resp = session.post(
                GEONOC_API,
                json=payload,
                headers=headers,
                params=params,
                verify=False,
                timeout=120
            )

            if resp.status_code == 200:
                data = resp.json()
                if isinstance(data, list) and len(data) > 0:
                    log.info(f"  ✅ {nombre}: {len(data)} registros (método {intento})")
                    return data
                elif isinstance(data, dict) and "error" in data:
                    log.warning(f"  Método {intento} — error API: {data['error']}")
                    continue
                else:
                    log.info(f"  {nombre}: 0 registros (respuesta vacía)")
                    return []
            elif resp.status_code in (401, 403):
                log.warning(f"  Método {intento} — auth rechazada ({resp.status_code})")
                continue
            else:
                log.warning(f"  Método {intento} — HTTP {resp.status_code}")
                continue

        except Exception as e:
            log.warning(f"  Método {intento} — error: {e}")
            continue

    log.error(f"❌ No se pudo descargar {nombre} tras 3 métodos de auth")
    return []


# ── Conversión JSON → CSV ─────────────────────────────────────────────────

# Mapeo de campos JSON → columnas CSV para Productividad Genérico
PG_HEADERS = (
    "Número Noc;Fecha NOC;Intervención;Unidad Operativa;Rut Empresario;"
    "Nombre Empresario;Equipo;Tipo Equipo;Código Calibrador;Nombre Calibrador;"
    "Código Predio;Origen;Id Especie;Desc Especie;Fecha Inicio;Fecha Fin;"
    "Hora Inicio;Hora Fin;Horómetro;Tiempo Colacion;Tiempo Efectivo;"
    "Número Ciclos;Número Personas;Árboles Madereados;Volumen SSC PU;Volumen SSC AS;"
    + ";".join(str(i) for i in range(1, 72)) + ";"
    "Id Zona;Zona;NOC Completa;Zona Predio;Zona Movil;Zona Cosecha;"
    "Número Acta;Secuencia Acta"
)

PG_FIELDS = [
    "numero_noc", "hora_inicio", "intervencion", "unidad_operativa",
    "rut_empresario", "nombre_empresario", "equipo", "tipo_equipo",
    "rut_calibrador", "nombre_calibrador", "codigo_predio", "origen",
    "id_epecie", "desc_especie", "hora_inicio", "hora_fin",
    "hora_inicio", "hora_fin", "horometro", "horas_colacion",
    "tiempo_efectivo", "numero_ciclos", "numero_personas",
    "arboles_madereados", "m3ssc_pu", "m3ssc_as",
] + [str(i) for i in range(1, 72)] + [
    "id_zona", "zona", "noc_completa", "zona_predio", "zona_movil",
    "zona_cosecha", "Numero_Acta", "secuencia_acta"
]

TP_HEADERS = (
    "N°;Empresario;Fecha;Número Noc;Estado Noc;Código Equipo;"
    "Código Tiempo Perdido;Descripción;Tiempo (Min);Observación"
)

TP_FIELDS = [
    "N", "Empresario", "Fecha", "Numero_noc", "estado_noc",
    "codigo_equipo", "codigo_tiempo_perdido", "descripcion",
    "tiempo", "observacion"
]


def formato_fecha(valor):
    """Convierte '2026-04-01T00:04:24' → '01-04-2026'."""
    if not valor:
        return ""
    try:
        if "T" in str(valor):
            dt = datetime.fromisoformat(str(valor))
        else:
            dt = datetime.strptime(str(valor), "%Y-%m-%d")
        return dt.strftime("%d-%m-%Y")
    except:
        return str(valor) if valor else ""


def formato_numero(valor):
    """Formatea números: usa coma como separador decimal."""
    if valor is None:
        return ""
    if isinstance(valor, float):
        # Usar coma como separador decimal (formato chileno)
        return str(valor).replace(".", ",")
    return str(valor)


def valor_a_segundos(valor):
    """Convierte un valor de hora a segundos desde medianoche.
    Maneja: número (ya en segundos), ISO datetime string, o HH:MM:SS."""
    if valor is None:
        return ""
    if isinstance(valor, (int, float)):
        return str(int(valor))
    val_str = str(valor)
    try:
        if "T" in val_str:
            # ISO datetime: "2026-03-30T10:00:00" → extraer hora y convertir a segundos
            dt = datetime.fromisoformat(val_str)
            return str(dt.hour * 3600 + dt.minute * 60 + dt.second)
        elif ":" in val_str:
            # HH:MM:SS
            parts = val_str.split(":")
            h, m, s = int(parts[0]), int(parts[1]), int(parts[2]) if len(parts) > 2 else 0
            return str(h * 3600 + m * 60 + s)
    except:
        pass
    return str(valor) if valor else ""


def registro_pg_a_csv(rec: dict) -> str:
    """Convierte un registro JSON de Productividad Genérico a línea CSV."""
    valores = []
    for i, campo in enumerate(PG_FIELDS):
        val = rec.get(campo)

        # FIX 2026-05-22 distribución por día Arauco:
        # Posiciones 1, 14, 15 = columnas Fecha NOC / Fecha Inicio / Fecha Fin → día (dd-mm-yyyy)
        # Usamos hora_inicio/hora_fin (día operativo real) en lugar de fecha (cierre administrativo)
        # que ponía turnos nocturnos en el día siguiente
        if i in (1, 14, 15):
            val = formato_fecha(val)
        # Posiciones 16, 17 = columnas Hora Inicio / Hora Fin → segundos desde medianoche
        elif i in (16, 17):
            val = valor_a_segundos(val)
        # Campos numéricos con decimales
        elif campo in ("m3ssc_pu", "m3ssc_as", "horometro", "tiempo_efectivo",
                        "horas_colacion") or campo.isdigit():
            val = formato_numero(val)
        # RUT sin guión
        elif campo in ("rut_empresario", "rut_calibrador"):
            val = str(val).replace("-", "") if val else ""
        else:
            val = str(val) if val is not None else ""

        valores.append(val)

    return ";".join(valores)


def registro_tp_a_csv(rec: dict, indice: int) -> str:
    """Convierte un registro JSON de Tiempos Perdidos a línea CSV."""
    valores = []
    for campo in TP_FIELDS:
        val = rec.get(campo)
        if campo == "Fecha":
            val = formato_fecha(val)
        elif campo == "N":
            val = str(indice) if val is None else str(val)
        else:
            val = str(val) if val is not None else ""
        valores.append(val)

    return ";".join(valores)


def guardar_csv(datos: list, reporte: str, destino: Path):
    """Guarda los datos como CSV en formato compatible con el existente."""
    if reporte == "PG":
        nombre = "ProductividadGenerico.csv"
        headers = PG_HEADERS
        parse_fn = lambda rec, i: registro_pg_a_csv(rec)
    else:
        nombre = "TiemposPerdidos.csv"
        headers = TP_HEADERS
        parse_fn = lambda rec, i: registro_tp_a_csv(rec, i)

    filepath = destino / nombre

    with open(filepath, "w", encoding="utf-8-sig", newline="") as f:
        f.write(headers + "\r\n")
        for i, rec in enumerate(datos, 1):
            linea = parse_fn(rec, i)
            f.write(linea + "\r\n")

    size_kb = filepath.stat().st_size / 1024
    log.info(f"  💾 {nombre}: {len(datos)} registros, {size_kb:.1f} KB")
    return filepath


# ══════════════════════════════════════════════════════════════════════════════
# FLUJO PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
def main():
    log.info("=" * 60)
    log.info(f"Descarga NOC vía API — {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    # Calcular rango de fechas: 1er día del mes hasta hoy
    hoy = datetime.now()
    primer_dia = hoy.replace(day=1)
    fecha_ini = primer_dia.strftime("%Y-%m-%d")
    fecha_fin = hoy.strftime("%Y-%m-%d")
    log.info(f"📅 Rango: {fecha_ini} → {fecha_fin}")

    # Eliminar CSVs previos
    for fname in ["ProductividadGenerico.csv", "TiemposPerdidos.csv"]:
        fpath = BASE_DIR / fname
        if fpath.exists():
            fpath.unlink()
            log.info(f"  🗑️  Eliminado: {fname}")

    # Crear sesión HTTP
    session = requests.Session()
    session.verify = False
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/131.0.0.0 Safari/537.36",
    })

    # Obtener token
    token = obtener_token_arcgis(session)

    # Establecer sesión con GeoNOC
    establecer_sesion(session, token)

    # Descargar Productividad Genérico
    datos_pg = descargar_reporte(session, token, "PG", fecha_ini, fecha_fin)
    pg_ok = False
    if datos_pg:
        guardar_csv(datos_pg, "PG", BASE_DIR)
        pg_ok = True
    else:
        log.warning("⚠️  Sin datos de Productividad Genérico")

    # Descargar Tiempos Perdidos
    datos_tp = descargar_reporte(session, token, "TP", fecha_ini, fecha_fin)
    tp_ok = False
    if datos_tp:
        guardar_csv(datos_tp, "TP", BASE_DIR)
        tp_ok = True
    else:
        log.warning("⚠️  Sin datos de Tiempos Perdidos")

    # Resumen
    if pg_ok and tp_ok:
        log.info("✅ Ambos archivos descargados correctamente")
    elif pg_ok or tp_ok:
        log.info("⚠️  Solo se descargó un archivo")
    else:
        log.error("❌ No se descargó ningún archivo")
        sys.exit(1)

    # Actualizar Dashboard
    script_update = BASE_DIR / "ACTUALIZAR_DASHBOARD.py"
    if pg_ok and tp_ok and script_update.exists():
        log.info("📊 Cerrando Excel si está abierto...")
        subprocess.run(
            ["osascript", "-e", 'tell application "Microsoft Excel" to quit saving no'],
            capture_output=True, text=True, timeout=15
        )
        import time; time.sleep(3)
        log.info("📊 Actualizando Dashboard Excel...")
        try:
            res = subprocess.run(
                [sys.executable, str(script_update)],
                capture_output=True, text=True, timeout=120
            )
            if res.returncode == 0:
                log.info("✅ Dashboard actualizado correctamente")
            else:
                log.warning(f"⚠️  Error en actualización: {res.stderr[:2000]}")
        except Exception as e:
            log.error(f"❌ Error ejecutando actualización: {e}")

    log.info("🎯 Proceso finalizado")
    log.info("=" * 60)


if __name__ == "__main__":
    main()
