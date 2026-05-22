#!/usr/bin/env python3
"""
descargar_noc_api.py — Descarga GeoNOC vía API REST (sin Selenium)
═══════════════════════════════════════════════════════════════════════════════
Versión cloud-ready validada al peso 2026-05-21:
  - Auth: POST /portal/sharing/rest/generateToken (form-urlencoded)
  - API:  POST /geonocAPI/informevistanoc con header `Authorization: Bearer {token}`

Diferencias vs versión Mac:
  - Lee credenciales desde ENV vars (ARAUCO_USER, ARAUCO_PASS) — no .noc_config.json
  - Eliminados los 3 métodos de auth fallidos del intento anterior
  - Eliminada la sub-rutina establecer_sesion (innecesaria con Bearer)
  - Eliminada llamada a ACTUALIZAR_DASHBOARD.py (Mac/Excel only)
  - Eliminado subprocess osascript (Mac only)

Uso (Render cron job):
  ARAUCO_USER=<RUT_empresa> ARAUCO_PASS=<password> python3 descargar_noc_api.py
"""

import os
import sys
import json
import logging
from datetime import datetime
from pathlib import Path

try:
    from zoneinfo import ZoneInfo
    CHILE_TZ = ZoneInfo("America/Santiago")
except Exception:
    # Fallback Python < 3.9 o sin tzdata: offset fijo UTC-4 (Chile verano)
    from datetime import timezone, timedelta
    CHILE_TZ = timezone(timedelta(hours=-4))

import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── Configuración ──────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
LOG_FILE = BASE_DIR / "descarga_log.txt"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_FILE, encoding="utf-8")
    ]
)
log = logging.getLogger(__name__)

# ── Credenciales desde ENV (con fallback a .noc_config.json para dev local) ──
USUARIO = os.environ.get("ARAUCO_USER")
PASSWORD = os.environ.get("ARAUCO_PASS")

if not USUARIO or not PASSWORD:
    CONFIG_FILE = BASE_DIR / ".noc_config.json"
    if CONFIG_FILE.exists():
        log.info("ENV vars no presentes — usando .noc_config.json (modo dev local)")
        with open(CONFIG_FILE, encoding="utf-8") as f:
            cfg = json.load(f)
        USUARIO = cfg["username"]
        PASSWORD = cfg["password"]
    else:
        log.error("Faltan credenciales: definir ENV vars ARAUCO_USER y ARAUCO_PASS")
        sys.exit(1)

EMSEFOR = USUARIO  # mismo RUT empresa

# URLs ArcGIS Enterprise + GeoNOC
PORTAL_URL = "https://araucaria.arauco.com/portal"
TOKEN_URL = f"{PORTAL_URL}/sharing/rest/generateToken"
GEONOC_API = "https://geonoc.arauco.com/geonocAPI/informevistanoc"
GEONOC_BASE = "https://geonoc.arauco.com"


# ── Autenticación ──────────────────────────────────────────────────────────
def obtener_token_arcgis(session: requests.Session) -> str:
    """Obtiene token ArcGIS via POST form-urlencoded a generateToken.

    Validado al peso 2026-05-21 — devuelve token con expiración 14 días.
    """
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
        expires_ms = result.get("expires", 0)
        expires = datetime.fromtimestamp(expires_ms / 1000) if expires_ms else None
        log.info(f"Token obtenido — expira: {expires.strftime('%d/%m/%Y %H:%M') if expires else 'desconocido'}")
        return token

    err = result.get("error", {})
    log.error(f"Error de autenticación: {err.get('message', err) if err else result}")
    sys.exit(1)


# ── Descarga de reportes ───────────────────────────────────────────────────
def descargar_reporte(session: requests.Session, token: str,
                      reporte: str, fecha_ini: str, fecha_fin: str) -> list:
    """Descarga un reporte de GeoNOC con Authorization Bearer.

    Método validado al peso 2026-05-21: 199 registros TP bajados en ~2 seg.

    Args:
        reporte: "PG" (Productividad Genérico) o "TP" (Tiempos Perdidos)
        fecha_ini: "YYYY-MM-DD"
        fecha_fin: "YYYY-MM-DD"
    """
    nombre = "Productividad Genérico" if reporte == "PG" else "Tiempos Perdidos"
    log.info(f"Descargando {nombre} ({fecha_ini} → {fecha_fin})...")

    payload = [{
        "EMSEFOR": EMSEFOR,
        "FECHA_INI": fecha_ini,
        "FECHA_FIN": fecha_fin,
        "ZONA": None,
        "REPORTE": reporte,
    }]

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": f"{GEONOC_BASE}/planificacion/",
    }

    try:
        resp = session.post(
            GEONOC_API,
            json=payload,
            headers=headers,
            verify=False,
            timeout=120,
        )

        if resp.status_code == 200:
            data = resp.json()
            if isinstance(data, list):
                log.info(f"  ✅ {nombre}: {len(data)} registros")
                return data
            if isinstance(data, dict) and "error" in data:
                log.warning(f"  Error API: {data['error']}")
                return []
            log.info(f"  {nombre}: respuesta vacía o inesperada")
            return []

        log.warning(f"  HTTP {resp.status_code} — body: {resp.text[:300]}")
        return []

    except Exception as e:
        log.error(f"  ❌ Excepción: {e}")
        return []


# ── Conversión JSON → CSV ─────────────────────────────────────────────────
# Base 2 NOC (REPORTE=BN) — 49 columnas, mismo header que el CSV que descarga Selenium
BN_HEADERS = (
    "FOLIO;FECHA;CODIGO_PREDIO;NOMBRE_PREDIO;EQUIPO;UNIDAD_OPERATIVA;INTERVENCION;"
    "HORA_INICIO;HORA_TERMINO;TIEMPO_EFECTIVO;HORAS_COLACION;NUMERO_CICLOS;NUMERO_PERSONAS;"
    "RUT_EMPRESA;NOMBRE_EMPRESA;NUMERO_ACTA;SECUENCIA;RUT_CALIBRADOR;NOMBRE_CALIBRADOR;"
    "DIAMETRO;CODIGO_PRODUCTO;FACTOR_CORTEZA;ESPECIE;TROZOS;CODIGO_DESTINO;"
    "LARGO;LARGO_REAL;M3SSC;M3SSC_CABEZAL;STOCK;TIEMPOS_MUERTOS;TOTAL_HR_TMP_MUERTOS;"
    "TIPO_NOC;ESTADO_MADERA;FECHA_REGISTRO;FECHA_CORTE;ID_DETALLE_NOC;OBSERVACION;"
    "TIPO_EQUIPO;BIOMASA;PRODUCTO;NOMBRE_DESTINO;TEMPORADA;COMPLETA;"
    "ZONA_FORESTAL_PREDIO;ZONA_FORESTAL_MOVIL;ZONA_COSECHA;FSC_PRODUCCION;CERTFOR_PRODUCCION"
)

# Mapeo de columnas CSV → claves del JSON de Arauco
# Notar: API usa snake_case minúsculas mezclado con UPPERCASE en algunos campos
# y typo "ZONA_COCECHA" (que en CSV se llama ZONA_COSECHA)
BN_FIELDS = [
    "folio", "fecha", "codigo_predio", "nombre_predio", "equipo",
    "unidad_operativa", "intervencion", "hora_inicio", "hora_termino",
    "tiempo_efectivo", "horas_colacion", "numero_ciclos", "numero_personas",
    "rut_empresa", "nombre_empresa", "Numero_Acta", "secuencia",
    "rut_calibrador", "nombre_calibrador", "diametro", "codigo_producto",
    "factor_corteza", "especie", "trozos", "codigo_destino",
    "largo", "largo_real", "m3ssc", "m3ssc_cabezal", "stock",
    "tiempos_muertos", "TOTAL_HR_TMP_MUERTOS", "TIPO_NOC", "estado_madera",
    "FECHA_REGISTRO", "FECHA_CORTE", "ID_DETALLE_NOC", "OBSERVACION",
    "tipo_equipo", "BIOMASA", "PRODUCTO", "NOMBRE_DESTINO", "TEMPORADA",
    "completa", "ZONA_FORESTAL_PREDIO", "ZONA_FORESTAL_MOVIL", "ZONA_COCECHA",
    "FSC_PRODUCCION", "CERTFOR_PRODUCCION"
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
    if not valor:
        return ""
    try:
        s = str(valor)
        # Soporte 'Z' suffix (UTC) que fromisoformat<3.11 no acepta
        if s.endswith("Z"):
            s = s[:-1] + "+00:00"
        if "T" in s:
            dt = datetime.fromisoformat(s)
            # Si Arauco devuelve datetime timezone-aware (UTC u otro offset),
            # convertir a hora Chile antes de formatear el día — sino el último
            # turno del día se desplaza al día siguiente UTC
            if dt.tzinfo is not None:
                dt = dt.astimezone(CHILE_TZ)
        else:
            dt = datetime.strptime(s, "%Y-%m-%d")
        return dt.strftime("%d-%m-%Y")
    except Exception:
        return str(valor) if valor else ""


def formato_numero(valor):
    if valor is None:
        return ""
    if isinstance(valor, float):
        return str(valor).replace(".", ",")
    return str(valor)


def valor_a_segundos(valor):
    if valor is None:
        return ""
    if isinstance(valor, (int, float)):
        return str(int(valor))
    val_str = str(valor)
    try:
        if val_str.endswith("Z"):
            val_str = val_str[:-1] + "+00:00"
        if "T" in val_str:
            dt = datetime.fromisoformat(val_str)
            # Misma conversión timezone: si viene aware, convertir a hora Chile
            if dt.tzinfo is not None:
                dt = dt.astimezone(CHILE_TZ)
            return str(dt.hour * 3600 + dt.minute * 60 + dt.second)
        if ":" in val_str:
            parts = val_str.split(":")
            h, m = int(parts[0]), int(parts[1])
            s = int(parts[2]) if len(parts) > 2 else 0
            return str(h * 3600 + m * 60 + s)
    except Exception:
        pass
    return str(valor) if valor else ""


def _iso_a_hhmm(valor):
    """Convierte ISO timestamp '2026-04-29T11:00:00' → 'HH:MM' en hora Chile."""
    if not valor:
        return ""
    try:
        s = str(valor)
        if s.endswith("Z"):
            s = s[:-1] + "+00:00"
        if "T" in s:
            dt = datetime.fromisoformat(s)
            if dt.tzinfo is not None:
                dt = dt.astimezone(CHILE_TZ)
            return dt.strftime("%H:%M")
    except Exception:
        pass
    return str(valor) if valor else ""


def _minutos_a_hmm(valor):
    """Convierte número de minutos (ej. 60) → formato 'H:MM' (1:00)."""
    if valor is None or valor == "":
        return ""
    try:
        mins = int(valor)
        h, m = divmod(mins, 60)
        return f"{h}:{m:02d}"
    except Exception:
        return str(valor) if valor else ""


def _fecha_registro(valor):
    """Convierte ISO '2026-05-01T08:49:26' → 'dd-mm-yyyy HH:MM:SS' en hora Chile."""
    if not valor:
        return ""
    try:
        s = str(valor)
        if s.endswith("Z"):
            s = s[:-1] + "+00:00"
        if "T" in s:
            dt = datetime.fromisoformat(s)
            if dt.tzinfo is not None:
                dt = dt.astimezone(CHILE_TZ)
            return dt.strftime("%d-%m-%Y %H:%M:%S")
    except Exception:
        pass
    return str(valor) if valor else ""


def registro_bn_a_csv(rec: dict) -> str:
    """Convierte 1 registro JSON Base 2 NOC → 1 línea CSV con formato Selenium."""
    valores = []
    for campo in BN_FIELDS:
        val = rec.get(campo)
        # Campos fecha solo (dd-mm-yyyy)
        if campo in ("fecha", "FECHA_CORTE"):
            val = formato_fecha(val)
        # Campo fecha con hora (FECHA_REGISTRO)
        elif campo == "FECHA_REGISTRO":
            val = _fecha_registro(val)
        # Campos hora HH:MM
        elif campo in ("hora_inicio", "hora_termino"):
            val = _iso_a_hhmm(val)
        # Tiempos en minutos → H:MM
        elif campo in ("tiempo_efectivo", "horas_colacion"):
            val = _minutos_a_hmm(val)
        # RUTs sin guión
        elif campo in ("rut_empresa", "rut_calibrador"):
            val = str(val).replace("-", "") if val else ""
        # Decimales con coma (largo, largo_real ya vienen así desde API
        # pero por seguridad convierto floats)
        elif campo in ("m3ssc", "m3ssc_cabezal", "largo"):
            val = formato_numero(val)
        # Resto: string sin transformar
        else:
            val = str(val) if val is not None else ""
        valores.append(val)
    return ";".join(valores)


def registro_tp_a_csv(rec: dict, indice: int) -> str:
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
    if reporte == "BN":
        nombre = "Base2NOC.csv"
        headers = BN_HEADERS
        parse_fn = lambda rec, i: registro_bn_a_csv(rec)
    else:
        nombre = "TiemposPerdidos.csv"
        headers = TP_HEADERS
        parse_fn = lambda rec, i: registro_tp_a_csv(rec, i)

    filepath = destino / nombre
    with open(filepath, "w", encoding="utf-8-sig", newline="") as f:
        f.write(headers + "\r\n")
        for i, rec in enumerate(datos, 1):
            f.write(parse_fn(rec, i) + "\r\n")

    size_kb = filepath.stat().st_size / 1024
    log.info(f"  💾 {nombre}: {len(datos)} registros, {size_kb:.1f} KB")
    return filepath


# ══════════════════════════════════════════════════════════════════════════════
# FLUJO PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
def main():
    log.info("=" * 60)
    log.info(f"Descarga NOC vía API — {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    hoy = datetime.now()
    primer_dia = hoy.replace(day=1)
    fecha_ini = primer_dia.strftime("%Y-%m-%d")
    fecha_fin = hoy.strftime("%Y-%m-%d")
    log.info(f"📅 Rango: {fecha_ini} → {fecha_fin}")

    for fname in ["Base2NOC.csv", "TiemposPerdidos.csv", "ProductividadGenerico.csv"]:
        fpath = BASE_DIR / fname
        if fpath.exists():
            fpath.unlink()
            log.info(f"  🗑️  Eliminado: {fname}")

    session = requests.Session()
    session.verify = False
    session.headers.update({
        "User-Agent": ("Mozilla/5.0 (X11; Linux x86_64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/131.0.0.0 Safari/537.36"),
    })

    token = obtener_token_arcgis(session)

    datos_bn = descargar_reporte(session, token, "BN", fecha_ini, fecha_fin)
    bn_ok = False
    if datos_bn:
        guardar_csv(datos_bn, "BN", BASE_DIR)
        bn_ok = True
    else:
        log.warning("⚠️  Sin datos de Base 2 NOC")

    datos_tp = descargar_reporte(session, token, "TP", fecha_ini, fecha_fin)
    tp_ok = False
    if datos_tp:
        guardar_csv(datos_tp, "TP", BASE_DIR)
        tp_ok = True
    else:
        log.warning("⚠️  Sin datos de Tiempos Perdidos")

    if bn_ok and tp_ok:
        log.info("✅ Ambos archivos descargados correctamente")
    elif bn_ok or tp_ok:
        log.info("⚠️  Solo se descargó un archivo")
    else:
        log.error("❌ No se descargó ningún archivo")
        sys.exit(1)

    log.info("🎯 Proceso finalizado")
    log.info("=" * 60)


if __name__ == "__main__":
    main()
