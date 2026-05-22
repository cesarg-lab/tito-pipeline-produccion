#!/usr/bin/env python3
"""
SUBIR_FTP.py — Sube Dashboard HTML + data.json + snapshots al hosting vía FTP
══════════════════════════════════════════════════════════════════════════════
Versión cloud-ready 2026-05-21:
  - Lee credenciales FTP desde ENV vars (FTP_HOST, FTP_USER, FTP_PASS)
  - Sin credenciales hardcoded — seguro para repo público

Uso (Render cron job):
  FTP_HOST=<ip_servidor> FTP_USER=<usuario> FTP_PASS=<password> python3 SUBIR_FTP.py

Uso (dev local):
  Definir las mismas ENV vars en el shell o en un .env (no commitear).
"""

import ftplib
import io
import os
import sys
import glob as _glob
from datetime import datetime

# ── Credenciales desde ENV vars ────────────────────────────────────────────
FTP_HOST = os.environ.get("FTP_HOST")
FTP_USER = os.environ.get("FTP_USER")
FTP_PASS = os.environ.get("FTP_PASS")
FTP_DIR  = os.environ.get("FTP_DIR", "/")  # default = raíz del home FTP

if not (FTP_HOST and FTP_USER and FTP_PASS):
    print("❌ Faltan credenciales FTP: definir ENV vars FTP_HOST, FTP_USER, FTP_PASS")
    sys.exit(1)

# ── Archivos a subir ───────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_FILE = os.path.join(BASE_DIR, "Dashboard_Cosecha.html")
REMOTE_NAME = "index.html"  # se sube como index.html para que abra directo

# ══════════════════════════════════════════════════════════════════════════

def subir_dashboard():
    if not os.path.exists(HTML_FILE):
        print("❌ No se encontró Dashboard_Cosecha.html")
        print("   Ejecuta primero: python3 GENERAR_HTML.py")
        sys.exit(1)

    print(f"📤 Conectando a {FTP_HOST}...")
    try:
        ftp = ftplib.FTP(FTP_HOST, timeout=30)
        ftp.login(FTP_USER, FTP_PASS)
        print(f"✅ Conectado como {FTP_USER}")

        # Verificar/crear directorio remoto
        try:
            ftp.cwd(FTP_DIR)
        except ftplib.error_perm:
            print(f"📁 Creando directorio {FTP_DIR}...")
            dirs = FTP_DIR.strip("/").split("/")
            current = ""
            for d in dirs:
                current += f"/{d}"
                try:
                    ftp.cwd(current)
                except ftplib.error_perm:
                    ftp.mkd(current)
                    ftp.cwd(current)
            print(f"✅ Directorio creado: {FTP_DIR}")

        # .htaccess (UTF-8 + anti-cache)
        htaccess = (
            b"AddDefaultCharset UTF-8\n"
            b"AddCharset UTF-8 .html\n"
            b"\n"
            b"# Anti-cache: forzar que el navegador siempre pida la version mas reciente\n"
            b"<IfModule mod_headers.c>\n"
            b"  Header set Cache-Control \"no-cache, no-store, must-revalidate\"\n"
            b"  Header set Pragma \"no-cache\"\n"
            b"  Header set Expires \"0\"\n"
            b"</IfModule>\n"
            b"<IfModule mod_expires.c>\n"
            b"  ExpiresActive On\n"
            b"  ExpiresByType text/html \"access plus 0 seconds\"\n"
            b"</IfModule>\n"
        )
        ftp.storbinary("STOR .htaccess", io.BytesIO(htaccess))

        # Dashboard principal (mes en curso)
        file_size = os.path.getsize(HTML_FILE) / 1024
        print(f"📤 Subiendo Dashboard ({file_size:.1f} KB) como {REMOTE_NAME}...")
        with open(HTML_FILE, "rb") as f:
            ftp.storbinary(f"STOR {REMOTE_NAME}", f)

        # data.json (resumen optimizado para Tito JARVIS)
        json_path = os.path.join(BASE_DIR, "data.json")
        if os.path.exists(json_path):
            json_size = os.path.getsize(json_path) / 1024
            print(f"📤 Subiendo data.json ({json_size:.1f} KB)...")
            with open(json_path, "rb") as f:
                ftp.storbinary("STOR data.json", f)
            print(f"   📍 JSON disponible en: http://produccion.millalemu.com/data.json")
        else:
            print("⚠️  data.json no existe — ejecuta EXTRAER_JSON.py antes de SUBIR_FTP.py")

        # Snapshots históricos: Dashboard_Cosecha_YYYY-MM.html
        snapshot_htmls = sorted(_glob.glob(os.path.join(BASE_DIR, "Dashboard_Cosecha_2*.html")))
        for sh in snapshot_htmls:
            name = os.path.basename(sh)
            size_kb = os.path.getsize(sh) / 1024
            print(f"📤 Subiendo snapshot {name} ({size_kb:.1f} KB)...")
            with open(sh, "rb") as f:
                ftp.storbinary(f"STOR {name}", f)

        print(f"✅ Dashboard subido exitosamente")
        print(f"   📍 URL: http://produccion.millalemu.com")
        print(f"   🕐 Actualizado: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

        ftp.quit()

    except ftplib.all_errors as e:
        print(f"❌ Error FTP: {e}")
        sys.exit(1)


if __name__ == "__main__":
    subir_dashboard()
