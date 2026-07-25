#!/usr/bin/env bash
# ══════════════════════════════════════════════════════════════════════════════
# run_pipeline.sh — Pipeline diario de producción Forestal Millalemu
# ══════════════════════════════════════════════════════════════════════════════
# Orden:
#   1. descargar_noc_api.py   — baja ProductividadGenerico.csv + TiemposPerdidos.csv
#   2. cp → Base2NOC.csv      — alias para GENERAR_IMAGEN/RESUMEN
#   3. GENERAR_HTML.py        — produce Dashboard_Cosecha.html
#   4. EXTRAER_JSON.py        — extrae data.json del HTML
#   5. SUBIR_FTP.py           — sube todo al hosting
#   6. GENERAR_IMAGEN.py ×3   — grilla_produccion(_aereo|_terrestre).png
#   7. GENERAR_RESUMEN.py ×3  — resumen_diario(_aereo|_terrestre).txt
#   8. sendPhoto Telegram ×3  — imagen + resumen a chat tito_jarvis_bot
#   9. sendMessage Telegram   — notificación final con KPIs del mes
#
# ENV vars requeridas (configurar en GitHub Settings → Secrets and variables → Actions):
#   ARAUCO_USER, ARAUCO_PASS
#   FTP_HOST, FTP_USER, FTP_PASS
#   TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
# ══════════════════════════════════════════════════════════════════════════════

set -e  # abortar al primer error en pasos críticos (1-5)

cd "$(dirname "$0")"

FECHA_INICIO="$(date '+%Y-%m-%d %H:%M:%S')"
LOG_PIPELINE="pipeline_$(date '+%Y%m%d_%H%M').log"

echo "════════════════════════════════════════════════════════════════"
echo "🚀 Pipeline Producción Millalemu — $FECHA_INICIO"
echo "════════════════════════════════════════════════════════════════"

# ── 1. Descarga GeoNOC ────────────────────────────────────────────────────
echo ""
echo "▶️  [1/9] Descargando NOC desde Arauco..."
python3 descargar_noc_api.py 2>&1 | tee -a "$LOG_PIPELINE"

# ── 2. Verificar Base2NOC.csv (descargado por descargar_noc_api.py paso 1) ──
echo ""
echo "▶️  [2/9] Verificando Base2NOC.csv..."
if [ -f "Base2NOC.csv" ]; then
    SIZE=$(wc -c < Base2NOC.csv)
    echo "✅ Base2NOC.csv presente ($SIZE bytes)"
else
    echo "❌ Base2NOC.csv NO encontrado — paso 1 falló"
    exit 1
fi

# ── 2.5 Archivar mes anterior si falta (idempotente; recupera el mes cerrado) ──
echo ""
echo "▶️  [2.5/9] Archivando mes anterior en el histórico si falta..."
python3 archivar_mes_anterior.py 2>&1 | tee -a "$LOG_PIPELINE" || echo "  ⚠️  archivar_mes_anterior falló (no crítico, sigue)"

# ── 2.7 KPIs Uso/Ritmo/Carga/VMA + su dashboard (no crítico) ──────────────
echo ""
echo "▶️  [2.7/9] Calculando KPIs Uso/Ritmo/Carga/VMA y generando Dashboard_KPIs.html..."
( python3 compute_kpis.py && python3 generar_dashboard_kpis.py kpis.json tm_por_faena.json Dashboard_KPIs.html ) 2>&1 | tee -a "$LOG_PIPELINE" || echo "  ⚠️  KPIs fallaron (no crítico, el pipeline sigue)"

# ── 2.8. Informe de Faena (tablero Arauco + guía productividad, imprimible A4) ──
# Entregable único: consolida el tablero pre-llenado y la guía VMA×especie por faena.
# (generar_tablero_faena.py queda solo como librería: el informe importa sus funciones.)
echo ""
echo "▶️  [2.8/8] Generando Informe de Faena (tablero Arauco, mitad productividad, 1 hoja A4/faena)..."
python3 generar_informe_faena.py 2>&1 | tee -a "$LOG_PIPELINE" || echo "  ⚠️  Informe de faena falló (no crítico, el pipeline sigue)"

# ── 3. Generar HTML ───────────────────────────────────────────────────────
echo ""
echo "▶️  [3/9] Generando Dashboard HTML..."
python3 GENERAR_HTML.py 2>&1 | tee -a "$LOG_PIPELINE"

# ── 4. Extraer JSON ───────────────────────────────────────────────────────
echo ""
echo "▶️  [4/9] Extrayendo data.json..."
python3 EXTRAER_JSON.py 2>&1 | tee -a "$LOG_PIPELINE"

# ── 5. Subir FTP ──────────────────────────────────────────────────────────
echo ""
echo "▶️  [5/9] Subiendo a produccion.millalemu.com..."
python3 SUBIR_FTP.py 2>&1 | tee -a "$LOG_PIPELINE"

# ── 5.5 Generar snapshots de meses cerrados (idempotente; solo los que faltan) ──
echo ""
echo "▶️  [5.5/9] Generando snapshots de meses pasados si faltan..."
python3 generar_snapshots.py --todos 2>&1 | tee -a "$LOG_PIPELINE" || echo "  ⚠️  generar_snapshots falló (no crítico, sigue)"

# ── Pasos siguientes son nice-to-have: no abortan el pipeline si fallan ──
set +e

# ── 6. Generar 3 imágenes ─────────────────────────────────────────────────
echo ""
echo "▶️  [6/9] Generando 3 imágenes (general / aéreo / terrestre)..."
python3 GENERAR_IMAGEN.py                    2>&1 | tee -a "$LOG_PIPELINE"
python3 GENERAR_IMAGEN.py --grupo aereo      2>&1 | tee -a "$LOG_PIPELINE"
python3 GENERAR_IMAGEN.py --grupo terrestre  2>&1 | tee -a "$LOG_PIPELINE"

# ── 7. Generar 3 resúmenes ────────────────────────────────────────────────
echo ""
echo "▶️  [7/9] Generando 3 resúmenes de texto..."
python3 GENERAR_RESUMEN.py                    2>&1 | tee -a "$LOG_PIPELINE"
python3 GENERAR_RESUMEN.py --grupo aereo      2>&1 | tee -a "$LOG_PIPELINE"
python3 GENERAR_RESUMEN.py --grupo terrestre  2>&1 | tee -a "$LOG_PIPELINE"

# ── 8. Enviar imágenes + resúmenes a Telegram (Cesar copy/paste a WhatsApp) ──
echo ""
echo "▶️  [8/9] Enviando 3 imágenes + resúmenes a Telegram..."

# Diagnóstico: ¿el bot y el chat existen? (sin esto, un token o chat_id malo pasaba mudo)
if [ -n "$TELEGRAM_BOT_TOKEN" ]; then
    ME=$(curl -s "https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/getMe")
    if printf '%s' "$ME" | grep -q '"ok":true'; then
        BOT=$(printf '%s' "$ME" | sed -n 's/.*"username":"\([^"]*\)".*/\1/p')
        echo "  🤖 bot @${BOT} OK · chat_id=${TELEGRAM_CHAT_ID}"
    else
        echo "  ❌ TELEGRAM_BOT_TOKEN inválido — Telegram: $(printf '%s' "$ME" | head -c 200)"
    fi
else
    echo "  ⚠️  TELEGRAM_BOT_TOKEN vacío"
fi

# Telegram responde HTTP 200 incluso cuando falla ("ok":false + description). curl devuelve 0
# igual, así que el `&&` marcaba TODO como enviado aunque no llegara nada. Esta función mira
# la respuesta de verdad y muestra el motivo.
tg_check() {
    local QUE="$1" RESP="$2"
    if printf '%s' "$RESP" | grep -q '"ok":true'; then
        echo "  ✅ $QUE"
    else
        local MOTIVO
        MOTIVO=$(printf '%s' "$RESP" | sed -n 's/.*"description":"\([^"]*\)".*/\1/p')
        echo "  ❌ $QUE — Telegram rechazó: ${MOTIVO:-respuesta vacía o sin red}"
    fi
}

enviar_grupo() {
    local TITULO="$1"
    local IMG="$2"
    local TXT="$3"

    if [ -z "$TELEGRAM_BOT_TOKEN" ] || [ -z "$TELEGRAM_CHAT_ID" ]; then
        echo "  ⚠️  $TITULO — sin credenciales Telegram, skip"
        return
    fi

    # Enviar foto con caption del título
    if [ -f "$IMG" ]; then
        RESP=$(curl -s -F "chat_id=${TELEGRAM_CHAT_ID}" \
             -F "photo=@${IMG}" \
             -F "caption=📊 ${TITULO}" \
             "https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendPhoto")
        tg_check "$TITULO — imagen" "$RESP"
    else
        echo "  ❌ $TITULO — imagen $IMG no existe"
    fi

    # Enviar texto del resumen (separado para que Cesar pueda copiar/pegar limpio)
    if [ -f "$TXT" ]; then
        # Truncar a 4000 chars por seguridad (limite Telegram 4096)
        local CONTENT
        CONTENT=$(head -c 4000 "$TXT")
        RESP=$(curl -s -X POST "https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage" \
             -d "chat_id=${TELEGRAM_CHAT_ID}" \
             --data-urlencode "text=${CONTENT}")
        tg_check "$TITULO — texto" "$RESP"
    fi
}

enviar_grupo "Cosecha Forestal (General)" "grilla_produccion.png"           "resumen_diario.txt"
enviar_grupo "Millalemu Aéreo"            "grilla_produccion_aereo.png"     "resumen_diario_aereo.txt"
enviar_grupo "Millalemu Terrestre"        "grilla_produccion_terrestre.png" "resumen_diario_terrestre.txt"

# ── 8.5. Informes de faena en PDF a Telegram ─────────────────────────────────
# Una hoja A4 por faena, que es lo que el jefe imprime y llena en terreno.
# Chrome headless ya viene instalado en los runners de GitHub (ubuntu-latest).
echo ""
echo "▶️  [8.5/9] Convirtiendo informes de faena a PDF y enviando a Telegram..."

CHROME=""
for c in google-chrome google-chrome-stable chromium chromium-browser; do
    command -v "$c" >/dev/null 2>&1 && CHROME="$c" && break
done

if [ -z "$CHROME" ]; then
    echo "  ⚠️  Sin Chrome/Chromium en el runner — no se generan PDFs (no crítico)"
else
    PDF_OK=0
    for f in Informe_M*.html; do
        [ -f "$f" ] || continue
        pdf="${f%.html}.pdf"
        "$CHROME" --headless --disable-gpu --no-sandbox --no-pdf-header-footer \
                  --print-to-pdf="$pdf" "file://$(pwd)/$f" >/dev/null 2>&1 || true
        if [ -s "$pdf" ]; then
            PDF_OK=$((PDF_OK+1))
            if [ -n "$TELEGRAM_BOT_TOKEN" ] && [ -n "$TELEGRAM_CHAT_ID" ]; then
                FAENA="${f#Informe_}"; FAENA="${FAENA%.html}"
                FAENA="${FAENA//-/.}"      # M1-1 → M1.1, como se lee en la faena
                RESP=$(curl -s -F "chat_id=${TELEGRAM_CHAT_ID}" \
                     -F "document=@${pdf}" \
                     -F "caption=📋 Informe de Faena ${FAENA}" \
                     "https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendDocument")
                tg_check "${FAENA} — PDF" "$RESP"
            fi
        else
            echo "  ❌ $f — no se pudo generar el PDF"
        fi
    done
    echo "  📄 $PDF_OK PDF(s) generados"
fi

# ── 9. Cierre ────────────────────────────────────────────────────────────────
# El mensaje final con KPIs se retiró a pedido de gerencia (2026-07-25): por Telegram van
# SOLO las 3 tablas con su resumen (paso 8) y los informes de faena en PDF (paso 8.5).
FECHA_FIN="$(date '+%Y-%m-%d %H:%M:%S')"

echo ""
echo "════════════════════════════════════════════════════════════════"
echo "🎯 Pipeline finalizado — $FECHA_FIN"
echo "════════════════════════════════════════════════════════════════"
