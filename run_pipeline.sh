#!/usr/bin/env bash
# ══════════════════════════════════════════════════════════════════════════════
# run_pipeline.sh — Pipeline diario de producción Forestal Millalemu
# ══════════════════════════════════════════════════════════════════════════════
# Orden:
#   1. descargar_noc_api.py  — baja ProductividadGenerico.csv + TiemposPerdidos.csv
#   2. GENERAR_HTML.py        — produce Dashboard_Cosecha.html
#   3. EXTRAER_JSON.py        — extrae data.json del HTML
#   4. SUBIR_FTP.py           — sube todo al hosting
#   5. curl Telegram          — notifica resumen al canal de Cesar
#
# ENV vars requeridas (configurar en Render Dashboard → Environment):
#   ARAUCO_USER, ARAUCO_PASS
#   FTP_HOST, FTP_USER, FTP_PASS
#   TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
# ══════════════════════════════════════════════════════════════════════════════

set -e  # abortar al primer error

cd "$(dirname "$0")"

FECHA_INICIO="$(date '+%Y-%m-%d %H:%M:%S')"
LOG_PIPELINE="pipeline_$(date '+%Y%m%d_%H%M').log"

echo "════════════════════════════════════════════════════════════════"
echo "🚀 Pipeline Producción Millalemu — $FECHA_INICIO"
echo "════════════════════════════════════════════════════════════════"

# ── 1. Descarga GeoNOC ────────────────────────────────────────────────────
echo ""
echo "▶️  [1/4] Descargando NOC desde Arauco..."
python3 descargar_noc_api.py 2>&1 | tee -a "$LOG_PIPELINE"

# ── 2. Generar HTML ───────────────────────────────────────────────────────
echo ""
echo "▶️  [2/4] Generando Dashboard HTML..."
python3 GENERAR_HTML.py 2>&1 | tee -a "$LOG_PIPELINE"

# ── 3. Extraer JSON ───────────────────────────────────────────────────────
echo ""
echo "▶️  [3/4] Extrayendo data.json..."
python3 EXTRAER_JSON.py 2>&1 | tee -a "$LOG_PIPELINE"

# ── 4. Subir FTP ──────────────────────────────────────────────────────────
echo ""
echo "▶️  [4/4] Subiendo a produccion.millalemu.com..."
python3 SUBIR_FTP.py 2>&1 | tee -a "$LOG_PIPELINE"

# ── 5. Notificación Telegram ──────────────────────────────────────────────
FECHA_FIN="$(date '+%Y-%m-%d %H:%M:%S')"
echo ""
echo "▶️  [5/5] Notificando a Telegram..."

# Construir mensaje (extrae info de data.json si existe)
if [ -f "data.json" ] && command -v python3 >/dev/null 2>&1; then
    RESUMEN=$(python3 -c "
import json
with open('data.json') as f:
    d = json.load(f)
mes = d.get('mes', {})
tot = d.get('totales', {})
print(f\"📊 <b>Producción {mes.get('nombre','')} {mes.get('anio','')}</b>\")
print(f\"\")
print(f\"• Acumulado: {tot.get('acumulado_m3',0):,.0f} m³\")
print(f\"• Meta:      {tot.get('meta_m3',0):,.0f} m³\")
print(f\"• Cumplim:   {tot.get('cumplimiento_pct',0):.1f}%\")
print(f\"• Proyec:    {tot.get('proyeccion_fin_mes_m3',0):,.0f} m³ ({tot.get('proyeccion_pct',0):.1f}%)\")
print(f\"• Día {mes.get('dias_con_datos','?')} de {mes.get('dias_mes','?')}\")
" 2>/dev/null || echo "📊 Pipeline ejecutado correctamente")
else
    RESUMEN="📊 Pipeline ejecutado correctamente"
fi

MENSAJE="✅ <b>Pipeline Producción OK</b>
🕐 ${FECHA_INICIO} → ${FECHA_FIN}

${RESUMEN}

🔗 http://produccion.millalemu.com"

if [ -n "$TELEGRAM_BOT_TOKEN" ] && [ -n "$TELEGRAM_CHAT_ID" ]; then
    curl -s -X POST "https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage" \
        -d "chat_id=${TELEGRAM_CHAT_ID}" \
        -d "parse_mode=HTML" \
        --data-urlencode "text=${MENSAJE}" \
        > /dev/null && echo "✅ Notificación Telegram enviada"
else
    echo "⚠️  TELEGRAM_BOT_TOKEN o TELEGRAM_CHAT_ID no definidos — sin notificación"
fi

echo ""
echo "════════════════════════════════════════════════════════════════"
echo "🎯 Pipeline finalizado — $FECHA_FIN"
echo "════════════════════════════════════════════════════════════════"
