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
        curl -s -F "chat_id=${TELEGRAM_CHAT_ID}" \
             -F "photo=@${IMG}" \
             -F "caption=📊 ${TITULO}" \
             "https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendPhoto" > /dev/null \
             && echo "  ✅ $TITULO — imagen enviada"
    else
        echo "  ❌ $TITULO — imagen $IMG no existe"
    fi

    # Enviar texto del resumen (separado para que Cesar pueda copiar/pegar limpio)
    if [ -f "$TXT" ]; then
        # Truncar a 4000 chars por seguridad (limite Telegram 4096)
        local CONTENT
        CONTENT=$(head -c 4000 "$TXT")
        curl -s -X POST "https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage" \
             -d "chat_id=${TELEGRAM_CHAT_ID}" \
             --data-urlencode "text=${CONTENT}" > /dev/null \
             && echo "  ✅ $TITULO — texto enviado"
    fi
}

enviar_grupo "Cosecha Forestal (General)" "grilla_produccion.png"           "resumen_diario.txt"
enviar_grupo "Millalemu Aéreo"            "grilla_produccion_aereo.png"     "resumen_diario_aereo.txt"
enviar_grupo "Millalemu Terrestre"        "grilla_produccion_terrestre.png" "resumen_diario_terrestre.txt"

# ── 9. Notificación final ─────────────────────────────────────────────────
FECHA_FIN="$(date '+%Y-%m-%d %H:%M:%S')"
echo ""
echo "▶️  [9/9] Notificación final OK..."

if [ -f "data.json" ]; then
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
         --data-urlencode "text=${MENSAJE}" > /dev/null \
         && echo "✅ Notificación final Telegram enviada"
else
    echo "⚠️  Sin credenciales Telegram"
fi

echo ""
echo "════════════════════════════════════════════════════════════════"
echo "🎯 Pipeline finalizado — $FECHA_FIN"
echo "════════════════════════════════════════════════════════════════"
