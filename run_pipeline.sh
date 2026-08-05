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

# ── 2.72 VMA diario al CMMS — es el factor con que la app de terreno convierte ──
# El jefe de faena ya no estima m³ en /t/avance: CUENTA árboles y viajes, y el volumen sale de
# m³ = árboles × VMA. Ese VMA vive acá (reporte PG del NOC) y el CMMS no lo sabe, así que hay
# que empujárselo. La RPC además COMPLETA los conteos que el jefe declaró antes de que llegara
# el VMA del día —que es el caso normal: él declara al cerrar el turno y esto corre en la
# noche—. NO crítico: sin esto el conteo igual queda guardado y se convierte mañana.
echo ""
echo "▶️  [2.72/9] Publicando el VMA diario por faena al CMMS..."
python3 publicar_vma.py 2>&1 | tee -a "$LOG_PIPELINE" || echo "  ⚠️  VMA no publicado (no crítico, se reintenta en la próxima corrida)"

# ── 2.75 Desplazamiento del GPS (Wialon) — alimenta las filas de km/día del informe ──
# Va ANTES del informe porque el informe lee wialon_km.json. Es incremental: refresca los
# últimos días y acumula sobre el archivo versionado (bajar el mes entero son ~20 min).
# NO crítico: sin token o sin red, se conserva el archivo anterior y el informe muestra "rep.".
echo ""
echo "▶️  [2.75/9] Actualizando desplazamiento GPS (Wialon)..."
timeout 300 python3 descargar_wialon.py 2>&1 | tee -a "$LOG_PIPELINE" || echo "  ⚠️  Wialon no respondió a tiempo (no crítico, se usa lo ya guardado)"

# ── 2.8. Informe de Faena (tablero Arauco + guía productividad, imprimible A4) ──
# Entregable único: consolida el tablero pre-llenado y la guía VMA×especie por faena.
# (generar_tablero_faena.py queda solo como librería: el informe importa sus funciones.)
echo ""
echo "▶️  [2.8/8] Generando Informe de Faena (tablero Arauco, mitad productividad, 1 hoja A4/faena)..."
# REINTENTO: la falla típica acá es que el NOC devuelva basura por un momento. El 31-07-2026
# contestó "1 registros" para el mes entero, el DataFrame quedó sin la columna hora_inicio, el
# generador reventó y —como esto es "no crítico"— el run siguió VERDE dejando en el hosting el
# informe de la corrida anterior. Nadie se entera de eso mirando el ✅ de GitHub.
# Un minuto después el mismo endpoint devolvía los 283 registros: por eso un reintento resuelve
# la mayoría de los casos sin molestar a nadie.
INFORME_FALLO=""
for INTENTO in 1 2; do
    python3 generar_informe_faena.py 2>&1 | tee -a "$LOG_PIPELINE"
    # PIPESTATUS[0] y no $?: con el `tee` de por medio $? es el de tee (siempre 0), así que el
    # `|| echo` que había acá NUNCA se ejecutaba — el informe podía reventar en silencio.
    INFORME_RC=${PIPESTATUS[0]}
    if [ "$INFORME_RC" = "3" ]; then
        # 3 = EXIT_CMMS_AUTH: el CMMS rechazó la credencial (401/403). Esto NO es transitorio y
        # NO es "no crítico": el informe saldría sin la mitad del contenido. Se corta el
        # pipeline para que el run quede en rojo, en vez de publicar un PDF mutilado con un ✅.
        echo ""
        echo "  ❌ Informe de faena: el CMMS rechazó la credencial. Se aborta el pipeline."
        echo "     Arreglar SUPABASE_KEY / los permisos de las RPC informe_* y volver a correr."
        exit 1
    fi
    [ "$INFORME_RC" = "0" ] && break
    if [ "$INTENTO" = "1" ]; then
        echo "  ⚠️  Informe de faena falló (código $INFORME_RC) — reintentando en 45 s…"
        sleep 45
    else
        INFORME_FALLO="código $INFORME_RC en dos intentos"
        echo "  ❌ Informe de faena falló dos veces. El pipeline sigue, pero se avisa por Telegram."
    fi
done

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
# timeout de 4 min: cuando decide regenerar, descarga mes por mes del NOC y ha dejado el
# pipeline colgado 10+ min (cancelado a mano dos veces el 2026-07-25). Es un paso opcional —
# si no alcanza, se retoma en la corrida siguiente.
timeout 240 python3 generar_snapshots.py --todos 2>&1 | tee -a "$LOG_PIPELINE" || echo "  ⚠️  generar_snapshots no terminó a tiempo o falló (no crítico, sigue)"

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

# Las 3 TABLAS + UN solo resumen, el general (gerencia 2026-07-25): los resúmenes de aéreo y
# terrestre repetían lo mismo desglosado y llenaban el chat. Sus imágenes sí se mantienen.
enviar_grupo "Cosecha Forestal (General)" "grilla_produccion.png"           "resumen_diario.txt"
enviar_grupo "Millalemu Aéreo"            "grilla_produccion_aereo.png"     ""
enviar_grupo "Millalemu Terrestre"        "grilla_produccion_terrestre.png" ""

# ── 8.5. Informes de faena en PDF a Telegram, agrupados por ZONA ─────────────
# 2 PDFs (Aéreo / Terrestre) con sus 4 hojas A4 cada uno — una por faena. Agrupados por
# zona a pedido de gerencia: 8 archivos sueltos llenaban el chat y eran incómodos de
# reenviar. Cada hoja es lo que el jefe imprime y llena en terreno.
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
    # Último día con datos del NOC, sacado del propio informe. Va en el caption porque el PDF
    # se manda varias veces al día (10 veces el 31-07) y todos los mensajes se ven idénticos:
    # sin fecha ni hora no hay forma de saber cuál es el nuevo, y se termina revisando uno viejo.
    ULT_DIA_NOC=$(grep -oE '<span>Fecha</span><b>[^<]+' Informe_Faena.html 2>/dev/null \
                  | head -1 | sed 's/.*<b>//' | tr -d ' ')
    for f in Informe_Zona_*.html; do
        [ -f "$f" ] || continue
        pdf="${f%.html}_$(date '+%Y-%m-%d_%H%M').pdf"
        # `timeout` + `--virtual-time-budget` son OBLIGATORIOS: sin ellos Chrome se queda
        # esperando recursos (fuentes/red) y cuelga el pipeline entero — pasó con el HTML por
        # zona, 9 min sin terminar hasta cancelar a mano. El budget corta el reloj virtual de
        # la página; el timeout mata el proceso si igual se traba.
        timeout 90 "$CHROME" --headless --disable-gpu --no-sandbox --no-pdf-header-footer \
                  --virtual-time-budget=10000 --run-all-compositor-stages-before-draw \
                  --print-to-pdf="$pdf" "file://$(pwd)/$f" >/dev/null 2>&1 || true
        if [ -s "$pdf" ]; then
            PDF_OK=$((PDF_OK+1))
            if [ -n "$TELEGRAM_BOT_TOKEN" ] && [ -n "$TELEGRAM_CHAT_ID" ]; then
                ZONA="${f#Informe_Zona_}"; ZONA="${ZONA%.html}"
                RESP=$(curl -s -F "chat_id=${TELEGRAM_CHAT_ID}" \
                     -F "document=@${pdf}" \
                     -F "caption=📋 Informes de Faena — ${ZONA}
🗓️ Datos al ${ULT_DIA_NOC:-—} · generado $(date '+%d-%m-%Y %H:%M')" \
                     "https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendDocument")
                tg_check "${ZONA} — PDF" "$RESP"
            fi
        else
            echo "  ❌ $f — no se pudo generar el PDF"
        fi
    done
    echo "  📄 $PDF_OK PDF(s) generados"

    # ── Gráfico de desplazamiento (GPS) ──
    # Va acá y no en el paso 6 porque se rasteriza con el MISMO Chrome de los PDF, en vez de
    # sumar una segunda librería de gráficos al pipeline. Es el único indicador del GPS que no
    # aparece en ninguna otra salida: la grilla ya manda acumulado, meta, proyección y brecha.
    python3 generar_grafico_desplazamiento.py 2>&1 | tee -a "$LOG_PIPELINE"
    if [ -f grafico_desplazamiento.html ]; then
        timeout 60 "$CHROME" --headless --disable-gpu --no-sandbox --hide-scrollbars \
                  --force-device-scale-factor=2 --window-size=1080,520 \
                  --virtual-time-budget=8000 \
                  --screenshot="grafico_desplazamiento.png" \
                  "file://$(pwd)/grafico_desplazamiento.html" >/dev/null 2>&1 || true
        if [ -s grafico_desplazamiento.png ] && [ -n "$TELEGRAM_BOT_TOKEN" ] && [ -n "$TELEGRAM_CHAT_ID" ]; then
            RESP=$(curl -s -F "chat_id=${TELEGRAM_CHAT_ID}" \
                 -F "photo=@grafico_desplazamiento.png" \
                 -F "caption=🛰️ Desplazamiento diario por faena — datos al ${ULT_DIA_NOC:-—}" \
                 "https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendPhoto")
            tg_check "gráfico de desplazamiento" "$RESP"
        else
            echo "  ⚠️  gráfico de desplazamiento no se pudo generar (no crítico)"
        fi
    fi
fi

# Aviso al chat cuando el informe NO se generó. Va a Telegram y no solo al log porque es donde
# la gente efectivamente mira: el run puede quedar VERDE (el informe es no crítico) y sin este
# mensaje los jefes simplemente no reciben su PDF, sin ninguna señal de que algo pasó. Lo que
# queda publicado en el hosting es el informe de la corrida ANTERIOR, que parece al día.
if [ -n "$INFORME_FALLO" ] && [ -n "$TELEGRAM_BOT_TOKEN" ] && [ -n "$TELEGRAM_CHAT_ID" ]; then
    AVISO="⚠️ <b>Informe de Faena no se generó hoy</b>
Falló $INFORME_FALLO (se reintentó una vez).
El informe publicado en produccion.millalemu.com es el de la corrida anterior — <b>no está al día</b>.
La causa más común es que el NOC devuelva datos incompletos por un rato; suele resolverse solo en la próxima corrida."
    RESP=$(curl -s -X POST "https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage" \
         -d "chat_id=${TELEGRAM_CHAT_ID}" -d "parse_mode=HTML" --data-urlencode "text=${AVISO}")
    tg_check "aviso de informe no generado" "$RESP"
fi

# ── 9. Cierre ────────────────────────────────────────────────────────────────
# El mensaje final con KPIs se retiró a pedido de gerencia (2026-07-25): por Telegram van
# SOLO las 3 tablas con su resumen (paso 8) y los informes de faena en PDF (paso 8.5).
FECHA_FIN="$(date '+%Y-%m-%d %H:%M:%S')"

echo ""
echo "════════════════════════════════════════════════════════════════"
echo "🎯 Pipeline finalizado — $FECHA_FIN"
echo "════════════════════════════════════════════════════════════════"
