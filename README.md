# Tito Pipeline Producción — Render Cron Bundle

Pipeline diario automatizado que reemplaza la corrida manual de la Mac.

**Flujo:** Arauco GeoNOC → CSVs → Dashboard HTML → data.json → FTP hosting → Notificación Telegram.

**Tecnología:** Python + bash, corriendo en Render Cron Job (free tier, $0/mes).

---

## 📦 Contenido del bundle

| Archivo | Propósito |
|---|---|
| `descargar_noc_api.py` | Descarga PG + TP desde GeoNOC vía API REST (Bearer token) |
| `normalizar_produccion.py` | Convierte formato Base 2 NOC → PG legacy |
| `GENERAR_HTML.py` | Construye Dashboard_Cosecha.html con todos los KPIs |
| `EXTRAER_JSON.py` | Extrae el objeto `D` del HTML como data.json (para Tito JARVIS) |
| `SUBIR_FTP.py` | Sube HTML + data.json + snapshots al hosting Bluehosting |
| `Dashboard_CosechaForestal.xlsx` | Excel maestro con metas mensuales por faena |
| `historico_cierres_mensuales.csv` | Histórico para tab "meses anteriores" |
| `run_pipeline.sh` | Orquestador bash + notificación Telegram final |
| `requirements.txt` | Dependencias Python |
| `render.yaml` | Blueprint Render Cron Job (schedule + ENV vars) |
| `.gitignore` | Excluye credenciales, CSVs generados y logs |

---

## 🚀 Pasos de deploy (orden exacto)

### 1. Crear repo GitHub público

1. Andate a https://github.com/new
2. Nombre sugerido: `tito-pipeline-produccion`
3. Visibility: **Public** (Render free necesita repo público)
4. NO inicializar con README/license (ya lo tienes acá)
5. Click "Create repository"

### 2. Subir bundle al repo

Lo más simple sin tocar Git: usar la web de GitHub con drag-and-drop.

1. En el repo recién creado, click "uploading an existing file"
2. Arrastra TODOS los archivos de esta carpeta `render_bundle/`
3. Commit message: "Initial commit — bundle pipeline producción"
4. Click "Commit changes"

**Verificación:** revisa que NO aparezcan:
- `ProductividadGenerico.csv`, `TiemposPerdidos.csv` (data sensible)
- `.env`, `.noc_config.json` (credenciales)
- `Dashboard_Cosecha.html`, `data.json` (data sensible)
- Logs

Si aparecen, parar y avisar a Tito antes de seguir.

### 3. Conectar Render al repo

1. Cuenta en https://render.com (gratis, login con GitHub recomendado)
2. Dashboard → "New +" → "Blueprint"
3. Conectar el repo `tito-pipeline-produccion`
4. Render detecta `render.yaml` automáticamente → click "Apply"
5. Se crea un servicio "tito-pipeline-produccion" tipo Cron

### 4. Configurar ENV vars (7 secretos)

En el dashboard del servicio → tab **Environment** → agregar las 7 keys siguientes.

> **⚠️ Los valores reales NO están en este repo.** Tito te los pasa por chat aparte (o los tienes en `.noc_config.json` local + cPanel de Bluehosting + bot F2 v2 de Make).

| Key | Dónde sacar el value |
|---|---|
| `ARAUCO_USER` | RUT empresa Forestal Millalemu (formato `XXXXXXXX-X`) |
| `ARAUCO_PASS` | password GeoNOC Arauco (la que usas en el browser) |
| `FTP_HOST` | IP del servidor Bluehosting (cPanel → Cuentas FTP) |
| `FTP_USER` | usuario FTP completo `produccion@millalemu.com` |
| `FTP_PASS` | password FTP de cPanel |
| `TELEGRAM_BOT_TOKEN` | token bot `@tito_jarvis_bot` (Make F2 v2 rama Telegram) |
| `TELEGRAM_CHAT_ID` | chat_id de Cesar |

Save Changes.

### 5. Primera corrida manual

1. Dashboard del servicio → botón **"Trigger Run"** (arriba a la derecha)
2. Esperar 1-3 min, mirar el log en vivo
3. Validaciones esperadas:
   - ✅ `Token obtenido — expira: ...`
   - ✅ `Productividad Genérico: N registros`
   - ✅ `Tiempos Perdidos: N registros`
   - ✅ `Dashboard subido exitosamente`
   - ✅ `Notificación Telegram enviada`
4. Abrir http://produccion.millalemu.com y verificar fecha "Actualizado" arriba
5. Telegram: revisar que llegó el mensaje "✅ Pipeline Producción OK"

### 6. Activar cron diario

Si el paso 5 salió bien:
- El cron ya queda activado automáticamente con `schedule: "0 2 * * *"` (02:00 UTC = 23:00 Chile horario verano)
- Mañana a las 23:00 corre solo, sin tocar nada

### 7. Apagar pipeline Mac

Una vez confirmada al menos 1 corrida automática exitosa:
- Cesar puede dejar de correr `actualizar_y_enviar.sh` manualmente
- La Mac queda libre del compromiso diario

---

## 🔧 Troubleshooting

**El cron no aparece después de "Apply" del blueprint**
→ Refresca la página del Dashboard. A veces Render tarda 30 seg.

**ENV vars no se aplican**
→ Después de cambiar una ENV var, tienes que volver a hacer "Trigger Run" — no se reaplican automáticamente al cron pendiente.

**HTTP 401 desde Arauco**
→ Token expirado (válido 14 días pero ya no debería caducar acá). Revisa que `ARAUCO_USER` y `ARAUCO_PASS` estén bien.

**FTP timeout**
→ Bluehosting puede tener firewall por IP. Si pasa, levantar ticket a Bluehosting con la IP de Render (la verás en el log de error).

**No llega notificación Telegram**
→ Verifica `TELEGRAM_BOT_TOKEN` y `TELEGRAM_CHAT_ID`. Probar el token con: `curl https://api.telegram.org/bot<TOKEN>/getMe`.

**Quiero correr en otra hora**
→ Edita `render.yaml`, cambia `schedule: "0 2 * * *"` (formato cron UTC), commit + push. Render detecta y reaplica.

---

## 📊 Logs y monitoreo

- **Render Dashboard** → Logs (live + historial completo de cada corrida)
- **Telegram** → notificación al final de cada corrida (✅ OK o ❌ Falla con detalle)
- **produccion.millalemu.com** → si la fecha de "Actualizado" coincide con hoy a las 23:xx, todo OK

---

## 🆘 Volver atrás

Si algo se rompe en producción:
1. Apagar el cron en Render (Dashboard → Suspend)
2. Volver a correr el `actualizar_y_enviar.sh` en la Mac mientras tanto
3. Avisar a Tito para diagnosticar

El bundle no toca nada del flujo Mac ni de Tito JARVIS — solo agrega una corrida cloud paralela. Mientras esté activo el cron Render, NO correr la Mac (se pisarían).

---

_Última actualización: 2026-05-21 — Sprint Render bundle completo._
