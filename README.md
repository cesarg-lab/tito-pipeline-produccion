# Tito Pipeline Producción — GitHub Actions Bundle

Pipeline diario automatizado que reemplaza la corrida manual de la Mac.

**Flujo:** Arauco GeoNOC → CSVs → Dashboard HTML → data.json → FTP hosting → Notificación Telegram.

**Tecnología:** Python + bash, corriendo en **GitHub Actions (gratis, sin límites)**.

> Nota: el archivo `render.yaml` quedó del intento previo con Render. Render eliminó su Cron Job free tier (mínimo US$7/mes), así que cambiamos a GitHub Actions. El `render.yaml` se puede ignorar o borrar — no afecta nada.

---

## 📦 Contenido del bundle

| Archivo | Propósito |
|---|---|
| `.github/workflows/pipeline.yml` | **Workflow GitHub Actions** — schedule cron + trigger manual |
| `descargar_noc_api.py` | Descarga PG + TP desde GeoNOC vía API REST (Bearer token) |
| `normalizar_produccion.py` | Convierte formato Base 2 NOC → PG legacy |
| `GENERAR_HTML.py` | Construye Dashboard_Cosecha.html con todos los KPIs |
| `EXTRAER_JSON.py` | Extrae el objeto `D` del HTML como data.json (para Tito JARVIS) |
| `SUBIR_FTP.py` | Sube HTML + data.json + snapshots al hosting Bluehosting |
| `Dashboard_CosechaForestal.xlsx` | Excel maestro con metas mensuales por faena |
| `historico_cierres_mensuales.csv` | Histórico para tab "meses anteriores" |
| `run_pipeline.sh` | Orquestador bash + notificación Telegram final |
| `requirements.txt` | Dependencias Python |
| `.gitignore` | Excluye credenciales, CSVs generados y logs |

---

## 🚀 Pasos de deploy

### 1. Cargar los 7 secrets en GitHub

En tu repo `tito-pipeline-produccion` → **Settings** → **Secrets and variables** → **Actions** → click **"New repository secret"** y agregar uno por uno:

| Secret name | De dónde sacas el value |
|---|---|
| `ARAUCO_USER` | RUT empresa Forestal Millalemu (formato `XXXXXXXX-X`) |
| `ARAUCO_PASS` | password GeoNOC Arauco |
| `FTP_HOST` | IP servidor Bluehosting (cPanel → Cuentas FTP) |
| `FTP_USER` | `produccion@millalemu.com` |
| `FTP_PASS` | password FTP de cPanel |
| `TELEGRAM_BOT_TOKEN` | token bot `@tito_jarvis_bot` (Make F2 v2 rama Telegram) |
| `TELEGRAM_CHAT_ID` | chat_id numérico de Cesar (lo tienes guardado en `.noc_config.json` y en F2 v2) |

Después de cada uno, click **"Add secret"**. Los valores NO se ven después de guardarlos — GitHub los inyecta como ENV vars al workflow en cada corrida.

### 2. Primera corrida manual

1. En el repo → tab **"Actions"** (arriba)
2. Sidebar izquierdo → click **"Pipeline Producción Diaria"**
3. Botón **"Run workflow"** (arriba derecha del listado) → **"Run workflow"** (en el dropdown que aparece)
4. Refresca la página → aparece una corrida en gris (queued) → amarillo (running) → verde (success) o rojo (failed)
5. Click en la corrida → click en el job **"Descarga + HTML + FTP + Telegram"** → ves el log en vivo

**Validaciones esperadas en el log:**
- ✅ `Token obtenido — expira: ...`
- ✅ `Productividad Genérico: N registros`
- ✅ `Tiempos Perdidos: N registros`
- ✅ `Dashboard subido exitosamente`
- ✅ `Notificación Telegram enviada`

Después: abrí http://produccion.millalemu.com y verificá fecha "Actualizado" arriba. Revisa Telegram, debe estar la notif "✅ Pipeline Producción OK".

### 3. El cron ya queda activo automáticamente

`schedule: "0 2 * * *"` = 02:00 UTC todos los días = **23:00 Chile horario verano** (22:00 invierno).
Mañana corre solo a esa hora, sin tocar nada.

### 4. Apagar pipeline Mac

Una vez confirmada al menos 1 corrida automática exitosa (próximo día):
- Cesar puede dejar de correr `actualizar_y_enviar.sh` manualmente
- La Mac queda libre del compromiso diario

---

## 🔧 Troubleshooting

**El workflow no aparece en "Actions"**
→ Refrescá. GitHub a veces tarda 30 seg en detectar workflows nuevos.

**HTTP 401 desde Arauco**
→ Verificá que `ARAUCO_USER` y `ARAUCO_PASS` estén correctos. Token Arauco se renueva cada 14 días automáticamente.

**FTP timeout**
→ Bluehosting puede tener firewall por IP. GitHub Actions usa IPs públicas dinámicas. Si pasa, levantá ticket a Bluehosting para abrir el FTP a 0.0.0.0/0 (más simple) o configurá whitelist por rangos.

**No llega notificación Telegram**
→ Probá el token: `curl https://api.telegram.org/bot<TOKEN>/getMe` debe devolver info del bot.

**Quiero cambiar la hora**
→ Editá `.github/workflows/pipeline.yml`, cambiá `cron: '0 2 * * *'` (formato cron UTC), commit + push. GitHub detecta el cambio en la próxima ventana.

**Quiero correr manual fuera de hora**
→ Tab Actions → "Pipeline Producción Diaria" → "Run workflow" → "Run workflow". Corre en 1-2 min.

---

## 📊 Logs y monitoreo

- **GitHub Actions UI** → tab Actions → cada corrida tiene log completo (queda 90 días)
- **Telegram** → notif "✅ OK" al final de cada corrida (o "❌ FALLÓ" con link al log si revienta)
- **produccion.millalemu.com** → si la fecha "Actualizado" coincide con hoy a las 23:xx, todo OK
- **Artefactos en caso de falla** → GitHub guarda automáticamente los logs (`descarga_log.txt`, `pipeline_*.log`) 7 días, descargables desde la página del run fallido

---

## 💰 Costos

**GitHub Actions free tier:**
- Workflows en repos **públicos**: **ILIMITADOS minutos, gratis para siempre**
- Workflows en repos privados: 2000 min/mes gratis (sobra para 1 corrida/día de ~3 min = 90 min/mes)

Tu repo es público → 0 costo, sin límite.

---

## 🆘 Volver atrás

Si algo se rompe:
1. **Desactivá el workflow:** Actions tab → "Pipeline Producción Diaria" → menú "..." → "Disable workflow"
2. Volvé a correr el `actualizar_y_enviar.sh` en la Mac mientras tanto
3. Avisame para diagnosticar

El bundle no toca el flujo Mac ni Tito JARVIS — sólo agrega una corrida cloud paralela. Mientras esté activo el workflow GitHub Actions, NO correr la Mac (se pisarían los uploads FTP).

---

_Última actualización: 2026-05-22 — GitHub Actions setup._
