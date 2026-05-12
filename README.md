# Leaseir SAT dashboard

Página HTML pública con la **evolución de incidencias abiertas SAT** (por estado y por cadena) sobre fechas seleccionadas (6 fin-de-mes + últimos 15 días laborables).

Los datos los toma de Jira, se guardan en caché en este repo (`cache/jira_status_timeline.json`) y se sirven via GitHub Pages.

## Cómo está montado

1. **`scripts/fetch_jira.py`** — pulla el changelog de cada ticket relevante de Jira y lo guarda en la cache. Modo `seed` (primera vez) baja todo. Modo `update` (siguientes) solo baja los que han cambiado desde la última ejecución.
2. **`scripts/build_report.py`** — lee la cache, replay las transiciones para cada cutoff, y construye Excel + HTML.
3. **`.github/workflows/refresh.yml`** — corre diariamente (06:00 UTC) o a demanda. Actualiza cache, regenera HTML, publica en GitHub Pages.

## Setup inicial (one-time)

### 1. Crear API token de Jira

1. Ve a https://id.atlassian.com/manage-profile/security/api-tokens
2. Crea un token nuevo, etiquétalo "leaseir-sat-dashboard"
3. **Cópialo ya** — Atlassian no te lo enseña otra vez

### 2. Crear el repo en GitHub

```bash
# Local
cd "Analista de Jira/sat-dashboard"
git init
git add .
git commit -m "initial commit"

# Crear el repo en GitHub (vacío, sin README) y luego:
git remote add origin git@github.com:<tu-usuario>/leaseir-sat-dashboard.git
git branch -M main
git push -u origin main
```

### 3. Configurar secrets en GitHub

Settings → Secrets and variables → Actions → New repository secret. Añadir:

| Secret | Valor |
|---|---|
| `JIRA_BASE_URL` | `https://leaseir.atlassian.net` |
| `JIRA_USER` | `iortiz@leaseir.com` |
| `JIRA_TOKEN` | el token del paso 1 |

### 4. Habilitar GitHub Pages

Settings → Pages → Source: **GitHub Actions**.

### 5. Lanzar el primer run

Actions → "Refresh SAT dashboard" → Run workflow → Run.

El primer run baja unos 250 changelogs (~30 segundos), construye la cache y publica el HTML. Verás el link en el output del job (`page_url`).

A partir de ahí, cron diario lo refresca solo.

## Ejecutar localmente

```bash
pip install -r requirements.txt
export JIRA_BASE_URL=https://leaseir.atlassian.net
export JIRA_USER=iortiz@leaseir.com
export JIRA_TOKEN=*** # el token

# seed inicial
python scripts/fetch_jira.py --cache cache/jira_status_timeline.json --mode seed

# o update incremental
python scripts/fetch_jira.py --cache cache/jira_status_timeline.json --mode update

# generar Excel + HTML
python scripts/build_report.py \
    --cache cache/jira_status_timeline.json \
    --output-xlsx output/Evolucion_SAT.xlsx \
    --output-html output/index.html
```

## Configurar fechas/cadenas

- **Cutoffs**: `scripts/build_report.py::compute_cutoffs()` — los calcula dinámicamente (6 últimos fin-de-mes + 15 últimos días laborables). Cambiar la función si quieres otro patrón.
- **Cadenas**: `CHAIN_ORDER` y `chain_of()` en `build_report.py`.
- **Estados "abiertos"**: `OPEN_STATUSES` en ambos scripts. Deben coincidir.

## Privacidad

Si quieres que el HTML sea privado, en lugar de GitHub Pages usa:
- Cloudflare Pages con Access (recomendado, fácil)
- Repo privado + servirlo via Vercel con auth de SSO
- Subirlo a un bucket S3 con autenticación

Habla conmigo si quieres montar alguna de estas opciones.
