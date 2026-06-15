# report-i-need

Bot en GitHub Actions para:

1. abrir `https://www.ineedtours.com/es/tours.html`,
2. iniciar sesión con `USER_INEED` y `PASS_INEED`,
3. grabar video del flujo completo,
4. guardar evidencia en Artifacts,
5. registrar una fila de ejecución en Google Sheets usando `CREDENCIALES_JSON`.

## Secretos requeridos

En el repositorio (`Settings` -> `Secrets and variables` -> `Actions`) crea:

- `USER_INEED`: email de acceso a iNeedTours.
- `PASS_INEED`: contraseña del usuario.
- `CREDENCIALES_JSON`: JSON completo de una Service Account de Google con permisos de edición sobre la hoja.

## Hoja de cálculo

Actualmente el script apunta al spreadsheet:

- `https://docs.google.com/spreadsheets/d/1wBZKRRFBJZAUWdPsGa9hX3Ifak8zjKktaTdHFsx3Rms/edit`

Y escribe en:

- `Hoja 1`, rango `A:E` (un log por corrida).

Si quieres cambiar la pestaña de destino, modifica `SHEET_NAME` en `scripts/login-and-record.js`.

## Workflow

Archivo: `.github/workflows/login-report-bot.yml`

Se ejecuta:

- manualmente con `workflow_dispatch`.
- diariamente (cron `0 10 * * *`).

Sube artifacts en cada corrida (incluyendo fallos):

- `artifacts/videos/login-session.webm`
- `artifacts/screenshots/post-login.png` (o `login-error.png`)
- `artifacts/run-metadata.json`
- `artifacts/storage-state.json`

## Correr localmente

```bash
npm install
npx playwright install chromium
```

En PowerShell:

```powershell
$env:USER_INEED="tu-email"
$env:PASS_INEED="tu-password"
$env:CREDENCIALES_JSON='{"type":"service_account", ... }'
npm run bot:login
```

## Próximo paso

Esta primera versión ya valida login + evidencia en video + log en Sheets.

Luego podemos añadir la parte de:

- descargar el reporte exacto,
- parsear el archivo,
- pegar el contenido estructurado en la hoja objetivo.