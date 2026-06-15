# report-i-need

Bot en GitHub Actions para:

1. abrir `https://www.ineedtours.com/es/tours.html`,
2. iniciar sesión con `USER_INEED` y `PASS_INEED`,
3. ir a reservas privadas y descargar el Excel,
4. transformar los datos (sin tocar encabezados en la hoja),
5. pegar datos en Google Sheets,
6. grabar video del flujo completo y guardar evidencia en Artifacts.

## Secretos requeridos

En el repositorio (`Settings` -> `Secrets and variables` -> `Actions`) crea:

- `USER_INEED`: email de acceso a iNeedTours.
- `PASS_INEED`: contraseña del usuario.
- `CREDENCIALES_JSON`: JSON completo de una Service Account de Google con permisos de edición sobre la hoja.

## Hoja de cálculo

Actualmente el script apunta al spreadsheet:

- `https://docs.google.com/spreadsheets/d/1wBZKRRFBJZAUWdPsGa9hX3Ifak8zjKktaTdHFsx3Rms/edit`

Y escribe en:

- `Reservas I Need Tours`.
- Encabezados: se conservan como están (no se reemplazan).
- Datos: se limpian desde `A2:ZZ` y se vuelven a cargar desde `A2` (sin copiar encabezado del Excel descargado).

Si quieres cambiar la pestaña de destino, modifica `SHEET_NAME` en `scripts/login-and-record.js`.

## Transformaciones aplicadas

Sobre el Excel descargado de `listado_reservas.aspx`:

1. Se omite la primera fila (encabezados del Excel origen).
2. Columna `H`: se calcula como `F - 15 días` (fecha de pago).
3. Columnas `J` e `K`: se limpia formato monetario (`$ 48` -> `48`, `$ 0` -> `0`).

## Workflow

Archivo: `.github/workflows/login-report-bot.yml`

Se ejecuta:

- manualmente con `workflow_dispatch`.
- diariamente (cron `0 10 * * *`).

Sube artifacts en cada corrida (incluyendo fallos):

- `artifacts/videos/login-session.webm`
- `artifacts/screenshots/post-login.png` (o `login-error.png`)
- `artifacts/downloads/reservas.xlsx`
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

## Estado actual

Ya está implementado:

- login automático,
- descarga de reservas,
- transformación de columnas solicitadas,
- pegado en Sheets respetando encabezados,
- evidencia por video y screenshots.