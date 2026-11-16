# DocuSend

**English** | [Español](#español)

---

## English

A Windows desktop application that automates the generation of Word/PDF documents from an Excel spreadsheet and optionally sends them by email through Outlook — one document per row, in seconds.

### Features

- **Wizard interface** — 4-step guided workflow: Files → Configuration → Email → Generate
- **Word template rendering** — uses `{{ColumnName}}` placeholders anywhere in the `.docx` file
- **Excel data source** — reads any `.xlsx` file; each row produces one document. Choose the sheet and, optionally, a row range (e.g. rows 20 to 40)
- **PDF or Word output** — converts to PDF via Microsoft Word (no extra tools needed) or keeps `.docx`
- **Digital signature** — sign PDFs with a PFX/P12 certificate or with [AutoFirma](https://firmaelectronica.gob.es/Home/Descargas.html) (Spanish government tool, Windows certificate store)
- **Email automation via Outlook** — attach the generated document and send or save as draft
  - Choose the recipient from an Excel column
  - Add extra recipients (To, CC, BCC) — either fixed addresses or `{{Column}}` placeholders
  - Select which Outlook account to send from (supports multiple accounts)
  - Use `{{ColumnName}}` in subject and body to personalise each email
  - Optionally use an Outlook `.oft` template instead of writing subject/body manually
- **Static email attachments** — attach one or more fixed files (e.g. terms & conditions, brochures) to every outgoing email, on top of the generated document
- **"Generate only" mode** — produce all documents without sending any emails
- **Review before you run** — the last step summarises what is about to happen (rows, format, signature, send mode, recipients) and asks for confirmation before sending directly
- **Trial run** — *"Probar con la primera fila"* generates just the first row and opens the document, or shows the email in Outlook without saving or sending it
- **Safety checks** — fields that don't match any Excel column, rows with a missing or malformed email, and certificate problems are reported *before* anything is generated. An email is never sent if its document could not be signed or converted, or if an attachment is missing: it is left in Drafts instead
- **Result report** — every run saves a CSV with the outcome of each row, plus the full log, in a `_DocuSend_informes` subfolder of the output folder
- **Appearance mode** — switch between System, Light, and Dark themes from the footer; preference is saved between sessions
- **Auto-updater** — checks for new releases on GitHub at startup; asks before downloading and replacing the executable
- **Persistent settings** — remembers all your configuration between sessions

### Requirements

- Windows 10 / 11
- Microsoft Word installed (for PDF conversion)
- Microsoft Outlook installed (only if you want to send/save emails)
- AutoFirma installed (only if you want to sign with the Windows certificate store)

No Python installation required — just download the `.exe` from [Releases](../../releases).

### Quick Start

1. **Download** `docusend.exe` from the latest [Release](../../releases/latest) and run it from a **local folder** (not OneDrive or a network drive — see note below).
2. **Step 1 — Files**: select your Word template (`.docx`), your Excel data file (`.xlsx`), and the output folder. Pick the sheet and, if you want, a row range. Optionally add static attachments that will be included in every email.
3. **Step 2 — Configuration**: set the file naming pattern using `{{ColumnName}}` placeholders. Optionally enable digital signing (with AutoFirma you can load your certificates and pick one, so you are not asked for every document).
4. **Step 3 — Email**: configure subject, body, recipients, Outlook account, output format, and send mode. Set to *"Solo generar archivos"* if you don't want to send emails.
5. **Step 4 — Generate**: check the summary, use **Probar con la primera fila** to try one row, then click **✓ Generar** and watch the progress log. When it finishes, **Ver informe** opens the per-row report.

> **Important — run from a local folder.**
> If you run the `.exe` from a OneDrive-synced folder (e.g. `OneDrive\Downloads`), OneDrive may lock the temporary DLL files that the app extracts on startup, causing a *"Failed to load Python DLL"* error.
> Move `docusend.exe` to a local folder such as `C:\Users\YourName\AppData\Local\DocuSend\` or your Desktop (if it is not OneDrive-synced).

### Word Template

Use `{{ColumnName}}` anywhere in the `.docx` file to insert data from the Excel row:

```
Dear {{Name}} {{Surname}},

Your contract start date is {{StartDate}}.
```

The column names must match exactly (case-sensitive) the headers in your Excel file. Names with spaces or hyphens work too (`{{Full name}}`, `{{Start-date}}`). If the template uses a field that is not a column in the Excel file, DocuSend tells you which one and does not start.

### Excel File

Any `.xlsx` file works. The first row must contain the column headers. Each subsequent row generates one document (and one email, if configured).

Dates are written as `dd/mm/yyyy` and whole numbers without decimals. To keep leading zeros (postcodes, phone numbers), format those cells as **Text** in Excel.

| Name   | Surname | StartDate  | Email                 |
|--------|---------|------------|-----------------------|
| María  | García  | 01/09/2026 | maria@example.com     |
| Carlos | López   | 15/09/2026 | carlos@example.com    |

### Email Placeholders

In the email subject and body you can also use `{{ColumnName}}`:

```
Subject: Contract attached — {{Name}} {{Surname}}

Hi {{Name}},
Please find your contract for the position starting on {{StartDate}}.
```

### Building from Source

```bash
git clone https://github.com/aruizciee/docusend.git
cd docusend
pip install -r requirements.txt
python docusend.py
```

To compile (or just run `build.bat`):

```bash
pyinstaller --noconfirm --onefile --windowed --collect-all customtkinter --runtime-tmpdir "%TEMP%" --icon assets/icon.ico --name DocuSend_App docusend.py
```

### Tests

```bash
pip install -r requirements.txt -r requirements-dev.txt
pytest
```

The tests run the real application against a fake Outlook, so nothing is ever sent. They run automatically on every push to `main`.

### Publishing a release

Pushing to `main` does **not** publish anything. To release a new version to users (the app's auto-updater will offer it), either:

- open **Actions → Build y Release → Run workflow** on GitHub, or
- push a tag in the format `vYYYY.MM.DD-N`, e.g. `git tag v2026.11.16-1 && git push origin v2026.11.16-1`

---

## Español

Una aplicación de escritorio para Windows que automatiza la generación de documentos Word/PDF a partir de una hoja de Excel y los envía por correo a través de Outlook — un documento por fila, en segundos.

### Funcionalidades

- **Interfaz tipo asistente** — flujo de trabajo en 4 pasos: Archivos → Configuración → Correo → Generar
- **Plantillas Word con variables** — usa `{{NombreColumna}}` en cualquier parte del `.docx`
- **Fuente de datos Excel** — lee cualquier `.xlsx`; cada fila genera un documento. Puedes elegir la hoja y, si quieres, un rango de filas (p. ej. de la 20 a la 40)
- **Salida en PDF o Word** — convierte a PDF mediante Microsoft Word (sin herramientas adicionales) o guarda en `.docx`
- **Firma digital** — firma los PDFs con un certificado PFX/P12 o con [AutoFirma](https://firmaelectronica.gob.es/Home/Descargas.html) (almacén de certificados de Windows)
- **Automatización de correo con Outlook** — adjunta el documento generado y envía o guarda como borrador
  - Elige el destinatario desde una columna del Excel
  - Añade destinatarios adicionales (Para, CC, CCO) — direcciones fijas o `{{Columna}}`
  - Selecciona la cuenta de Outlook desde la que enviar (compatible con múltiples cuentas)
  - Usa `{{NombreColumna}}` en asunto y cuerpo para personalizar cada correo
  - Opcionalmente usa una plantilla Outlook `.oft` en lugar de escribir asunto/cuerpo
- **Adjuntos estáticos adicionales** — adjunta uno o varios archivos fijos (p. ej. condiciones generales, folletos) a todos los correos, además del documento generado
- **Modo "Solo generar"** — genera todos los documentos sin enviar ningún correo
- **Resumen antes de generar** — el último paso muestra lo que va a ocurrir (filas, formato, firma, modo de envío, destinatarios) y pide confirmación antes de enviar directamente
- **Prueba con la primera fila** — genera solo esa fila y abre el documento, o muestra el correo en Outlook sin guardarlo ni enviarlo
- **Comprobaciones de seguridad** — los campos que no coinciden con ninguna columna del Excel, las filas sin email o con el email mal escrito y los problemas con el certificado se avisan *antes* de generar nada. Nunca se envía un correo cuyo documento no se pudo firmar o convertir, o al que le falta un adjunto: se deja en Borradores
- **Informe de resultados** — cada ejecución guarda un CSV con el resultado de cada fila, además del log completo, en la subcarpeta `_DocuSend_informes` de la carpeta de salida
- **Modo de apariencia** — cambia entre los temas Sistema, Claro y Oscuro desde el pie de la ventana; la preferencia se guarda entre sesiones
- **Actualizador automático** — comprueba nuevas versiones en GitHub al arrancar; pregunta antes de descargar y sustituir el ejecutable
- **Configuración persistente** — recuerda todos los ajustes entre sesiones

### Requisitos

- Windows 10 / 11
- Microsoft Word instalado (para la conversión a PDF)
- Microsoft Outlook instalado (solo si quieres enviar o guardar borradores)
- AutoFirma instalado (solo si quieres firmar con el almacén de certificados de Windows)

No necesitas instalar Python — descarga el `.exe` desde [Releases](../../releases).

### Inicio rápido

1. **Descarga** `docusend.exe` desde la última [Release](../../releases/latest) y ejecútalo desde una **carpeta local** (no OneDrive ni unidad de red — ver nota abajo).
2. **Paso 1 — Archivos**: selecciona tu plantilla Word (`.docx`), el archivo de datos Excel (`.xlsx`) y la carpeta de salida. Elige la hoja y, si quieres, un rango de filas. Opcionalmente añade adjuntos estáticos que se incluirán en todos los correos.
3. **Paso 2 — Configuración**: define el patrón de nombre de archivo usando `{{NombreColumna}}`. Opcionalmente activa la firma digital (con AutoFirma puedes cargar tus certificados y elegir uno, para que no te lo pida en cada documento).
4. **Paso 3 — Correo**: configura asunto, cuerpo, destinatarios, cuenta de Outlook, formato de salida y modo de envío. Elige *"Solo generar archivos"* si no quieres enviar correos.
5. **Paso 4 — Generar**: revisa el resumen, usa **Probar con la primera fila** para probar una fila, y después pulsa **✓ Generar** y observa el log de progreso. Al terminar, **Ver informe** abre el informe fila a fila.

> **Importante — ejecuta desde una carpeta local.**
> Si ejecutas el `.exe` desde una carpeta sincronizada con OneDrive (p. ej. `OneDrive\Descargas`), OneDrive puede bloquear los archivos DLL temporales que la aplicación extrae al arrancar, provocando el error *"Failed to load Python DLL"*.
> Mueve `docusend.exe` a una carpeta local como `C:\Users\TuNombre\AppData\Local\DocuSend\` o al Escritorio (si no está sincronizado con OneDrive).

### Plantilla Word

Usa `{{NombreColumna}}` en cualquier parte del `.docx` para insertar datos de la fila del Excel:

```
Estimado/a {{Nombre}} {{Apellidos}}:

La fecha de inicio de tu contrato es el {{FechaInicio}}.
```

Los nombres de columna deben coincidir exactamente (distingue mayúsculas/minúsculas) con las cabeceras del Excel. También funcionan los nombres con espacios o guiones (`{{Nombre completo}}`, `{{Fecha-inicio}}`). Si la plantilla usa un campo que no es una columna del Excel, DocuSend te dice cuál y no empieza.

### Archivo Excel

Vale cualquier `.xlsx`. La primera fila debe tener las cabeceras de columna. Cada fila siguiente genera un documento (y un correo, si está configurado).

Las fechas se escriben como `dd/mm/aaaa` y los números enteros sin decimales. Para conservar los ceros iniciales (códigos postales, teléfonos), pon esas celdas con formato **Texto** en Excel.

| Nombre | Apellidos | FechaInicio | Email                  |
|--------|-----------|-------------|------------------------|
| María  | García    | 01/09/2026  | maria@ejemplo.com      |
| Carlos | López     | 15/09/2026  | carlos@ejemplo.com     |

### Variables en el correo

En el asunto y el cuerpo del correo también puedes usar `{{NombreColumna}}`:

```
Asunto: Contrato adjunto — {{Nombre}} {{Apellidos}}

Hola {{Nombre}},
Adjunto encontrarás tu contrato para el puesto con inicio el {{FechaInicio}}.
```

### Compilar desde el código fuente

```bash
git clone https://github.com/aruizciee/docusend.git
cd docusend
pip install -r requirements.txt
python docusend.py
```

Para compilar (o simplemente ejecuta `build.bat`):

```bash
pyinstaller --noconfirm --onefile --windowed --collect-all customtkinter --runtime-tmpdir "%TEMP%" --icon assets/icon.ico --name DocuSend_App docusend.py
```

### Tests

```bash
pip install -r requirements.txt -r requirements-dev.txt
pytest
```

Los tests ejecutan la aplicación real contra un Outlook simulado, así que nunca se envía nada. Se ejecutan automáticamente en cada push a `main`.

### Publicar una versión

Hacer push a `main` **no** publica nada. Para publicar una versión nueva a los usuarios (el actualizador de la app se la ofrecerá), hay dos opciones:

- en GitHub, abrir **Actions → Build y Release → Run workflow**, o
- subir una etiqueta con el formato `vAAAA.MM.DD-N`, por ejemplo `git tag v2026.11.16-1 && git push origin v2026.11.16-1`
