# DocuSend

**English** | [Español](#español)

---

## English

A Windows desktop application that automates the generation of Word/PDF documents from an Excel spreadsheet and optionally sends them by email through Outlook — one document per row, in seconds.

### Features

- **Wizard interface** — 4-step guided workflow: Files → Configuration → Email → Generate
- **Word template rendering** — uses `{{ColumnName}}` placeholders anywhere in the `.docx` file
- **Excel data source** — reads any `.xlsx` file; each row produces one document
- **PDF or Word output** — converts to PDF via Microsoft Word (no extra tools needed) or keeps `.docx`
- **Digital signature** — sign PDFs with a PFX/P12 certificate or with [AutoFirma](https://firmaelectronica.gob.es/Home/Descargas.html) (Spanish government tool, Windows certificate store)
- **Email automation via Outlook** — attach the generated document and send or save as draft
  - Choose the recipient from an Excel column
  - Add extra recipients (To, CC, BCC) — either fixed addresses or `{{Column}}` placeholders
  - Select which Outlook account to send from (supports multiple accounts)
  - Use `{{ColumnName}}` in subject and body to personalise each email
  - Optionally use an Outlook `.oft` template instead of writing subject/body manually
- **"Generate only" mode** — produce all documents without sending any emails
- **Auto-updater** — checks for new releases on GitHub at startup; downloads and installs automatically
- **Persistent settings** — remembers all your configuration between sessions

### Requirements

- Windows 10 / 11
- Microsoft Word installed (for PDF conversion)
- Microsoft Outlook installed (only if you want to send/save emails)
- AutoFirma installed (only if you want to sign with the Windows certificate store)

No Python installation required — just download the `.exe` from [Releases](../../releases).

### Quick Start

1. **Download** `docusend.exe` from the latest [Release](../../releases/latest) and run it from a **local folder** (not OneDrive or a network drive — see note below).
2. **Step 1 — Files**: select your Word template (`.docx`), your Excel data file (`.xlsx`), and the output folder.
3. **Step 2 — Configuration**: set the file naming pattern using `{{ColumnName}}` placeholders. Optionally enable digital signing.
4. **Step 3 — Email**: configure subject, body, recipients, Outlook account, output format, and send mode. Set to *"Solo generar archivos"* if you don't want to send emails.
5. **Step 4 — Generate**: click **✓ Generar** and watch the progress log.

> **Important — run from a local folder.**
> If you run the `.exe` from a OneDrive-synced folder (e.g. `OneDrive\Downloads`), OneDrive may lock the temporary DLL files that the app extracts on startup, causing a *"Failed to load Python DLL"* error.
> Move `docusend.exe` to a local folder such as `C:\Users\YourName\AppData\Local\DocuSend\` or your Desktop (if it is not OneDrive-synced).

### Word Template

Use `{{ColumnName}}` anywhere in the `.docx` file to insert data from the Excel row:

```
Dear {{Name}} {{Surname}},

Your contract start date is {{StartDate}}.
```

The column names must match exactly (case-sensitive) the headers in your Excel file.

### Excel File

Any `.xlsx` file works. The first row must contain the column headers. Each subsequent row generates one document (and one email, if configured).

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

To compile:

```bash
pyinstaller --noconfirm --onefile --windowed --collect-all customtkinter --runtime-tmpdir "%TEMP%" --icon assets/icon.ico docusend.py
```

---

## Español

Una aplicación de escritorio para Windows que automatiza la generación de documentos Word/PDF a partir de una hoja de Excel y los envía por correo a través de Outlook — un documento por fila, en segundos.

### Funcionalidades

- **Interfaz tipo asistente** — flujo de trabajo en 4 pasos: Archivos → Configuración → Correo → Generar
- **Plantillas Word con variables** — usa `{{NombreColumna}}` en cualquier parte del `.docx`
- **Fuente de datos Excel** — lee cualquier `.xlsx`; cada fila genera un documento
- **Salida en PDF o Word** — convierte a PDF mediante Microsoft Word (sin herramientas adicionales) o guarda en `.docx`
- **Firma digital** — firma los PDFs con un certificado PFX/P12 o con [AutoFirma](https://firmaelectronica.gob.es/Home/Descargas.html) (almacén de certificados de Windows)
- **Automatización de correo con Outlook** — adjunta el documento generado y envía o guarda como borrador
  - Elige el destinatario desde una columna del Excel
  - Añade destinatarios adicionales (Para, CC, CCO) — direcciones fijas o `{{Columna}}`
  - Selecciona la cuenta de Outlook desde la que enviar (compatible con múltiples cuentas)
  - Usa `{{NombreColumna}}` en asunto y cuerpo para personalizar cada correo
  - Opcionalmente usa una plantilla Outlook `.oft` en lugar de escribir asunto/cuerpo
- **Modo "Solo generar"** — genera todos los documentos sin enviar ningún correo
- **Actualizador automático** — comprueba nuevas versiones en GitHub al arrancar; descarga e instala automáticamente
- **Configuración persistente** — recuerda todos los ajustes entre sesiones

### Requisitos

- Windows 10 / 11
- Microsoft Word instalado (para la conversión a PDF)
- Microsoft Outlook instalado (solo si quieres enviar o guardar borradores)
- AutoFirma instalado (solo si quieres firmar con el almacén de certificados de Windows)

No necesitas instalar Python — descarga el `.exe` desde [Releases](../../releases).

### Inicio rápido

1. **Descarga** `docusend.exe` desde la última [Release](../../releases/latest) y ejecútalo desde una **carpeta local** (no OneDrive ni unidad de red — ver nota abajo).
2. **Paso 1 — Archivos**: selecciona tu plantilla Word (`.docx`), el archivo de datos Excel (`.xlsx`) y la carpeta de salida.
3. **Paso 2 — Configuración**: define el patrón de nombre de archivo usando `{{NombreColumna}}`. Opcionalmente activa la firma digital.
4. **Paso 3 — Correo**: configura asunto, cuerpo, destinatarios, cuenta de Outlook, formato de salida y modo de envío. Elige *"Solo generar archivos"* si no quieres enviar correos.
5. **Paso 4 — Generar**: pulsa **✓ Generar** y observa el log de progreso.

> **Importante — ejecuta desde una carpeta local.**
> Si ejecutas el `.exe` desde una carpeta sincronizada con OneDrive (p. ej. `OneDrive\Descargas`), OneDrive puede bloquear los archivos DLL temporales que la aplicación extrae al arrancar, provocando el error *"Failed to load Python DLL"*.
> Mueve `docusend.exe` a una carpeta local como `C:\Users\TuNombre\AppData\Local\DocuSend\` o al Escritorio (si no está sincronizado con OneDrive).

### Plantilla Word

Usa `{{NombreColumna}}` en cualquier parte del `.docx` para insertar datos de la fila del Excel:

```
Estimado/a {{Nombre}} {{Apellidos}}:

La fecha de inicio de tu contrato es el {{FechaInicio}}.
```

Los nombres de columna deben coincidir exactamente (distingue mayúsculas/minúsculas) con las cabeceras del Excel.

### Archivo Excel

Vale cualquier `.xlsx`. La primera fila debe tener las cabeceras de columna. Cada fila siguiente genera un documento (y un correo, si está configurado).

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

Para compilar:

```bash
pyinstaller --noconfirm --onefile --windowed --collect-all customtkinter --runtime-tmpdir "%TEMP%" --icon assets/icon.ico docusend.py
```
