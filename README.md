# email-parser

A C# console application for Windows that reads emails from a specified Outlook folder and saves each email — together with all its attachments — as a single PDF file in your **My Documents** folder.

## Requirements

- Windows 10 (version 1607+) or Windows 11
- [.NET 8 Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/8.0) (x64)
- **Outlook mode only:** Microsoft Outlook (Office 365 / Microsoft 365) installed and configured with a profile

## Features

- Reads emails from any Outlook folder **or from a local directory of .msg files** (no Office required)
- Saves each email as a single merged PDF containing:
  - A formatted header (From, To, Subject, Received date)
  - The full email body (HTML or plain-text)
  - All attachments converted and appended as additional pages
- Supported attachment types:
  | Type | Extensions | Processing |
  |------|-----------|------------|
  | Images | `.jpg`, `.jpeg`, `.png`, `.bmp`, `.gif`, `.tiff` | Embedded as PDF pages |
  | PDF | `.pdf` | Merged directly |
  | Word | `.doc`, `.docx`, `.rtf` | Converted via Word (Office 365) |
  | Excel | `.xls`, `.xlsx`, `.csv` | Converted via Excel (Office 365) |
  | ZIP | `.zip` | Extracted; each contained file processed recursively |
- Output files are saved to `My Documents\EmailParser\<FolderName>\<Subject>.pdf`
- Each email's attachments are saved in their **original format** in a folder named `<subject> attachments` (e.g. `re: window enquiry attachments\`) placed alongside the PDF — one folder per email thread
- ZIP attachments are extracted into a named subfolder inside the email's attachment folder
- Input folder structures are fully preserved in the output — subdirectories in the source directory are mirrored in the output
- Duplicate subject lines are handled by appending a counter, e.g. `Report (2).pdf`

## Building

```bash
cd src/EmailParser
dotnet build -c Release
```

The compiled binary will be at `src/EmailParser/bin/Release/net8.0-windows/EmailParser.exe`.

## Usage

The tool operates in two modes depending on what you pass as the source:

| What you pass | Mode | Requires Office? |
|---|---|---|
| Outlook folder name (e.g. `Inbox`) | Outlook COM interop | Yes |
| Path to a local directory (e.g. `C:\exports\emails`) | .msg file reader | No |

### Option 1 — Outlook folder (requires Microsoft Office)

```cmd
EmailParser.exe "Inbox"
EmailParser.exe "Inbox/Projects"
EmailParser.exe "Archive/2024/Q1"
```

### Option 2 — Local directory of .msg files (no Office required)

Export your emails from Outlook as .msg files (drag them to a folder, or use
File → Save As), then point the tool at that directory:

```cmd
EmailParser.exe "C:\Users\you\exports\inbox"
EmailParser.exe ".\my_emails"
```

The tool auto-detects which mode to use: if the path you supply is an existing
directory it reads `.msg` files from it; otherwise it connects to Outlook.

### Option 3 — Interactive prompt

Run `EmailParser.exe` without arguments and enter the source when prompted:

```
Email Parser — Save Outlook Emails to PDF
==========================================

Enter an Outlook folder name (e.g. 'Inbox' or 'Inbox/Projects'),
or a path to a local directory containing .msg files: C:\exports\inbox
```

### Folder path format (Outlook mode)

- Use the **Outlook display name** of the folder (case-insensitive).
- Separate nested levels with `/` or `\`, e.g. `Inbox/Projects/Active`.
- Common special-folder names such as `Inbox`, `Sent Items`, `Drafts`, `Deleted Items`, `Junk Email`, and `Outbox` are resolved automatically.

## Running without Microsoft Office

If Microsoft Office is not installed you can still convert emails to PDF by
supplying a directory of `.msg` files:

1. On a machine that **does** have Outlook, select the emails you want, then
   drag-and-drop them into a folder on disk (Windows saves each as a `.msg` file),
   or use **File → Save As** in Outlook.
2. Copy that folder to the machine where you want to run the tool.
3. Run `EmailParser.exe "C:\path\to\folder"`.

Word and Excel attachments still require Microsoft Office to be converted; if
Office is absent they are skipped with a warning and the rest of the email is
still saved as PDF.

## Output structure

```
My Documents\
  EmailParser\
    inbox\
      Meeting agenda.pdf
      Meeting agenda attachments\   ← "<subject> attachments" folder
        presentation.pptx
        data\                       ← extracted from data.zip
          report.docx
      Project status update.pdf
                                    ← no folder created (no attachments)
      subfolder\                    ← mirrors input subfolder structure
        Q1 Budget Review.pdf
        Q1 Budget Review attachments\
          budget.xlsx
```

## Notes

- Word and Excel conversions launch the respective Office application silently in the background and close it afterwards.
- Unsupported attachment types (e.g. `.exe`, `.mp4`) are skipped with a warning.
- Nested ZIP files are extracted and processed recursively.
- If an individual email fails to convert, the error is reported and processing continues with the remaining emails.
- When running in `.msg` directory mode the tool searches **all subdirectories** recursively and replicates the source folder hierarchy in the output directory.
- Each email's attachments are saved in their original format in a folder named `<subject> attachments` (e.g. `re: window enquiry attachments\`) placed next to the PDF.  Each email thread therefore gets its own clearly-labelled attachment folder.  ZIP attachments are extracted into a named subfolder inside that folder.
