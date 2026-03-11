# Preprod Evidence Global Script

This folder contains a **global/portable** evidence generator script, without changing your existing script.

## Folder Contents

- `generate-preprod-evidence-global.ps1`
- `README.md`

## Prerequisites

1. Windows + PowerShell.
2. LibreOffice installed.
   - Recommended: `soffice` is available in PATH, or
   - use the `-SofficePath` parameter.

## Simple Folder Structure

By default, the script uses folders relative to the script location:

- Input default: `./input`  → put **PDF files directly** here
- Output default: `./out`

Contoh:

```text
scripts/preprod-evidence-global/
├─ generate-preprod-evidence-global.ps1
├─ README.md
├─ input/
│  ├─ file1.pdf
│  ├─ file2.pdf
│  └─ ...
└─ out/
```

## Usage

Jalankan dari root repo:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\preprod-evidence-global\generate-preprod-evidence-global.ps1
```

Example: override input/output paths:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\preprod-evidence-global\generate-preprod-evidence-global.ps1 -InputDir "D:\evidence\input" -OutputDir "D:\evidence\output"
```

Example: if `soffice` is not in PATH:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\preprod-evidence-global\generate-preprod-evidence-global.ps1 -SofficePath "C:\Program Files\LibreOffice\program\soffice.exe"
```

Example: overwrite output without timestamp:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\preprod-evidence-global\generate-preprod-evidence-global.ps1 -ForceOverwrite
```

Example: custom output file name:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\preprod-evidence-global\generate-preprod-evidence-global.ps1 -OutputFileName "evidence-custom-name.docx"
```

## Output

DOCX will be generated in the output folder with default file name:

- `UFO - FMC Pre-Deployment Evidence.docx`

If the file already exists and `-ForceOverwrite` is not used, the script automatically appends a timestamp.

## Available Parameters

- `-InputDir` (opsional)
- `-OutputDir` (opsional)
- `-OutputFileName` (optional, exact output file name)
- `-OutputBaseName` (optional, used when `-OutputFileName` is not set)
- `-ReportTitle` (optional)
- `-SofficePath` (optional)
- `-ForceOverwrite` (optional)

## Push to GitHub

```bash
git add scripts/preprod-evidence-global
git commit -m "add global preprod evidence generator and usage docs"
git push
```
