# Audit Report Compiler

![Windows](https://img.shields.io/badge/platform-Windows-0078d4?style=flat-square&logo=windows)
![Python](https://img.shields.io/badge/built%20with-Python-3776ab?style=flat-square&logo=python)
![License](https://img.shields.io/badge/license-MIT-green?style=flat-square)
![No install](https://img.shields.io/badge/no%20install-just%20run%20the%20.exe-success?style=flat-square)

A lightweight Windows desktop tool that compiles a professional **Cyber Audit Report** by merging a Word template, an Excel checklist, and custom header/footer images into a single, formatted `.docx` file — all in one click.

<img width="771" height="339" alt="image" src="https://github.com/user-attachments/assets/b5fb2b68-6a4f-44b7-bbed-244ca0543580" />
<img width="1866" height="698" alt="image" src="https://github.com/user-attachments/assets/dbd26cf5-876d-467b-82ad-09c6183734b8" />
<img width="1899" height="780" alt="image" src="https://github.com/user-attachments/assets/cd7aa9d6-3e92-4417-818a-51d9b7d77d97" />
<img width="1902" height="789" alt="image" src="https://github.com/user-attachments/assets/a2423e19-5539-4958-8c86-83faac0abbb1" />

---

## What It Does

The tool takes four inputs and produces one output:

| Input | Description |
|---|---|
| **Word Template** (.docx) | Your existing audit report narrative/template |
| **Excel Checklist** (.xlsx) | The filled audit checklist (e.g. SEBI CSCRF TOR sheet) |
| **Header Image** (.png/.jpg) | Your firm's header banner image |
| **Footer Image** (.png/.jpg) | Your firm's footer banner image |
| **Output File** (.docx) | The compiled final report |

The Excel sheet is automatically embedded as a formatted table inside the Word document, with merged cells, proper column widths, and consistent fonts — ready to share or print.

---

## Getting Started

### Download

1. Go to the [**Releases**](../../releases) page
2. Download `AuditReportCompiler.exe`
3. Double-click to run — **no installation required**

> Windows may show a SmartScreen warning on first run since the exe is unsigned. Click **"More info" → "Run anyway"** to proceed.

---

## How to Use

1. **Launch** `AuditReportCompiler.exe`
2. Click **Browse** next to each field and select your files:
   - Word Template (.docx)
   - Excel Checklist (.xlsx)
   - Header Image (.png)
   - Footer Image (.png)
3. Click **Save As** to choose where the output report should be saved
4. Click **Generate Report**
5. A success message will confirm the report has been saved

---

## Requirements

- **Windows 10 or 11** (64-bit)
- No Python installation needed
- No additional software required — everything is bundled in the `.exe`

---

## Output Formatting

The generated report automatically applies the following formatting:

- **Landscape A4** layout throughout
- **Custom header and footer** images applied to every page
- **Excel table** embedded with borders, merged cells preserved, and column widths proportionally scaled
- **Times New Roman, 9pt** font inside the table
- **0.1cm cell margins** for clean, compact layout
- Trailing blank pages and ghost tables removed automatically

---

## Tips & Known Behaviour

- Make sure the **output `.docx` file is not already open** in Word before generating — you will get a file-in-use error
- Header and footer images should ideally be **full page width** (e.g. 2480px wide) for best results
- Excel column widths in your `.xlsx` file directly influence how wide each column appears in the report — adjust them in Excel before running if needed
- Merged cells in the Excel sheet are preserved in the output table

---

## Repository Contents

```
.
└── AuditReportCompiler.exe    # Standalone Windows executable
└── README.md                  # This file
```

The source code is not included in this repository. The `.exe` is self-contained.

---

## License

MIT — free to use and distribute.

---

## Support

If you encounter a bug or the report doesn't generate as expected, please open an [Issue](../../issues) and include:
- A description of what happened
- The error message shown (if any)
- Your Windows version
