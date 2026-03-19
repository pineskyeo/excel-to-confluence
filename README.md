# Excel to Confluence Converter

Convert Excel (`.xlsx`) files into Word (`.docx`) documents ready for import into Confluence.

Supports **tables**, **images**, **merged cells**, and **text** across multiple sheets.

---

## How It Works

```
Excel (.xlsx)
  └── Sheet 1  ──►  Heading + Text + Tables + Images
  └── Sheet 2  ──►  Heading + Text + Tables + Images
  └── Sheet 3  ──►  ...
                          │
                          ▼
                    output.docx
                          │
                          ▼
              Confluence Word Import
```

Each sheet becomes a section in the Word document, preserving:
- Cell text and font styles (bold, italic, color)
- Table structures including merged cells and background colors
- Embedded images, anchored in their original row order

---

## Requirements

- Python 3.9+
- pip packages listed in `requirements.txt`

---

## Installation

```bash
git clone https://github.com/pineskyeo/excel-to-confluence.git
cd excel-to-confluence

pip install -r requirements.txt
```

---

## Usage

### 1. Convert your Excel file

```bash
python converter.py input.xlsx
```

Output file is saved as `input.docx` in the same directory.

To specify a custom output path:

```bash
python converter.py input.xlsx output.docx
```

### 2. Import into Confluence

1. Open a Confluence page and click **Edit**
2. Click the **···** (more options) menu
3. Select **Import Word Document**
4. Upload the generated `.docx` file

---

## Quick Test with Sample File

Generate a sample Excel file that includes text, tables, merged cells, and images:

```bash
python create_sample.py
```

This creates `sample.xlsx` with 3 sheets:

| Sheet | Contents |
|-------|----------|
| Project Overview | Title text + image + task table |
| Technical Specs | Merged cell headers + image + 2 tables |
| Summary | KPI text + image + summary table |

Then convert it:

```bash
python converter.py sample.xlsx
```

---

## Project Structure

```
excel-to-confluence/
├── converter.py        # Main converter
├── create_sample.py    # Sample Excel generator (for testing)
├── requirements.txt    # Python dependencies
└── README.md
```

---

## Supported Excel Elements

| Element | Supported |
|---------|-----------|
| Text cells | ✅ |
| Bold / Italic font | ✅ |
| Font color | ✅ |
| Cell background color | ✅ |
| Tables (named) | ✅ |
| Tables (auto-detected) | ✅ |
| Merged cells | ✅ |
| Embedded images (PNG/JPEG) | ✅ |
| Multiple sheets | ✅ |
| Formulas | ✅ (computed value only) |
| Charts | ❌ (skipped) |
