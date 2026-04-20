# Excel to Word Test Spec Converter

This project converts a filled SKF Excel request file (`.xlsm`) into an editable Word project specification (`.docx`) using a predefined Word template.

## What the tool does

- Reads key administrative and technical values from `Page 1` of Excel (requester, customer, application, purpose, sample count, shaft/bore specs, setup data, contamination data, notes).
- Reads duty-cycle and acceptance data from `Page 2` (oil change interval, duration, failure criteria).
- Captures:
  - Pre/Post measurement block (`Page 1!A30:L44`) as an image.
  - Duty cycle block (`Page 2!A8:O33`) as an image (up to last populated row).
- Places the extracted content into the Word template while keeping fixed sections unchanged (for example monitoring procedure, disclaimer, and tolerance statement text).

## Setup

Install dependencies:

```bash
python3 -m pip install openpyxl python-docx pillow
```

## Usage

```bash
python3 tools/excel_to_word_converter.py \
  --excel "/path/to/input.xlsm" \
  --template "/path/to/template.docx" \
  --output "/path/to/output.docx"
```

Example with your files:

```bash
python3 tools/excel_to_word_converter.py \
  --excel "/Users/anoopnarasimhalu/Downloads/26-01-50x64x6 HMSA10 RG - Mud & Slurry.xlsm" \
  --template "/Users/anoopnarasimhalu/Downloads/Project Specification  - RAE1010 Mud & Slurry Test.docx" \
  --output "output/Project Specification - Generated.docx"
```
