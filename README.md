# SKF Report Generator

Desktop software to attach an Excel test request and generate:

- Word report (`.docx`)
- PDF report (`.pdf`)

using the predefined SKF project specification template.

## UI Flow

- Pitch-dark black UI with slightly shiny black center attachment box
- Metadata row above attachment box:
  - `Date` (calendar picker)
  - `Revision` (numeric)
  - `Revision Date` (calendar picker; blank keeps `DD/MM/YYYY` in report)
  - `Project No.` (example: `TR26-0002-BTS`)
  - `Project Leader` (free-text name)
  - `Tooling Lead Time` dropdown:
    - `Available`, `1 Week` ... `10 Weeks`
- Center large box with:
  - Pin symbol (`📌`) and `Attach your excel` text initially
  - Drag-and-drop support and click-to-browse support
  - After attach: Excel symbol (`📊`) and selected file name
- Curved green action buttons with black text:
  - `Generate Report in Word`
  - `Generate Report in PDF`
  - `Exit`
- Top menu bar:
  - `File` -> `Reset`, `Exit`
- `Help` -> `Download Work Instruction (PDF)`

These metadata values are written to the generated Word/PDF report headers:
- Date -> `Original Date`
- Revision -> `Revision No.`
- Revision Date -> `Revision Date` (or `DD/MM/YYYY` if empty)
- Project No. -> replaces value next to `Project#` in header text
- Project No. -> also updates right-top header text `Project Specification: <ProjectNo>` on report pages
- Project Leader -> updates `Project Leader` value in project details table
- Tooling Lead Time -> `Tooling Design, Manufacture and Inspection` duration row

Additional report logic:
- Testing row text is auto-built from Excel values:
  - `Testing of <N> samples @ <Duration Hrs. each>`
- Testing duration is auto-calculated:
  - Total hours = `samples * duration hours`
  - If up to 4 weeks -> shown in weeks
  - If more than 4 weeks -> shown in months
- Setup notes from Excel setup notes section are inserted under Test Specification (after Fluid section).
- Setup notes and inserted Reciprocation lines use the same report font style as surrounding Test Specification content.
- Contamination section is shown only when contamination values exist; otherwise no contamination/slurry subheading is generated.
- Pre/Post and On-Test snapshots preserve sheet formatting colors and are captured to the last filled On-Test row before the Removal section.
- Product Drawing image is auto-filled from embedded main drawing in Excel `Part Drawing 1`.
- Every generated report includes the `Decision Rule` heading and image block (Choice 2 image) without font/color changes.

Generated files are automatically saved into the user's `Downloads` folder.

## Extraction Rules Implemented

- Pulls requester/customer/application/purpose/number of samples from `Page 1`.
- Pulls technical values for shaft and bore: diameter, tolerance, material, Ra, shaft hardness.
- Pulls setup parameters: DRO, STBM, seal cock, reciprocation, oil type, pre-lube/setup notes.
- For Mud Slurry / Dry Dust / contamination tests, also maps contamination section (type, mix ratio, amount, recip. frequency, STBM orientation).
- Captures pre/post measurements block image from `Page 1!A30:L44`.
- Captures duty-cycle image from `Page 2!A8:O33` (to last filled row).
- Pulls oil change interval (`Page 2!E4`) and acceptance criteria duration/failure.
- Keeps fixed Word content unchanged:
  - `Tolerance for Speed is ± 50 RPM and Temperature is ± 5°C`
  - `Procedure for monitoring`
  - `Disclaimer`
  - `Expected outcome from the Project`

## Template File

Bundled template:

- `assets/Project Specification - Template.docx`
- `assets/Project Specification - Decision Rule Source.docx` (Decision Rule heading/image source)

## Run from Source

Install dependencies:

```bash
python3 -m pip install -r requirements.txt
```

Start GUI:

```bash
python3 app/report_generator_app.py
```

Optional CLI conversion:

```bash
python3 tools/excel_to_word_converter.py --excel <input.xlsm> --template <template.docx> --output <output.docx>
```

## Build Windows `.exe`

Run on a Windows machine:

```bat
build_windows_exe.bat
```

This generates:

- `dist/SKF_Report_Generator.exe`

Notes:

- For reliable PDF generation in Windows `.exe`, Microsoft Word is recommended (used via automation).
- The `.exe` includes the bundled template file.
