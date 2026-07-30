# Distributing the tool to the team (Windows)

Each person gets one file: `SKF_Report_Generator.exe`. They double-click
it and a normal Windows app window opens.

- No Python to install.
- No browser, no link, no HTML.
- **No host PC.** Nothing runs in the background, nothing stays on.
- No network port is opened. The tool never listens for connections.

Everything happens on the user's own machine, and generated reports go
straight to their own Downloads folder.

## Building the .exe (once, on a Windows PC)

Building is the only step that needs Python, and only on the machine doing
the build - not on anyone else's.

```
build_windows_exe.bat
```

Output: `dist\SKF_Report_Generator.exe`, plus a `.sha256.txt` checksum.

The script refuses to build if either security check fails:

1. **`pip-audit`** - blocks the build if any dependency has a known
   published vulnerability.
2. **`bandit`** - blocks the build on any medium or high severity code
   finding.

Neither is advisory. A failing check stops the build rather than warning
and continuing, so a vulnerable build cannot be produced by accident.

## Code signing - do not skip this

Without a signature, Windows SmartScreen shows your team an **"unknown
publisher"** warning, and corporate antivirus may quarantine the file
outright. This is the single most likely reason an internal `.exe` fails
to reach people.

Ask IT for a code-signing certificate, then:

```
set SIGN_CERT=C:\path\to\cert.pfx
set CERT_PASSWORD=your-password
build_windows_exe.bat
```

The build signs the executable with a timestamp, so it stays valid after
the certificate itself expires. Unsigned builds still work, but the script
prints a loud warning, because handing colleagues a file that trips a
security warning trains them to click through security warnings.

## Distributing it

Put the `.exe` on a shared drive or send it via Teams, **with the
checksum**. Recipients can verify the file matches what you built:

```
certutil -hashfile SKF_Report_Generator.exe SHA256
```

Compare against `SKF_Report_Generator.exe.sha256.txt`. This is what makes
"is this file safe?" answerable instead of a matter of trust.

## What each user needs

- **Microsoft Word** for *Generate Report in PDF* - the PDF is produced by
  driving real Word (`app/report_export.py`), which is why it matches your
  templates exactly. *Generate Report in Word* does not need it.
- **Tesseract OCR**, only for scanned/handwritten monitoring sheets with
  no embedded text layer. A separate install; pip cannot provide it.

## How the data is handled

Relevant if IT asks, and worth knowing before you distribute:

- **Nothing leaves the machine.** No network connections, no telemetry, no
  cloud service, no logging of report contents.
- **Inputs are read from where the user picked them** and are never copied
  anywhere persistent.
- **Intermediate files are auto-deleted.** Generation uses
  `tempfile.TemporaryDirectory()`, so the temp folder is removed even if
  generation fails partway.
- **Output goes to the user's own Downloads folder** and nowhere else.
- **Photo handling is the one place untrusted file parsing happens**
  (users open image files). Pillow is pinned to a security floor of
  12.3.0 for exactly this reason - see the note in `requirements.txt`.

## Trade-offs of the single-file build

- **First launch is slow** - a one-file build unpacks itself to a temp
  folder on every run, and the bundled Word templates alone are ~9 MB.
  Expect a few seconds before the window appears, with no splash screen.
  This is normal and not a hang.
- **Updates mean redistributing the file.** There is no auto-update; when
  you rebuild, everyone needs the new copy. Bump the version in
  `version_info.txt` each release so you can tell which copy someone is
  running.
