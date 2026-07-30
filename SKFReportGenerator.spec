# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for SKF Report Generator - macOS .app bundle.

Builds the browser (HTML) tool (app/web_app.py): launching the .app starts
a local web server (waitress) and opens it in the default browser, rather
than opening a Tkinter window. The old Tkinter desktop app still exists in
the repo and runs from source (`python3 app/report_generator_app.py`), but
is no longer what this bundle packages.
"""

import sys
from pathlib import Path

block_cipher = None

# Collect any optional packages only if they are installed
hidden_imports = [
    "openpyxl",
    "docx",
    "docx2pdf",
    "PIL",
    "appscript",
    "lxml",
    "lxml.etree",
    "lxml._elementpath",
    "flask",
    "waitress",
    "jinja2",
]

datas = [
    # Bundle the Word template assets inside the .app
    (str(Path("assets") / "Project Specification - Template.docx"), "assets"),
    (str(Path("assets") / "Project Specification - Decision Rule Source.docx"), "assets"),
    (str(Path("assets") / "Final Test Report - Template.docx"), "assets"),
    # The web app's HTML/CSS/JS - resolved via app.paths.resource_path()
    # the same way the .docx templates above are, so it works the same
    # way inside this frozen bundle's extraction dir.
    (str(Path("app") / "templates"), str(Path("app") / "templates")),
    (str(Path("app") / "static"), str(Path("app") / "static")),
]

try:
    import tkinterdnd2
    tkdnd_path = Path(tkinterdnd2.__file__).parent
    datas.append((str(tkdnd_path), "tkinterdnd2"))
    hidden_imports.append("tkinterdnd2")
except ImportError:
    pass

try:
    import tkcalendar
    hidden_imports.append("tkcalendar")
    hidden_imports.append("babel")
    hidden_imports.append("babel.numbers")
except ImportError:
    pass

a = Analysis(
    [str(Path("app") / "web_app.py")],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="SKFReportGenerator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="SKFReportGenerator",
)

app = BUNDLE(
    coll,
    name="SKF Report Generator.app",
    icon=None,
    bundle_identifier="com.skf.reportgenerator",
    info_plist={
        "CFBundleName": "SKF Report Generator",
        "CFBundleDisplayName": "SKF Report Generator",
        "CFBundleVersion": "1.0.0",
        "CFBundleShortVersionString": "1.0.0",
        "NSHighResolutionCapable": True,
        "NSRequiresAquaSystemAppearance": False,
        "CFBundleDocumentTypes": [
            {
                "CFBundleTypeName": "Excel Spreadsheet",
                "CFBundleTypeExtensions": ["xlsx", "xlsm"],
                "CFBundleTypeRole": "Viewer",
            }
        ],
    },
)
