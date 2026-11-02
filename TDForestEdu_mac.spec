# -*- mode: python ; coding: utf-8 -*-
# Spec PyInstaller cho macOS: tạo TDForestEdu.app (chạy trên Mac, không cross-compile được từ Windows)
APP_NAME = "TDForestEdu"
VERSION = "1.0.0"

a = Analysis(
    ["TDForestEdu.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        "psycopg2", "psycopg2._psycopg", "psycopg2.extensions",
        "openpyxl", "openpyxl.chart", "openpyxl.styles", "xlrd",
        "docx", "docx.oxml", "docx.oxml.ns", "docx.enum.text", "docx.shared",
        "matplotlib.backends.backend_qtagg", "matplotlib.backends.backend_pdf",
        "matplotlib.backends.backend_svg", "matplotlib.backends.backend_ps",
        "mpl_toolkits.mplot3d",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "PyQt5", "PyQt6", "PySide2", "IPython", "notebook"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name=APP_NAME,
    debug=False,
    strip=False,
    upx=False,              # UPX hay làm hỏng chữ ký trên macOS -> tắt
    console=False,          # app có giao diện, không mở Terminal
    argv_emulation=False,
    target_arch=None,       # build theo kiến trúc của máy Mac đang dùng (arm64 hoặc x86_64)
    codesign_identity=None,
    entitlements_file=None,
    icon="TDForestEdu.icns",
)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name=APP_NAME)

app = BUNDLE(
    coll,
    name=f"{APP_NAME}.app",
    icon="TDForestEdu.icns",
    bundle_identifier="com.tdforest.tracdorung",
    version=VERSION,
    info_plist={
        "CFBundleName": APP_NAME,
        "CFBundleDisplayName": "TDForest Edu",
        "CFBundleShortVersionString": VERSION,
        "CFBundleVersion": VERSION,
        "NSHighResolutionCapable": True,
        "LSMinimumSystemVersion": "11.0",
        "NSRequiresAquaSystemAppearance": False,
    },
)
