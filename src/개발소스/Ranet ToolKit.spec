# -*- mode: python ; coding: utf-8 -*-

import os
import re

from PyInstaller.utils.hooks import collect_submodules, collect_dynamic_libs, collect_data_files
from PyInstaller.utils.win32.versioninfo import (
    VSVersionInfo, FixedFileInfo, StringFileInfo, StringTable, StringStruct,
    VarFileInfo, VarStruct,
)

# 버전은 소스의 APP_VERSION 한 곳에서만 관리한다.
# exe 파일명은 'Ranet Toolkit.exe'로 고정(작업표시줄 핀 유지),
# 버전은 Windows 버전 리소스(파일 속성 → 자세히)에 심는다.
with open(os.path.join(SPECPATH, "..", "Ranet ToolKit.py"), encoding="utf-8") as _f:
    APP_VERSION = re.search(r'APP_VERSION\s*=\s*"([\d.]+)"', _f.read()).group(1)

# "6.1" → (6, 1, 0, 0) — 버전 리소스는 4자리 정수 튜플 요구
_ver4 = (tuple(int(x) for x in APP_VERSION.split(".")) + (0, 0, 0, 0))[:4]

version_res = VSVersionInfo(
    ffi=FixedFileInfo(filevers=_ver4, prodvers=_ver4),
    kids=[
        StringFileInfo([StringTable("040904B0", [
            StringStruct("FileDescription", "Ranet Toolkit"),
            StringStruct("FileVersion", APP_VERSION),
            StringStruct("ProductName", "Ranet Toolkit"),
            StringStruct("ProductVersion", APP_VERSION),
        ])]),
        VarFileInfo([VarStruct("Translation", [1033, 1200])]),
    ],
)


# 이 spec은 '개발소스' 하위로 이동됨 — 소스 경로는 프로젝트 루트(..) 기준
datas = [
    ("..\\assets", "assets"),
    ("..\\logo.png", "."),
    ("..\\HWP_PDF변환기_pywinauto\\hwp_to_pdf_gui.py", "HWP_PDF변환기_pywinauto"),
    ("..\\PPTX_이미지변환기\\pptx_converter_v2.py", "PPTX_이미지변환기"),
] + collect_data_files("fitz")

hiddenimports = [
    "pythoncom",
    "pywintypes",
    "win32api",
    "win32clipboard",
    "win32com",
    "win32com.client",
    "win32con",
    "win32gui",
    "win32print",
    "win32process",
    "pyautogui",
    "pyperclip",
    "pptx",
    "pypdf",
    "PIL",
    "requests",
    "reportlab",
    "reportlab.pdfgen",
    "reportlab.pdfgen.canvas",
    "fitz",
    "lxml",
    "lxml.etree",
    "lxml._elementpath",
] + collect_submodules("pptx") + collect_submodules("reportlab") + collect_submodules("fitz")


a = Analysis(
    ["..\\Ranet ToolKit.py"],
    pathex=[],
    binaries=collect_dynamic_libs("fitz"),
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # 미사용 대형 패키지 제외 → onefile 압축해제 빨라짐(스플래시·기동 속도 개선).
    # numpy(약 30~40MB)는 fitz/PIL이 옵션으로만 쓰고 본 앱은 import 안 함 → 안전.
    excludes=["numpy", "scipy", "tkinter", "matplotlib", "pandas", "PyQt5", "PyQt6", "IPython"],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

# 시작 스플래시 — onefile 압축 해제(가장 긴 대기 구간)가 시작되자마자 표시된다.
# 하단 텍스트는 해제 중 자동으로 흐르고, 파이썬 기동 후 main()이 pyi_splash로 단계 문구를 갱신한다.
splash = Splash(
    os.path.join(SPECPATH, "splash.png"),
    binaries=a.binaries,
    datas=a.datas,
    text_pos=(32, 228),
    text_size=9,
    text_color="#8A8FB0",
    text_default="시작 준비 중...",
    minify_script=True,
    always_on_top=False,
)

exe = EXE(
    pyz,
    splash,
    splash.binaries,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Ranet Toolkit",
    version=version_res,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="..\\assets\\RTK.ico",
)
