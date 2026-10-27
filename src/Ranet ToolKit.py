# Ranet Toolkit - Copyright (C) 2026 Region & Agriculture NETwork Cooperative, Jeonnam
# SPDX-License-Identifier: AGPL-3.0-only  (see LICENSE)
from __future__ import annotations

import ctypes               # 자동업데이트: 단일 실행 뮤텍스 / DWM 둥근 모서리
import hashlib              # 자동업데이트: 받은 exe SHA256 대조
from ctypes import wintypes  # PDF → EMF 벡터 변환(GDI 구조체)
import importlib.util
import io
import json                  # 최근 작업 기록(홈 대시보드)
import math                  # PDF 내보내기: 회전 각도 계산
import struct                # PDF 내보내기: EMF 헤더 보정
import os
import re                   # 자동업데이트: 버전 문자열 파싱
import shutil
import subprocess           # 자동업데이트: 교체용 .bat 실행
import sys
import tempfile
import threading
import time
import zipfile
import copy                  # 보고서 hwpx 템플릿 단락/표 복제
import datetime              # 보고서 변환일자 표기

# 일괄병합 자식 프로세스(--batch-worker)는 GUI가 없다. onefile exe로 재실행되면 부트로더가
# 스플래시를 띄우는데, 아래 무거운 임포트(PySide6·win32com·fitz 등)가 끝날 때까지 스플래시가
# 깜빡 보인다(병합 시작 시 스플래시 재출현). → 무거운 임포트 '전에' 여기서 즉시 닫아 깜빡임 제거.
if "--batch-worker" in sys.argv:
    try:
        import pyi_splash
        pyi_splash.close()
    except Exception:
        pass

import pythoncom
from lxml import etree       # 보고서 hwpx(zip+XML) 직접 생성
import requests             # 자동업데이트: version.txt / exe 다운로드
import win32com.client
import win32con
import win32gui
import win32process
import fitz                  # PyMuPDF: 일괄병합 후 이미지 깨짐 검사
from PIL import Image, ImageChops, ImageMath
from pypdf import PdfReader, PdfWriter, PageObject, Transformation
from pypdf.generic import NameObject, NumberObject
from PySide6.QtCore import QEvent, QMimeData, QObject, QPoint, QSize, QThread, QTimer, Qt, Signal, Slot
from PySide6.QtGui import QColor, QDrag, QDragEnterEvent, QDropEvent, QFont, QFontDatabase, QIcon, QImage, QIntValidator, QPainter, QPen, QPixmap, QTransform
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListView,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizeGrip,
    QSizePolicy,
    QStackedWidget,
    QStyle,
    QStyledItemDelegate,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)



# 실패 기록(log_err)은 rtk_common.py로 옮겼다 - 형제 모듈도 함께 쓰기 때문이다.
from rtk_common import ERROR_LOG_PATH, error_log_summary, log_err


# ─── 버전 / 자동업데이트 ───
# v{major}.{minor} 형식. exe 재빌드 시 한 단계 올린다.
APP_VERSION = "7.8"
SOURCE_URL = "https://github.com/skulloom/ranet-toolkit"   # AGPL-3.0 소스 공개 위치(화면 고지·README)
# GitHub 레포(Public). version.txt(최신 버전 평문)와 릴리스 자산(.exe)을 여기서 받는다.
UPDATE_REPO = "skulloom/ranet-toolkit"
# version.txt: 최신 버전 문자열(예: "1.4")만 담긴 평문 파일.
# raw.githubusercontent.com은 ~5분 CDN 캐시 + 쿼리스트링 무시라 즉시 못 받는다.
# 캐시 없는 GitHub API contents 엔드포인트로 version.txt를 읽는다
# (Accept: application/vnd.github.raw 헤더로 평문 그대로 수신).
UPDATE_VERSION_URL = f"https://api.github.com/repos/{UPDATE_REPO}/contents/version.txt"
# changelog.txt: 버전별 변경사항. 업데이트 알림/완료 팝업에 표시한다.
UPDATE_CHANGELOG_URL = f"https://api.github.com/repos/{UPDATE_REPO}/contents/changelog.txt"
# 최신 릴리스 메타. 자산 이름이 무엇이든(.exe면) 여기서 다운로드 URL을 찾는다.
UPDATE_RELEASES_URL = f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest"
# 폴백용 고정 이름 URL(릴리스 API 실패 시). 실제 자산 이름은 Ranet.Toolkit.exe —
# 예전 값 RanetToolkit.exe 는 404 였다(2026-10-27 실측: 폴백이 한 번도 동작 안 함).
UPDATE_DOWNLOAD_URL = f"https://github.com/{UPDATE_REPO}/releases/latest/download/Ranet.Toolkit.exe"

# 업데이트 통신용 공유 세션.
_session = requests.Session()


def _fetch_release_asset() -> dict:
    # 최신 릴리스에서 첫 .exe 자산의 다운로드 URL·크기·SHA256 을 찾아 돌려준다.
    # 자산 파일명이 무엇이든 동작한다. GitHub 이 자산마다 digest("sha256:...")를 자동으로 붙인다.
    # 실패 시 고정 이름 URL 만(크기·해시 없음 → 머리·길이·하한 검사만 한다).
    try:
        r = _session.get(
            UPDATE_RELEASES_URL,
            headers={"Accept": "application/vnd.github+json", "Cache-Control": "no-cache"},
            timeout=10,
        )
        if r.status_code == 200:
            for asset in r.json().get("assets", []):
                if str(asset.get("name", "")).lower().endswith(".exe"):
                    url = asset.get("browser_download_url")
                    if url:
                        digest = str(asset.get("digest") or "")
                        return {
                            "url": url,
                            "size": int(asset.get("size") or 0),
                            "sha256": digest.split(":", 1)[1].lower() if digest.startswith("sha256:") else "",
                        }
    except Exception as e:
        log_err("_fetch_release_asset", e)
    return {"url": UPDATE_DOWNLOAD_URL, "size": 0, "sha256": ""}


# 받은 exe 검증 기준. Windows 실행 파일은 반드시 'MZ' 로 시작한다(HTML 오류 페이지·빈 파일 거름).
EXE_SIGNATURE = b"MZ"
# 새 exe 가 지금 exe 의 이 비율보다 작으면 거부(릴리스 메타를 못 받았을 때의 마지막 방어선).
MINIMUM_SIZE_RATIO = 0.5


def describe_download_problem(received: dict, content_length: int, expected: dict,
                              minimum_size: int) -> str:
    # 받은 파일이 교체해도 되는 exe 인지 판정. 문제없으면 "", 있으면 사용자에게 보일 이유 한 줄.
    # received = {"head": 첫 바이트들, "size": 받은 바이트 수, "sha256": 16진 해시}.
    head, size, sha256_hex = received["head"], received["size"], received["sha256"]
    if not head.startswith(EXE_SIGNATURE):
        return "받은 파일이 실행 파일이 아닙니다"
    if content_length and size != content_length:
        return f"다운로드가 중간에 끊겼습니다 ({size:,} / {content_length:,} 바이트)"
    if expected.get("size") and size != expected["size"]:
        return "받은 파일 크기가 배포본과 다릅니다"
    if expected.get("sha256") and sha256_hex.lower() != expected["sha256"]:
        return "받은 파일 내용이 배포본과 다릅니다"
    if size < minimum_size:
        return "받은 파일이 너무 작습니다"
    return ""


def _minimum_update_size() -> int:
    # 지금 실행 중인 exe 크기의 절반. 소스 실행이면 하한 없음.
    if not getattr(sys, "frozen", False):
        return 0
    try:
        return int(os.path.getsize(sys.executable) * MINIMUM_SIZE_RATIO)
    except OSError:
        return 0


def _parse_version(v: str) -> tuple:
    # "1.2" → (1, 2). 숫자만 추출해 정수 튜플로 비교한다.
    nums = re.findall(r"\d+", str(v or ""))
    return tuple(int(n) for n in nums) if nums else (0,)


def is_newer_version(remote: str, current: str) -> bool:
    # 원격 버전이 현재보다 높으면 True.
    return _parse_version(remote) > _parse_version(current)


def fetch_changelog_text() -> str:
    # 레포의 changelog.txt 전문을 받아온다. 실패하면 빈 문자열(팝업은 변경사항 없이 표시).
    # version.txt와 같은 이유로 raw CDN 대신 캐시 없는 GitHub API contents를 쓴다.
    try:
        r = _session.get(
            UPDATE_CHANGELOG_URL,
            params={"ref": "main"},
            headers={
                "Accept": "application/vnd.github.raw",
                "Cache-Control": "no-cache",
                "Pragma": "no-cache",
            },
            timeout=5,
        )
        if r.status_code == 200:
            return r.text or ""
    except Exception as e:
        log_err("fetch_changelog_text", e)
    return ""


def changelog_notes(text: str, newer_than: str = "", only: str = "") -> str:
    # changelog.txt에서 표시할 항목을 골라낸다. 파일 형식:
    #   [6.2]
    #   - 변경 내용 한 줄
    # only="6.2"       → 해당 버전 섹션만 (업데이트 완료 팝업용)
    # newer_than="6.0" → 그보다 높은 버전 전부 (여러 버전 건너뛴 업데이트 알림용)
    sections = re.split(r"^\[([\d.]+)\]\s*$", text, flags=re.M)
    out = []
    # re.split 결과: [머리말, 버전, 본문, 버전, 본문, ...] 꼴로 번갈아 나온다.
    for i in range(1, len(sections) - 1, 2):
        ver, body = sections[i], sections[i + 1].strip()
        if not body:
            continue
        if only and _parse_version(ver) != _parse_version(only):
            continue
        if newer_than and not is_newer_version(ver, newer_than):
            continue
        out.append(body if only else f"v{ver}\n{body}")
    return "\n\n".join(out)


BASE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


def _load_module(name: str, rel_path: str):
    path = os.path.join(BASE_DIR, rel_path)
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load module: {path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


hwp_mod = _load_module("ranet_hwp_pywinauto", os.path.join("HWP_PDF변환기_pywinauto", "hwp_to_pdf_gui.py"))
ppt_mod = _load_module("ranet_pptx_converter", os.path.join("PPTX_이미지변환기", "pptx_converter_v2.py"))


# ── 색상 (모던 인디고, 공시지가조회기 UI 참고) ──
BG          = "#EDEEF6"   # 창 바깥(반투명 영역 뒤). 프레임리스라 거의 안 보임
CARD        = "#F4F5FB"   # 앱 카드(콘텐츠) 배경
SURFACE     = "#FFFFFF"   # 내부 카드/입력 표면
ACCENT      = "#6C5CE7"
ACCENT_H    = "#5A4FCF"
ACCENT_P    = "#4B3FB8"
ACCENT_SOFT = "#EEECFC"
DANGER      = "#EF4444"
OK          = "#10B981"   # 성공/정상 상태(에메랄드)
SUCCESS     = "#10B981"
BORDER      = "#E6E8F2"
DIVIDER     = "#ECEDF6"
TXT         = "#1F2937"
TXT_S       = "#4B5563"
TXT_T       = "#8A8FB0"

_make_light_palette = ppt_mod._make_light_palette
ConvertWorker = ppt_mod.ConvertWorker
ElideLabel = ppt_mod.ElideLabel
RoundedPreviewWidget = ppt_mod.RoundedPreviewWidget


# 임베드 폰트(Pretendard) 로드 후 실제 패밀리명으로 교체된다. 실패 시 맑은 고딕 폴백.
FONT_FAMILY = "Malgun Gothic"

# Static TTF만 등록 (OTF/Variable 금지 — 작은 한글 렌더가 무르고 weight 매칭이 모호해짐)
_FONT_FILES = ("Pretendard-Regular.ttf", "Pretendard-Medium.ttf",
               "Pretendard-SemiBold.ttf", "Pretendard-Bold.ttf")


def _font(pt: int, wt: str = "regular") -> QFont:
    # 'regular' 요청도 Medium(500)으로 — Regular(400)는 작은 한글에서 얇아 자글거려 보임.
    w_map = {
        "semibold": QFont.Weight.DemiBold,
        "bold": QFont.Weight.Bold,
        "medium": QFont.Weight.Medium,
        "regular": QFont.Weight.Medium,
        "light": QFont.Weight.Normal,
    }
    font = QFont(FONT_FAMILY, pt)
    font.setWeight(w_map.get(wt, QFont.Weight.Medium))
    # FreeType 엔진 + 풀힌팅 → 픽셀 정렬 렌더로 또렷해진다 (란맵 v8 검증)
    font.setHintingPreference(QFont.HintingPreference.PreferFullHinting)
    font.setStyleStrategy(QFont.StyleStrategy.PreferAntialias)
    return font


def _load_fonts() -> None:
    # Pretendard Static TTF 등록. assets/fonts는 spec datas로 exe에 자동 동봉된다.
    # QApplication 생성 후 호출해야 한다.
    global FONT_FAMILY
    font_dir = os.path.join(BASE_DIR, "assets", "fonts")
    family = ""
    for name in _FONT_FILES:
        p = os.path.join(font_dir, name)
        if not os.path.isfile(p):
            continue
        fid = QFontDatabase.addApplicationFont(p)
        fams = QFontDatabase.applicationFontFamilies(fid) if fid >= 0 else []
        if fams:
            family = fams[0]
    if family:
        FONT_FAMILY = family

# 선택/활성 채움 버튼(품질·DPI 선택 등)
BTN_ON = (
    f"QPushButton{{background:{ACCENT};color:white;border:none;border-radius:10px;}}"
    f"QPushButton:hover{{background:{ACCENT_H};}}"
    f"QPushButton:pressed{{background:{ACCENT_P};}}"
)
# 실행 준비(글로우 펄스) 버튼 — _PulseTimer가 배경색을 애니메이션으로 덮어씀
BTN_ACTION = f"background:#FFFFFF;color:{ACCENT};border:2px solid {ACCENT};border-radius:12px;"
# 비활성(실행 불가)
BTN_OFF = "background:#C2C0E6;color:#FFFFFF;border:none;border-radius:12px;"
# 중지(빨강 채움)
BTN_STOP = (
    f"QPushButton{{background:{DANGER};color:white;border:none;border-radius:12px;}}"
    "QPushButton:hover{background:#DC2626;}"
    "QPushButton:pressed{background:#B91C1C;}"
)
# 비활성/대기용 연한 회색
BTN_SOFT = (
    "QPushButton{background:#E9EAF6;color:#9AA0B5;border:none;border-radius:12px;}"
    f"QPushButton:hover{{background:{ACCENT_SOFT};color:{ACCENT};}}"
)
# 폴더/파일 찾기(인디고 채움 소형)
BTN_BROWSE = (
    f"QPushButton{{background:{ACCENT};color:white;border:none;border-radius:10px;padding:0 18px;}}"
    f"QPushButton:hover{{background:{ACCENT_H};}}"
    f"QPushButton:pressed{{background:{ACCENT_P};}}"
)
# 세그먼트 탭(활성/비활성)
# 사이드바 내비 버튼(활성=인디고 채움 / 비활성=투명, hover 연보라)
NAV_ON = (
    f"QPushButton {{ background:{ACCENT}; color:white; border:none;"
    f" border-radius:10px; text-align:left; padding-left:12px; }}"
)
NAV_OFF = (
    f"QPushButton {{ background:transparent; color:{TXT_S}; border:none;"
    f" border-radius:10px; text-align:left; padding-left:12px; }}"
    f"QPushButton:hover {{ background:{ACCENT_SOFT}; color:{ACCENT}; }}"
)
CHECK_ICON = os.path.join(BASE_DIR, "assets", "check.svg").replace("\\", "/")
RED_PLUS = os.path.join(BASE_DIR, "assets", "plus_red.svg").replace("\\", "/")
# 일괄병합 결과 보고서(.hwpx) 템플릿 — 표 서식을 복제해 채운다(HWP COM 불필요).
REPORT_TEMPLATE = os.path.join(BASE_DIR, "assets", "report_template.hwpx")

QSS = f"""
* {{ font-family: "Pretendard", "맑은 고딕", "Segoe UI", sans-serif; font-weight: 500; outline: none; }}
QMainWindow, QWidget#Central {{ background: {CARD}; }}
QWidget#Content {{ background: {CARD}; }}

QFrame#Sidebar {{ background: {SURFACE}; border: none; border-right: 1px solid {BORDER}; }}
QLabel#BrandName {{ color: {TXT}; background: transparent; }}
QLabel#BrandSub {{ color: {TXT_T}; background: transparent; }}
QLabel#NavSection {{ color: {TXT_T}; background: transparent; padding-left: 6px; }}
QLabel#VersionChip {{ color: {TXT_S}; background: {CARD}; border: none; border-radius: 10px; padding: 8px 12px; }}

QFrame#TitleBar {{ background: {SURFACE}; border: none; border-bottom: 1px solid {BORDER}; }}
QLabel#Crumb {{ color: {TXT_S}; background: transparent; }}
QPushButton#WinMin, QPushButton#WinMax, QPushButton#WinClose {{
    background: transparent; border: none; border-radius: 6px;
}}
QPushButton#WinMin:hover, QPushButton#WinMax:hover {{ background: rgba(31,41,55,0.08); }}
QPushButton#WinClose:hover {{ background: #EF4444; }}

QFrame#page_shell {{ background: transparent; border: none; }}
QWidget#root, QWidget#scroll_content {{ background: transparent; }}
QScrollArea {{ background: transparent; border: none; }}

QFrame#card {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 14px;
}}
QFrame#drop_card {{
    background: {SURFACE};
    border: 2px dashed #CBCFEA;
    border-radius: 14px;
}}

QTextEdit#log_box {{
    background: #FBFBFE;
    border: 1px solid {BORDER};
    border-radius: 10px;
    color: {TXT};
    padding: 6px;
    selection-background-color: {ACCENT_SOFT};
}}

QLineEdit {{
    background: #FFFFFF;
    border: 1px solid {BORDER};
    border-radius: 10px;
    color: {TXT};
    padding: 8px 11px;
}}
QLineEdit:focus {{ border: 2px solid {ACCENT}; }}
QLineEdit:disabled {{ color: {TXT_T}; background: #F1F2F8; }}

QLabel {{ color: {TXT}; background: transparent; }}

QCheckBox, QRadioButton {{ color: {TXT_S}; background: transparent; spacing: 8px; }}
QCheckBox::indicator {{
    width: 16px; height: 16px;
    border: 1px solid #CBCFEA; border-radius: 5px; background: white;
}}
QCheckBox::indicator:hover {{ border: 1px solid {ACCENT}; }}
QCheckBox::indicator:checked {{
    border: 1px solid {ACCENT}; background: white; image: url("{CHECK_ICON}");
}}
QRadioButton::indicator {{
    width: 16px; height: 16px;
    border: 1px solid #CBCFEA; border-radius: 8px; background: white;
}}
QRadioButton::indicator:hover {{ border: 1px solid {ACCENT}; }}
QRadioButton::indicator:checked {{
    border: 1px solid {ACCENT};
    background: qradialgradient(cx:0.5, cy:0.5, radius:0.55, fx:0.5, fy:0.5,
        stop:0 {ACCENT}, stop:0.40 {ACCENT}, stop:0.44 white, stop:1 white);
}}

QProgressBar {{
    background: #E9EAF6; border: none; border-radius: 4px;
}}
QProgressBar::chunk {{ background: {ACCENT}; border-radius: 4px; }}

QScrollBar:vertical {{ background: transparent; width: 10px; margin: 4px 2px 4px 2px; }}
QScrollBar::handle:vertical {{ background: #C2C6E0; min-height: 30px; border-radius: 5px; }}
QScrollBar::handle:vertical:hover {{ background: {ACCENT}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
    background: transparent; border: none; height: 0px;
}}
QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px 4px 2px 4px; }}
QScrollBar::handle:horizontal {{ background: #C2C6E0; min-width: 30px; border-radius: 5px; }}
QScrollBar::handle:horizontal:hover {{ background: {ACCENT}; }}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal,
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
    background: transparent; border: none; width: 0px;
}}

QMessageBox {{ background: {SURFACE}; }}
QMessageBox QLabel {{ color: {TXT}; background: transparent; }}
QMessageBox QPushButton {{
    background: {SURFACE}; color: {TXT_S};
    border: 1px solid {BORDER}; border-radius: 8px; padding: 6px 16px; min-width: 64px;
}}
QMessageBox QPushButton:hover {{ background: {ACCENT_SOFT}; border-color: {ACCENT}; color: {ACCENT}; }}
"""


def _clear_effect(widget: QWidget) -> None:
    widget.setGraphicsEffect(None)


def _apply_glow(widget: QWidget, color: str = ACCENT, blur: int = 20) -> None:
    # 글로우(QGraphicsDropShadowEffect) 사용 중지: 이펙트가 걸린 위젯은 픽스맵으로
    # 래스터화되어 서브픽셀 AA가 꺼진다(글자 자글거림 — 란맵 v8에서 확인).
    # 실행 준비 강조는 _PulseTimer 배경 애니메이션만으로 충분하다.
    return None


def _unique_path(path: str) -> str:
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    n = 1
    while True:
        cand = f"{root}({n}){ext}"
        if not os.path.exists(cand):
            return cand
        n += 1


def _natural_key(s: str) -> list:
    # 자연정렬 키: 문자열 속 숫자를 정수로 비교해 '1,2,10' 순서가 되게 한다.
    # ('2. 자격' < '10. 검토' 처럼 폴더/파일명 접두 번호를 의도대로 정렬)
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def _safe_name(name: str) -> str:
    # 파일명으로 못 쓰는 문자(\ / : * ? " < > |)를 _ 로 치환한다.
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip().rstrip(".") or "무제"


def _get_writer_size(writer: PdfWriter) -> int:
    buf = io.BytesIO()
    writer.write(buf)
    return buf.tell()


def _is_hidden_or_system(path: str) -> bool:
    # Windows 숨김(0x2)·시스템(0x4) 속성 파일이면 True.
    # OneDrive 온라인전용 스텁·OS 부산물 등 일반 사용자 문서가 아닌 유령 파일을 거른다.
    try:
        attrs = os.stat(path).st_file_attributes
    except (OSError, AttributeError):
        return False
    return bool(attrs & 0x2 or attrs & 0x4)


# ─── 진단 로그 전송 ───
# 오류 발생 시 사용자 동의(팝업 예)를 받아 작업 로그를 라넷 서버 보고 창구로 보낸다.
# 서버(데이터 통로 /data/report/toolkit)가 파일로 남기고 비공개 소스 레포에 이슈를 올린다.
# 토큰은 서버에만 둔다 — exe·소스에 넣지 않는다(소스 공개 대비).
DIAG_REPORT_URL = "https://jnranet.com/data/report/toolkit"


def _pc_id() -> str:
    # 같은 PC 에서 반복되는 문제인지 보려는 익명 표지(컴퓨터·사용자 이름 해시 앞 8자리)
    raw = (os.environ.get("COMPUTERNAME", "") + os.environ.get("USERNAME", "")).encode()
    return hashlib.sha1(raw).hexdigest()[:8]


class DiagReportSender(QThread):
    # 진단 로그를 서버 보고 창구로 전송 (백그라운드 — UI 비차단).
    done = Signal(bool, str)   # 성공 여부, 실패 시 오류 메시지

    def __init__(self, summary: str, detail: str, log: str, context: dict) -> None:
        super().__init__()
        import platform
        self._report = {
            "kind": "error",
            "time": datetime.datetime.now().isoformat(timespec="seconds"),
            "summary": summary,
            "detail": detail,
            "log": log,
            "context": context,
            "env": {"app_version": APP_VERSION, "frozen": bool(getattr(sys, "frozen", False)),
                    "python": sys.version.split()[0], "os": platform.platform(),
                    "cpu_count": os.cpu_count(), "pc": _pc_id()},
        }

    def run(self) -> None:
        try:
            r = requests.post(DIAG_REPORT_URL, json=self._report, timeout=15)
            self.done.emit(r.status_code == 200, "" if r.status_code == 200 else f"HTTP {r.status_code}")
        except Exception as e:
            self.done.emit(False, str(e) or e.__class__.__name__)


def _korean_buttons(msg: QMessageBox) -> None:
    # Qt 기본 버튼 라벨(OK/Yes/No)이 영문으로 떠서 한국어 안내문과 안 맞는 문제 —
    # 표준 버튼 텍스트를 한국어로 교체한다. (란맵과 동일)
    for sb, label in ((QMessageBox.StandardButton.Ok, "확인"),
                      (QMessageBox.StandardButton.Yes, "예"),
                      (QMessageBox.StandardButton.No, "아니요")):
        b = msg.button(sb)
        if b is not None:
            b.setText(label)


class _SHFILEOPSTRUCTW(ctypes.Structure):
    # SHFileOperationW 입력 구조체 (휴지통 이동용)
    _fields_ = [
        ("hwnd", ctypes.c_void_p),
        ("wFunc", ctypes.c_uint),
        ("pFrom", ctypes.c_wchar_p),
        ("pTo", ctypes.c_wchar_p),
        ("fFlags", ctypes.c_ushort),
        ("fAnyOperationsAborted", ctypes.c_int),
        ("hNameMappings", ctypes.c_void_p),
        ("lpszProgressTitle", ctypes.c_wchar_p),
    ]


def _delete_to_trash(path: str) -> None:
    # 영구 삭제 대신 휴지통으로 보낸다(FOF_ALLOWUNDO — 사용자가 복구 가능).
    # 실패 시 OSError. (네트워크 드라이브 등 휴지통이 없는 위치는 실패할 수 있음)
    FO_DELETE = 3
    FOF_SILENT, FOF_NOCONFIRMATION, FOF_ALLOWUNDO, FOF_NOERRORUI = 0x4, 0x10, 0x40, 0x400
    op = _SHFILEOPSTRUCTW()
    op.wFunc = FO_DELETE
    op.pFrom = os.path.abspath(path) + "\x00"   # 이중 null 종단(목록 형식) 필요
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI
    res = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    if res != 0 or op.fAnyOperationsAborted:
        raise OSError(f"휴지통 이동 실패 (code {res})")


def _tinted_pixmap(icon_name: str, color: str, size: int = 20) -> QPixmap:
    # 흰색 스트로크 SVG(assets/{icon_name}_on.svg)를 지정 색으로 틴트해 반환.
    # 홈 대시보드 카드(흰 배경)에서 인디고 아이콘으로 재사용하기 위함.
    src = QIcon(os.path.join(BASE_DIR, "assets", f"{icon_name}_on.svg")).pixmap(size, size)
    pm = QPixmap(src.size())
    pm.setDevicePixelRatio(src.devicePixelRatio())
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.drawPixmap(0, 0, src)
    p.setCompositionMode(QPainter.CompositionMode.CompositionMode_SourceIn)
    p.fillRect(pm.rect(), QColor(color))
    p.end()
    return pm


# ── 최근 작업 기록 (홈 대시보드 표시용) ──
RECENT_PATH = os.path.join(
    os.environ.get("LOCALAPPDATA", tempfile.gettempdir()), "RanetToolkit", "recent.json")
RECENT_MAX = 30


def _recent_load() -> list[dict]:
    # 최근 작업 목록(최신순). 파일 없음/손상 시 빈 목록.
    try:
        with open(RECENT_PATH, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except (OSError, ValueError):
        return []


def _recent_save(items: list[dict]) -> None:
    # 기록 저장. 실패는 작업 성패와 무관하므로 조용히 무시.
    try:
        # 임시 파일에 다 쓴 뒤 한 번에 바꾼다 — 쓰다 꺼져도 앞 기록이 남는다(깨지면 목록이 통째로 빔).
        os.makedirs(os.path.dirname(RECENT_PATH), exist_ok=True)
        draft_path = f"{RECENT_PATH}.{os.getpid()}.tmp"
        with open(draft_path, "w", encoding="utf-8") as f:
            json.dump(items[:RECENT_MAX], f, ensure_ascii=False)
        os.replace(draft_path, RECENT_PATH)
    except OSError as e:
        log_err("_recent_save", e, RECENT_PATH)


def _recent_add(tool: str, title: str, meta: str) -> None:
    # 작업 완료 시 1건 추가.
    items = _recent_load()
    items.insert(0, {
        "tool": tool, "title": title, "meta": meta,
        "when": datetime.datetime.now().strftime("%m-%d %H:%M"),
    })
    _recent_save(items)


def _card(title: str | None = None) -> tuple[QFrame, QVBoxLayout]:
    card = QFrame()
    card.setObjectName("card")
    # 드롭섀도 제거: 이펙트가 걸리면 카드+자식 전체가 픽스맵 래스터화되어
    # 글자 서브픽셀 AA가 꺼진다. 1px 보더(QSS #card)만으로 구분한다.
    layout = QVBoxLayout(card)
    layout.setContentsMargins(15, 13, 15, 15)
    layout.setSpacing(8)
    if title:
        lbl = QLabel(title)
        lbl.setFont(_font(9, "semibold"))
        lbl.setStyleSheet(f"color:{ACCENT};background:transparent;letter-spacing:0.2px;")
        layout.addWidget(lbl)
    return card, layout


class WheelLockScrollArea(QScrollArea):
    def wheelEvent(self, event) -> None:
        super().wheelEvent(event)
        event.accept()


class ThumbDelegate(QStyledItemDelegate):
    """끼워넣은(추가) 페이지를 연한 빨강으로 칠하는 델리게이트.

    위젯 스타일시트(::item)를 쓰면 QListWidgetItem.setBackground(BackgroundRole)가
    무시되므로, 배경은 델리게이트가 직접 그린다. 선택된 항목은 스타일시트의
    선택색이 그려지도록 칠하지 않는다.
    """

    INSERTED = QColor("#FCE3E3")

    def paint(self, painter, option, index) -> None:
        d = index.data(Qt.ItemDataRole.UserRole)
        selected = bool(option.state & QStyle.StateFlag.State_Selected)
        if isinstance(d, dict) and d.get("inserted") and not selected:
            painter.save()
            painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(self.INSERTED)
            painter.drawRoundedRect(option.rect.adjusted(2, 2, -2, -2), 6, 6)
            painter.restore()
        super().paint(painter, option, index)


class ThumbGridList(QListWidget):
    """PDF 편집 탭의 썸네일 격자.

    - 휠 스크롤을 이 위젯 안에서만 처리하고 항상 소비한다(맨 아래/위에 닿아도
      부모 스크롤영역=탭 전체로 전파되지 않게).
    - 드롭(순서변경)을 직접 처리한다. IconMode+InternalMove의 기본 동작은 아이콘을
      '자유 위치'로 옮기면서 행이 사라지거나 중복되는 버그가 있어, 행 순서만 바꾸도록
      재정의한다.
    """

    reordered = Signal()
    zoomRequested = Signal(int)   # Ctrl+휠: +1 확대 / -1 축소

    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self._drop_line: tuple[int, int, int] | None = None   # (x, y1, y2) 삽입 표시선

    def wheelEvent(self, event) -> None:
        # Ctrl을 누른 채 휠을 굴리면 스크롤 대신 썸네일 크기 확대/축소.
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            self.zoomRequested.emit(1 if event.angleDelta().y() > 0 else -1)
            event.accept()
            return
        bar = self.verticalScrollBar()
        bar.setValue(bar.value() - event.angleDelta().y())
        event.accept()

    # ── 커스텀 드래그앤드롭 ──
    # IconMode 기본 DnD는 (a)드래그 중 원본 아이콘이 사라져 보이고 (b)삽입 위치 표시가
    # 부실하다. 그래서 드래그 픽스맵(따라다니는 썸네일)·삽입선·재배치를 직접 그린다.
    def startDrag(self, supportedActions) -> None:
        items = self.selectedItems()
        if not items:
            return
        drag = QDrag(self)
        drag.setMimeData(QMimeData())   # 내부 이동: 선택 항목은 dropEvent에서 직접 읽음
        pm = items[0].icon().pixmap(self.iconSize())
        if len(items) > 1:
            # 여러 장이면 장수 배지를 픽스맵 위에 살짝 표시.
            pm = QPixmap(pm)
            p = QPainter(pm)
            p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
            p.setBrush(QColor(ACCENT))
            p.setPen(Qt.PenStyle.NoPen)
            p.drawRoundedRect(pm.width() - 34, 6, 28, 20, 6, 6)
            p.setPen(QColor("#FFFFFF"))
            p.setFont(_font(9, "bold"))
            p.drawText(pm.width() - 34, 6, 28, 20, Qt.AlignmentFlag.AlignCenter, str(len(items)))
            p.end()
        drag.setPixmap(pm)
        drag.setHotSpot(QPoint(pm.width() // 2, pm.height() // 2))
        drag.exec(Qt.DropAction.MoveAction)

    def dragEnterEvent(self, event) -> None:
        if event.source() is self:
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event) -> None:
        if event.source() is self:
            event.acceptProposedAction()
            pos = event.position().toPoint() if hasattr(event, "position") else event.pos()
            self._update_drop_line(pos)
            self.viewport().update()
        else:
            super().dragMoveEvent(event)

    def dragLeaveEvent(self, event) -> None:
        self._drop_line = None
        self.viewport().update()
        super().dragLeaveEvent(event)

    def _drop_target_row(self, pos: QPoint) -> int:
        item = self.itemAt(pos)
        if item is None:
            return self.count()
        row = self.row(item)
        return row + 1 if pos.x() > self.visualItemRect(item).center().x() else row

    def _update_drop_line(self, pos: QPoint) -> None:
        # 삽입될 위치(행 인덱스)를 구해 그 자리에 그릴 세로선 좌표를 계산.
        if self.count() == 0:
            self._drop_line = None
            return
        row = self._drop_target_row(pos)
        if row >= self.count():
            r = self.visualItemRect(self.item(self.count() - 1))
            self._drop_line = (r.right() + 4, r.top(), r.bottom())
        else:
            r = self.visualItemRect(self.item(row))
            self._drop_line = (r.left() - 4, r.top(), r.bottom())

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        if self._drop_line is not None:
            x, y1, y2 = self._drop_line
            p = QPainter(self.viewport())
            pen = QPen(QColor(ACCENT))
            pen.setWidth(3)
            pen.setCapStyle(Qt.PenCapStyle.RoundCap)
            p.setPen(pen)
            p.drawLine(x, y1 + 2, x, y2 - 2)
            p.end()

    def dropEvent(self, event) -> None:
        self._drop_line = None
        # 외부(다른 위젯/파일)에서 온 드롭은 기본 처리에 맡긴다.
        if event.source() is not self:
            super().dropEvent(event)
            return
        pos = event.position().toPoint() if hasattr(event, "position") else event.pos()
        target_row = self._drop_target_row(pos)
        rows = sorted(self.row(it) for it in self.selectedItems())
        if not rows:
            event.ignore()
            return
        self.move_rows(rows, target_row)
        event.setDropAction(Qt.DropAction.MoveAction)
        event.accept()
        self.viewport().update()
        self.reordered.emit()

    def move_rows(self, rows: list[int], target_row: int) -> None:
        """선택 행(rows)을 target_row 위치로 이동(행 순서만 재배치). 아이템 유실 없음."""
        # 높은 인덱스부터 떼어내고, 떼어낸 위치가 목표보다 앞이면 목표 인덱스를 보정.
        taken = []
        for r in sorted(rows, reverse=True):
            taken.insert(0, self.takeItem(r))
            if r < target_row:
                target_row -= 1
        for off, it in enumerate(taken):
            self.insertItem(target_row + off, it)
        self.clearSelection()
        for it in taken:
            it.setSelected(True)


def _make_scroll(content: QWidget) -> QScrollArea:
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
    scroll.setWidget(content)
    return scroll


_LIST_SCROLLBAR_QSS = (
    "QScrollBar:vertical { background:transparent; width:9px; margin:4px 2px 4px 2px; }"
    "QScrollBar::handle:vertical { background:#CAD4E4; min-height:28px; border-radius:4px; }"
    "QScrollBar::handle:vertical:hover { background:#AAB8CC; }"
    "QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,"
    "QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background:transparent; border:none; height:0px; }"
    "QScrollBar:horizontal { background:transparent; height:9px; margin:2px 4px 2px 4px; }"
    "QScrollBar::handle:horizontal { background:#CAD4E4; min-width:28px; border-radius:4px; }"
    "QScrollBar::handle:horizontal:hover { background:#AAB8CC; }"
    "QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal,"
    "QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background:transparent; border:none; width:0px; }"
)


def _make_check_list(height: int, accent: str = ACCENT, check_icon: str = CHECK_ICON) -> QListWidget:
    # 대량(수백 개) 파일 미리보기용 체크 리스트.
    # 행마다 QCheckBox 위젯을 만드는 방식은 수백 개에서 생성·스크롤이 무거워지므로
    # QListWidget 체크 아이템으로 처리한다(수천 행도 가벼움).
    lw = QListWidget()
    lw.setFixedHeight(height)
    lw.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
    lw.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
    lw.setStyleSheet(
        f"QListWidget {{ background:#FBFBFE; border:1px solid {BORDER};"
        f" border-radius:10px; outline:none; padding:4px; color:{TXT}; }}"
        "QListWidget::item { padding:2px 1px; border:none; }"
        "QListWidget::indicator { width:15px; height:15px; border:1px solid #CBCFEA;"
        " border-radius:4px; background:white; }"
        f"QListWidget::indicator:hover {{ border:1px solid {accent}; }}"
        f"QListWidget::indicator:checked {{ border:1px solid {accent}; background:white;"
        f" image:url('{check_icon}'); }}"
        + _LIST_SCROLLBAR_QSS
    )

    # 행 어디를 클릭해도 체크 토글(체크박스만 노리지 않아도 됨).
    # 아이템 플래그에서 UserCheckable을 빼 indicator 자체 토글과 중복되지 않게 한다.
    def _flip(it: QListWidgetItem) -> None:
        if it.data(Qt.ItemDataRole.UserRole) is None:
            return   # 안내문(placeholder) 행은 무시
        it.setCheckState(
            Qt.CheckState.Unchecked if it.checkState() == Qt.CheckState.Checked
            else Qt.CheckState.Checked)

    lw.itemClicked.connect(_flip)
    lw.setCursor(Qt.CursorShape.PointingHandCursor)
    return lw


def _add_check_item(lw: QListWidget, text: str, data, checked: bool) -> QListWidgetItem:
    # 체크 리스트에 행 1개 추가. data(파일/폴더 경로)는 UserRole에 보관.
    # 체크 토글은 _make_check_list의 행 클릭 핸들러가 담당(UserCheckable 플래그 없음).
    it = QListWidgetItem(text)
    it.setData(Qt.ItemDataRole.UserRole, data)
    it.setFont(_font(8, "regular"))
    it.setFlags(Qt.ItemFlag.ItemIsEnabled)
    it.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
    lw.addItem(it)
    return it


def _placeholder_item(lw: QListWidget, text: str) -> None:
    # 비어 있을 때 안내문 1줄(선택/체크 불가).
    it = QListWidgetItem(text)
    it.setFlags(Qt.ItemFlag.NoItemFlags)
    it.setFont(_font(9, "regular"))
    it.setForeground(QColor(TXT_S))
    lw.addItem(it)


class LogoLogTextEdit(QTextEdit):
    def __init__(self) -> None:
        super().__init__()
        self._logo = QPixmap(os.path.join(BASE_DIR, "logo.png"))

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        if self._logo.isNull():
            return
        viewport = self.viewport()
        target_w = min(250, max(140, int(viewport.width() * 0.42)))
        scaled = self._logo.scaledToWidth(target_w, Qt.TransformationMode.SmoothTransformation)
        x = (viewport.width() - scaled.width()) // 2
        y = (viewport.height() - scaled.height()) // 2
        painter = QPainter(viewport)
        painter.setOpacity(0.09)
        painter.drawPixmap(x, y, scaled)


# ── Office·한글 COM 안전 규칙 ──────────────────────────────────────────
# ⚠️ win32com 의 Dispatch 는 «새로 띄우는» 게 아니라, 이미 떠 있는 인스턴스가 있으면
#    거기에 그대로 붙는다. 사용자가 편집 중인 파워포인트·한글에 붙은 뒤 Quit 을 부르면
#    그 앱이 통째로 닫히고, 자동화 세션이라 저장 확인 창도 안 떠서 미저장 편집분이
#    그대로 사라진다. 그래서 «내가 띄운 것인지» 를 구분해 두고, 내가 띄운 것만 닫는다.
_APP_PROCS = {
    "Word.Application":       ("WINWORD.EXE",),
    "PowerPoint.Application": ("POWERPNT.EXE",),
    "HWPFrame.HwpObject":     ("Hwp.exe", "HwpApp.exe", "Hwp90.exe"),
}


def _proc_running(names: tuple) -> bool:
    """해당 실행파일이 지금 떠 있는지. 확인에 실패하면 «떠 있다»로 본다(보수적)."""
    if not names:
        return True
    try:
        out = subprocess.run(
            ["tasklist", "/NH", "/FO", "CSV"],
            capture_output=True, text=True, errors="ignore",
            creationflags=0x08000000,   # CREATE_NO_WINDOW
            timeout=10).stdout.upper()
    except Exception as exc:
        log_err("proc_running", exc, ",".join(names))
        return True
    return any(n.upper() in out for n in names)


def com_app(prog_id: str):
    """COM 앱 핸들. 반환 (app, created) — created=True 면 «내가 띄운 것»이라 닫아도 된다."""
    running_before = _proc_running(_APP_PROCS.get(prog_id, ()))
    try:
        # 이미 떠 있으면 실행 중인 객체를 그대로 받는다(붙기만 하고 닫지 않는다).
        return win32com.client.GetActiveObject(prog_id), False
    except Exception:
        pass
    app = win32com.client.Dispatch(prog_id)
    return app, (not running_before)


class HwpWorker(QObject):
    log = Signal(str)
    progress = Signal(int, int, str)
    done = Signal(int, int, str)
    error = Signal(str)

    def __init__(self, files: list[str], dst: str, engine: str) -> None:
        super().__init__()
        self.files = files
        self.dst = dst
        self.engine = engine
        self.stop_event = threading.Event()

    def request_stop(self) -> None:
        self.stop_event.set()

    @Slot()
    def run(self) -> None:
        pythoncom.CoInitialize()
        hwp = None
        hwp_mine = False          # 내가 띄운 한글인가(아니면 끝나도 닫지 않는다)
        stop_popup = threading.Event()
        threading.Thread(target=hwp_mod.HwpToPdfApp._popup_watcher, args=(stop_popup, 0), daemon=True).start()
        success = 0
        fail = 0
        try:
            os.makedirs(self.dst, exist_ok=True)
            total = len(self.files)
            self.log.emit(f"총 {total}개 파일 변환을 시작합니다.")
            self.progress.emit(0, total, "HWP 실행 준비 중")

            hwp, hwp_mine = com_app("HWPFrame.HwpObject")
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
            if not hwp_mine:
                self.log.emit("한글이 이미 실행 중입니다 - 그 창을 빌려 쓰고 끝나도 닫지 않습니다.")

            hwp_pid = hwp_mod.HwpToPdfApp._get_hwp_pid(hwp)
            stop_popup.set()
            stop_popup = threading.Event()
            threading.Thread(target=hwp_mod.HwpToPdfApp._popup_watcher, args=(stop_popup, hwp_pid), daemon=True).start()

            for idx, hwp_path in enumerate(self.files, 1):
                if self.stop_event.is_set():
                    self.log.emit(f"중지됨: 완료 {success}개 / 실패 {fail}개")
                    break

                filename = os.path.basename(hwp_path)
                self.progress.emit(idx - 1, total, f"{filename} 준비 중")
                pdf_name = os.path.splitext(filename)[0] + ".pdf"
                pdf_path = hwp_mod.HwpToPdfApp._unique_path(self.dst, pdf_name)
                tmp_dir = tempfile.mkdtemp(prefix="ranet_hwp_")
                try:
                    ext = os.path.splitext(hwp_path)[1].lower()
                    fmt = "HWPX" if ext == ".hwpx" else "HWP"
                    tmp_hwp = os.path.join(tmp_dir, f"src{ext}")
                    tmp_pdf = os.path.join(tmp_dir, "out.pdf")
                    shutil.copy2(hwp_path, tmp_hwp)
                    hwp.Open(tmp_hwp, fmt, "forceopen:true")

                    if self.engine == "mspdf":
                        self.log.emit(f"[진행] {filename} - Microsoft Print to PDF")
                        hwp_mod.HwpToPdfApp._save_via_ms_print_to_pdf(
                            hwp, hwp_pid, tmp_pdf, self.stop_event
                        )
                    else:
                        self.log.emit(f"[진행] {filename} - HWP 내장 PDF")
                        hwp.SaveAs(tmp_pdf, "PDF", "")
                        hwp_mod.HwpToPdfApp.wait_for_pdf_saved(tmp_pdf, self.stop_event, timeout=600)

                    # 스풀러/한글이 파일 핸들을 늦게 놓는 경우가 있어 넉넉히 재시도(최대 60초)
                    for _ in range(60):
                        try:
                            shutil.move(tmp_pdf, pdf_path)
                            break
                        except OSError:
                            time.sleep(1)
                    else:
                        raise RuntimeError("PDF 파일 이동 실패: 파일이 아직 잠겨 있습니다.")

                    success += 1
                    self.log.emit(f"[완료] {filename}")
                except Exception as exc:
                    fail += 1
                    self.log.emit(f"[실패] {filename}: {exc}")
                finally:
                    try:
                        hwp.Run("FileClose")
                    except Exception:
                        pass
                    shutil.rmtree(tmp_dir, ignore_errors=True)
                    self.progress.emit(idx, total, f"{idx} / {total} 처리 완료")

            self.done.emit(success, fail, self.dst)
        except Exception as exc:
            self.error.emit(str(exc))
        finally:
            stop_popup.set()
            # 내가 띄운 한글만 닫는다(사용자가 쓰던 창이면 그대로 둔다).
            if hwp and hwp_mine:
                try:
                    hwp.Quit()
                except Exception as exc:
                    log_err("hwp_quit", exc)
            pythoncom.CoUninitialize()


class _StopRequested(Exception):
    """[중지] 를 눌렀을 때 파일 처리 도중에 빠져나오기 위한 내부 신호."""


class HwpxWorker(QObject):
    # HWP → HWPX 일괄 교체 변환. 같은 한글 엔진의 네이티브 직렬화라 내용·서식 보존.
    # 변환본을 검증한 뒤에만 원본 .hwp를 **휴지통으로** 보낸다(복구 가능).
    log = Signal(str)
    progress = Signal(int, int, str)
    done = Signal(int, int)
    error = Signal(str)

    def __init__(self, files: list[str], delete_original: bool = True) -> None:
        super().__init__()
        self.files = files
        self.delete_original = delete_original   # True면 변환 성공한 원본을 휴지통으로
        self.stop_event = threading.Event()

    def request_stop(self) -> None:
        self.stop_event.set()

    @Slot()
    def run(self) -> None:
        pythoncom.CoInitialize()
        hwp = None
        hwp_mine = False          # 내가 띄운 한글인가(아니면 끝나도 닫지 않는다)
        stop_popup = threading.Event()
        # 한글 보안 팝업(모두 허용 등) 자동 클릭
        threading.Thread(target=hwp_mod.HwpToPdfApp._popup_watcher,
                         args=(stop_popup, 0), daemon=True).start()
        success = 0
        fail = 0
        try:
            total = len(self.files)
            tail = " (원본 .hwp는 변환 성공 시 휴지통으로 이동)" if self.delete_original else " (원본 .hwp 유지)"
            stopped = False
            self.log.emit(f"총 {total}개 파일 변환을 시작합니다.{tail}")
            self.progress.emit(0, total, "HWP 실행 준비 중")
            hwp, hwp_mine = com_app("HWPFrame.HwpObject")
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
            if not hwp_mine:
                self.log.emit("한글이 이미 실행 중입니다 - 그 창을 빌려 쓰고 끝나도 닫지 않습니다.")

            for idx, hwp_path in enumerate(self.files, 1):
                if self.stop_event.is_set():
                    self.log.emit(f"중지됨: 완료 {success}개 / 실패 {fail}개")
                    break
                filename = os.path.basename(hwp_path)
                self.progress.emit(idx - 1, total, f"{filename} 변환 중")
                # 한글 경로 문제 회피: 임시 폴더로 복사 후 열고, 임시 hwpx로 저장 → 이동
                tmp_dir = tempfile.mkdtemp(prefix="ranet_hwpx_")
                try:
                    tmp_hwp = os.path.join(tmp_dir, "src.hwp")
                    tmp_out = os.path.join(tmp_dir, "out.hwpx")
                    shutil.copy2(hwp_path, tmp_hwp)
                    if self.stop_event.is_set():
                        raise _StopRequested()
                    # 직전 한글 인스턴스 종료 직후 등 일시 오류 대비 1회 재시도
                    try:
                        hwp.Open(tmp_hwp, "HWP", "forceopen:true")
                    except Exception:
                        # 대기도 중지에 반응해야 한다(sleep 이면 2초를 그냥 흘린다).
                        if self.stop_event.wait(2):
                            raise _StopRequested()
                        # 재시도 때도 «내가 띄운 것인지» 를 다시 판단한다
                        # (사용자가 그 사이 한글을 열었을 수 있다).
                        hwp, hwp_mine = com_app("HWPFrame.HwpObject")
                        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
                        hwp.Open(tmp_hwp, "HWP", "forceopen:true")
                    hwp.SaveAs(tmp_out, "HWPX", "")
                    try:
                        hwp.Run("FileClose")
                    except Exception:
                        pass
                    # 변환본 검증: hwpx는 zip 컨테이너 — 깨졌으면 원본 보존하고 실패 처리
                    if not (os.path.isfile(tmp_out) and os.path.getsize(tmp_out) > 0
                            and zipfile.is_zipfile(tmp_out)):
                        raise RuntimeError("변환 결과가 올바른 hwpx가 아닙니다.")
                    target = _unique_path(os.path.splitext(hwp_path)[0] + ".hwpx")
                    # 한글이 저장 직후 파일을 잠깐 잡고 있을 수 있어 이동을 재시도한다.
                    for _ in range(20):
                        try:
                            shutil.move(tmp_out, target)
                            break
                        except OSError:
                            if self.stop_event.wait(0.5):
                                raise _StopRequested()
                    else:
                        raise RuntimeError("hwpx 파일 이동 실패: 파일이 아직 잠겨 있습니다.")
                    success += 1
                    if self.delete_original:
                        # 검증 통과 후에만 원본을 휴지통으로(복구 가능). 실패해도 변환은 성공.
                        try:
                            _delete_to_trash(hwp_path)
                            self.log.emit(f"[완료] {filename} → {os.path.basename(target)} (원본은 휴지통)")
                        except OSError:
                            self.log.emit(f"[완료] {filename} → {os.path.basename(target)}"
                                          f"  (원본 휴지통 이동 실패 — 원본 유지됨)")
                    else:
                        self.log.emit(f"[완료] {filename} → {os.path.basename(target)} (원본 유지)")
                except _StopRequested:
                    stopped = True
                    self.log.emit(f"[중지] {filename} 처리 중 중단 (원본 유지)")
                except Exception as exc:
                    fail += 1
                    self.log.emit(f"[실패] {filename}: {exc}  (원본 유지)")
                finally:
                    shutil.rmtree(tmp_dir, ignore_errors=True)
                    self.progress.emit(idx, total, f"{idx} / {total} 처리 완료")
                if stopped:
                    self.log.emit(f"중지됨: 완료 {success}개 / 실패 {fail}개")
                    break

            self.done.emit(success, fail)
        except Exception as exc:
            self.error.emit(str(exc))
        finally:
            stop_popup.set()
            # 내가 띄운 한글만 닫는다(사용자가 쓰던 창이면 그대로 둔다).
            if hwp and hwp_mine:
                try:
                    hwp.Quit()
                except Exception as exc:
                    log_err("hwp_quit", exc)
            pythoncom.CoUninitialize()


class PdfSplitWorker(QObject):
    log = Signal(str)
    progress = Signal(int, int, str)
    done = Signal(int, str)
    error = Signal(str)

    def __init__(self, input_path: str, mode: str, value: float) -> None:
        super().__init__()
        self.input_path = input_path
        self.mode = mode
        self.value = value

    @Slot()
    def run(self) -> None:
        try:
            count = self._split_by_page(int(self.value)) if self.mode == "page" else self._split_by_size(float(self.value))
            self.done.emit(count, os.path.dirname(self.input_path))
        except Exception as exc:
            self.error.emit(str(exc))

    def _split_by_page(self, pages_per: int) -> int:
        reader = PdfReader(self.input_path)
        total = len(reader.pages)
        base = os.path.splitext(self.input_path)[0]
        if total == 0:
            raise ValueError("PDF에 페이지가 없습니다.")
        part, start = 1, 0
        while start < total:
            end = min(start + pages_per, total)
            writer = PdfWriter()
            for i in range(start, end):
                writer.add_page(reader.pages[i])
            out_path = _unique_path(f"{base}_{part}.pdf")
            with open(out_path, "wb") as f:
                writer.write(f)
            self.log.emit(f"파일 {part}: {os.path.basename(out_path)} (페이지 {start + 1}~{end})")
            self.progress.emit(end, total, f"{end} / {total} 페이지 처리 완료")
            part += 1
            start = end
        return part - 1

    def _split_by_size(self, target_mb: float) -> int:
        target_bytes = int(target_mb * 1024 * 1024)
        reader = PdfReader(self.input_path)
        total = len(reader.pages)
        base = os.path.splitext(self.input_path)[0]
        if total == 0:
            raise ValueError("PDF에 페이지가 없습니다.")
        test_w = PdfWriter()
        test_w.add_page(reader.pages[0])
        if _get_writer_size(test_w) > target_bytes:
            raise ValueError("단일 페이지 크기가 목표 용량보다 큽니다. 목표 용량을 높여주세요.")

        self.log.emit(f"목표 용량: {target_mb} MB | 총 {total} 페이지")
        part, start = 1, 0
        while start < total:
            accepted = start
            i = start
            while i < total:
                tw = PdfWriter()
                for j in range(start, i + 1):
                    tw.add_page(reader.pages[j])
                if _get_writer_size(tw) > target_bytes:
                    break
                accepted = i + 1
                i += 1

            writer = PdfWriter()
            for j in range(start, accepted):
                writer.add_page(reader.pages[j])
            if len(writer.pages) == 0:
                raise RuntimeError(f"페이지 {start + 1}을 단독으로 저장할 수 없습니다.")

            end = start + len(writer.pages)
            out_path = _unique_path(f"{base}_{part}.pdf")
            with open(out_path, "wb") as f:
                writer.write(f)
            actual_mb = os.path.getsize(out_path) / 1024 / 1024
            self.log.emit(f"파일 {part}: {os.path.basename(out_path)} (페이지 {start + 1}~{end}, {actual_mb:.2f} MB)")
            self.progress.emit(end, total, f"{end} / {total} 페이지 처리 완료")
            part += 1
            start = end
        return part - 1


class PdfCompressWorker(QObject):
    log = Signal(str)
    progress = Signal(int, int, str)
    done = Signal(str, float, float, float, int, int)
    error = Signal(str)

    def __init__(self, input_path: str, quality_name: str, jpeg_q: int, target_dpi: int) -> None:
        super().__init__()
        self.input_path = input_path
        self.quality_name = quality_name
        self.jpeg_q = jpeg_q
        self.target_dpi = target_dpi

    @Slot()
    def run(self) -> None:
        try:
            before_mb = os.path.getsize(self.input_path) / 1024 / 1024
            base = os.path.splitext(self.input_path)[0]
            out_path = _unique_path(f"{base}_압축.pdf")
            self.log.emit("=== 압축 시작 ===")
            self.log.emit(f"압축 강도: {self.quality_name} (JPEG quality={self.jpeg_q}, 목표 DPI={self.target_dpi})")
            self.log.emit(f"원본 크기: {before_mb:.2f} MB")

            reader = PdfReader(self.input_path)
            total = len(reader.pages)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)

            total_found = 0
            total_done = 0
            for page_num, page in enumerate(writer.pages, 1):
                found, done = self._compress_page_images(page)
                total_found += found
                total_done += done
                self.log.emit(f"페이지 {page_num:3d}: 이미지 {found}개 발견, {done}개 압축 완료")
                self.progress.emit(page_num, total, f"{page_num} / {total} 페이지 처리")

            with open(out_path, "wb") as f:
                writer.write(f)
            after_mb = os.path.getsize(out_path) / 1024 / 1024
            pct = (1 - after_mb / before_mb) * 100 if before_mb > 0 else 0
            self.done.emit(out_path, before_mb, after_mb, pct, total_done, total_found)
        except Exception as exc:
            self.error.emit(str(exc))

    def _compress_page_images(self, page):
        found = 0
        done = 0
        try:
            resources = page.get("/Resources")
            if not resources:
                return found, done
            resources = resources.get_object()
            xobjects = resources.get("/XObject")
            if not xobjects:
                return found, done
            xobjects = xobjects.get_object()
        except Exception:
            return found, done

        for key in list(xobjects.keys()):
            try:
                xobj = xobjects[key].get_object()
                if xobj.get("/Subtype") != "/Image" or xobj.get("/ImageMask") is True:
                    continue
                found += 1
                width = int(xobj.get("/Width", 0))
                height = int(xobj.get("/Height", 0))
                if width == 0 or height == 0:
                    continue

                try:
                    raw = xobj.get_data()
                except Exception:
                    continue

                img = None
                try:
                    img = Image.open(io.BytesIO(raw))
                    img.load()
                except Exception:
                    try:
                        cs = xobj.get("/ColorSpace", "/DeviceRGB")
                        if hasattr(cs, "get_object"):
                            cs = cs.get_object()
                        if isinstance(cs, list):
                            cs = cs[0] if cs else "/DeviceRGB"
                        mode = "L" if str(cs) == "/DeviceGray" else "RGB"
                        if int(xobj.get("/BitsPerComponent", 8)) == 8:
                            img = Image.frombytes(mode, (width, height), raw)
                    except Exception as e:
                        log_err("_compress_page_images.rebuild", e)
                if img is None:
                    continue

                if max(width, height) > self.target_dpi * 10:
                    scale = self.target_dpi * 10 / max(width, height)
                    img = img.resize((max(1, int(width * scale)), max(1, int(height * scale))), Image.LANCZOS)

                if img.mode == "RGBA":
                    bg = Image.new("RGB", img.size, (255, 255, 255))
                    bg.paste(img, mask=img.split()[3])
                    img = bg
                elif img.mode in ("P", "CMYK") or img.mode not in ("RGB", "L"):
                    img = img.convert("RGB")

                buf = io.BytesIO()
                img.save(buf, format="JPEG", quality=self.jpeg_q, optimize=True)
                jpeg_bytes = buf.getvalue()
                xobj._data = jpeg_bytes
                xobj[NameObject("/Filter")] = NameObject("/DCTDecode")
                xobj[NameObject("/Length")] = NumberObject(len(jpeg_bytes))
                xobj[NameObject("/Width")] = NumberObject(img.width)
                xobj[NameObject("/Height")] = NumberObject(img.height)
                for unwanted in ("/DecodeParms", "/SMask"):
                    if NameObject(unwanted) in xobj:
                        del xobj[NameObject(unwanted)]
                done += 1
            except Exception:
                continue
        return found, done


# 라넷툴킷이 '주석 모드'로 박은 쪽수 표식. 이 title을 가진 FreeText 주석만
# 나중에 안전하게 찾아 삭제·재번호할 수 있다(본문 손상 없음).
RTK_PAGENUM_TAG = "RTK_PAGENUM"


class PdfMergeWorker(QObject):
    # 선택된 파일들을 순서대로 PDF로 변환한 뒤 하나로 병합한다.
    # pdf는 그대로, 이미지/word/ppt/hwp는 각각 PDF로 변환 후 페이지를 이어붙인다.
    log = Signal(str)
    progress = Signal(int, int, str)
    done = Signal(str, int, int)   # out_path, 성공 파일 수, 전체 파일 수
    error = Signal(str)

    IMG_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
    WORD_EXTS = {".doc", ".docx", ".rtf"}
    PPT_EXTS = {".ppt", ".pptx"}
    HWP_EXTS = {".hwp", ".hwpx"}

    # A4 크기(pt). 사진을 그대로 PDF로 만든 원본 등 A4보다 큰 페이지를 이 크기로 맞춘다.
    A4_W_PT = 595.32
    A4_H_PT = 841.92

    def __init__(self, files: list[str], out_path: str, page_numbers: bool = False,
                 font_size: int = 11, position: str = "center",
                 nups: list[int] | None = None, force_a4: bool = True,
                 bake_numbers: bool = False) -> None:
        super().__init__()
        self.files = files
        self.force_a4 = force_a4   # A4 초과 페이지를 A4로 축소(비율 유지·중앙 정렬)
        # 파일별 페이지 배치 코드. 1=그대로, 2/4=모아찍기, -2/-4=모아찍기 풀기.
        # 길이가 안 맞으면 부족분은 1로 채운다.
        self.nups = list(nups) if nups else [1] * len(files)
        if len(self.nups) < len(files):
            self.nups += [1] * (len(files) - len(self.nups))
        self.out_path = out_path
        self.page_numbers = page_numbers   # 병합 후 1쪽부터 - 1 - 형식 번호 부여
        self.font_size = font_size         # 페이지 번호 글자 크기(pt)
        self.position = position           # 하단 위치: left / center / right
        # 기본(False)=지울 수 있는 주석으로 쪽수 표시 / True=본문에 직접 새김(영구·제거 불가)
        self.bake_numbers = bake_numbers
        self.stop_event = threading.Event()
        # COM 앱 핸들은 필요할 때 한 번만 띄우고 끝까지 재사용한다.
        self._word = None
        self._ppt = None
        self._hwp = None
        self._hwp_pid = 0
        self._stop_popup = None
        self._stop_hide = None

    def request_stop(self) -> None:
        self.stop_event.set()

    @Slot()
    def run(self) -> None:
        pythoncom.CoInitialize()
        tmp_root = tempfile.mkdtemp(prefix="ranet_merge_")
        try:
            writer = PdfWriter()
            total = len(self.files)
            success = 0
            self.log.emit(f"총 {total}개 파일 병합을 시작합니다.")
            for idx, src in enumerate(self.files, 1):
                if self.stop_event.is_set():
                    self.log.emit(f"중지됨: {success}개 처리 후 중단")
                    break
                name = os.path.basename(src)
                self.progress.emit(idx - 1, total, f"{name} 변환 중")
                ext = os.path.splitext(src)[1].lower()
                nup = self.nups[idx - 1] if idx - 1 < len(self.nups) else 1
                try:
                    pdf_path = self._to_pdf(src, ext, tmp_root, idx)
                    if pdf_path is None:
                        self.log.emit(f"[건너뜀] {name}: 지원하지 않는 형식")
                    else:
                        if not self._is_valid_pdf(pdf_path):
                            raise RuntimeError(
                                "정상 PDF가 아닙니다 (파일 손상 또는 OneDrive 미동기화 스텁일 수 있음). "
                                "원본을 다시 받아 주세요."
                            )
                        reader = PdfReader(pdf_path)
                        src_pages = list(reader.pages)
                        if self.force_a4:
                            src_pages, fixed = self._normalize_oversized_pages(src_pages)
                            if fixed:
                                self.log.emit(f"[크기조정] {name}: A4 초과 {fixed}쪽을 A4로 축소")
                        if nup != 1:
                            # 목표대로 nup쪽을 한 면에 모아찍기(1→2/4)
                            out_pages = self._arrange_pages(src_pages, nup)
                            for page in out_pages:
                                writer.add_page(page)
                            self.log.emit(
                                f"[완료] {name} ({len(src_pages)}쪽 → {len(out_pages)}쪽, 1면 {nup}페이지)"
                            )
                        else:
                            for page in src_pages:
                                writer.add_page(page)
                            self.log.emit(f"[완료] {name} ({len(src_pages)}페이지)")
                        success += 1
                except Exception as exc:
                    self.log.emit(f"[실패] {name}: {exc}")
                self.progress.emit(idx, total, f"{idx} / {total} 처리 완료")

            if len(writer.pages) == 0:
                raise RuntimeError("병합할 수 있는 페이지가 없습니다.")

            if self.page_numbers:
                if self.bake_numbers:
                    self._stamp_page_numbers(writer)         # 본문에 직접 새김
                else:
                    self._normalize_page_rotations(writer)   # 주석 위치 정확하게(회전 흡수)

            os.makedirs(os.path.dirname(self.out_path), exist_ok=True)
            with open(self.out_path, "wb") as f:
                writer.write(f)
            if self.page_numbers and not self.bake_numbers:
                self._annotate_page_numbers(self.out_path)   # 지울 수 있는 주석 쪽수
            self.done.emit(self.out_path, success, total)
        except Exception as exc:
            self.error.emit(str(exc))
        finally:
            self._cleanup_apps()
            shutil.rmtree(tmp_root, ignore_errors=True)
            pythoncom.CoUninitialize()

    def _to_pdf(self, src: str, ext: str, tmp_root: str, idx: int) -> str | None:
        # 형식별로 PDF 경로를 돌려준다. pdf는 원본 경로를 그대로 사용한다.
        if ext == ".pdf":
            return src
        out_pdf = os.path.join(tmp_root, f"{idx}.pdf")
        if ext in self.IMG_EXTS:
            self._image_to_pdf(src, out_pdf)
        elif ext in self.WORD_EXTS:
            self._word_to_pdf(src, out_pdf)
        elif ext in self.PPT_EXTS:
            self._ppt_to_pdf(src, out_pdf)
        elif ext in self.HWP_EXTS:
            self._hwp_to_pdf(src, ext, out_pdf)
        else:
            return None
        return out_pdf

    @staticmethod
    def _is_valid_pdf(path: str) -> bool:
        # 앞부분에 %PDF 헤더가 있으면 정상 PDF로 본다.
        # 손상 파일·OneDrive 온라인전용 스텁(수백 바이트, 헤더 없음)을 걸러낸다.
        try:
            with open(path, "rb") as f:
                head = f.read(1024)
        except OSError:
            return False
        return b"%PDF" in head

    @classmethod
    def _normalize_oversized_pages(cls, pages: list) -> tuple[list, int]:
        """A4보다 큰 페이지를 방향(가로/세로) 유지한 A4 크기로 축소한다.

        사진을 그대로 PDF로 변환한 원본은 페이지가 A4의 수 배가 되어 병합본
        페이지 크기가 들쭉날쭉해진다. 비율 유지 + 중앙 정렬로 A4 안에 맞춘다.
        반환: (정규화된 페이지 목록, 축소한 페이지 수). 5% 이내 초과는 허용.
        """

        out: list = []
        fixed = 0
        for pg in pages:
            try:
                pw = float(pg.mediabox.width)
                ph = float(pg.mediabox.height)
            except Exception:
                out.append(pg)
                continue
            # 원본 방향에 맞춘 A4 한도(세로 원본→세로 A4, 가로 원본→가로 A4)
            lw, lh = (cls.A4_W_PT, cls.A4_H_PT) if ph >= pw else (cls.A4_H_PT, cls.A4_W_PT)
            if pw <= lw * 1.05 and ph <= lh * 1.05:
                out.append(pg)
                continue
            try:
                scale = min(lw / pw, lh / ph)
                blank = PageObject.create_blank_page(width=lw, height=lh)
                tx = (lw - pw * scale) / 2 - float(pg.mediabox.left) * scale
                ty = (lh - ph * scale) / 2 - float(pg.mediabox.bottom) * scale
                blank.merge_transformed_page(pg, Transformation().scale(scale).translate(tx, ty))
                out.append(blank)
                fixed += 1
            except Exception:
                out.append(pg)   # 실패하면 원본 그대로(병합은 계속)
        return out, fixed

    @staticmethod
    def _arrange_pages(pages: list, n: int) -> list:
        # n==1: 그대로, n>=2: n쪽을 한 면에 모아찍기.
        if n <= 1:
            return list(pages)
        return PdfMergeWorker._nup_pages(pages, n)

    @staticmethod
    def _nup_pages(pages: list, n: int) -> list:
        # 원본 여러 쪽을 한 면에 모아 배치한 새 PageObject 목록을 돌려준다.
        # 원본 방향에 맞춰 배치(강제 리사이즈 없음, 비율 유지·중앙 정렬):
        #  - 2쪽: 세로 원본 → 가로 면에 좌·우, 가로 원본 → 세로 면에 위·아래
        #  - 4쪽: 원본 방향 유지한 면에 2x2
        if n <= 1:
            return list(pages)
        out: list = []
        for start in range(0, len(pages), n):
            chunk = pages[start:start + n]
            ref = chunk[0]
            rw = float(ref.mediabox.width)
            rh = float(ref.mediabox.height)
            portrait = rh >= rw   # 원본이 세로 방향인가
            if n == 2:
                # 면을 원본과 반대 방향으로 눕히고, 세로원본은 좌우/가로원본은 상하로 나눈다.
                sheet_w, sheet_h = rh, rw
                cols, rows = (2, 1) if portrait else (1, 2)
            else:  # 4쪽: 원본 방향 유지 2x2
                sheet_w, sheet_h = rw, rh
                cols, rows = 2, 2
            blank = PageObject.create_blank_page(width=sheet_w, height=sheet_h)
            cell_w = sheet_w / cols
            cell_h = sheet_h / rows
            for j, pg in enumerate(chunk):
                col = j % cols
                row = j // cols
                pw = float(pg.mediabox.width)
                ph = float(pg.mediabox.height)
                scale = min(cell_w / pw, cell_h / ph)
                dw, dh = pw * scale, ph * scale
                x0 = col * cell_w
                y0 = sheet_h - (row + 1) * cell_h   # 행은 위→아래로 채움
                tx = x0 + (cell_w - dw) / 2 - float(pg.mediabox.left) * scale
                ty = y0 + (cell_h - dh) / 2 - float(pg.mediabox.bottom) * scale
                blank.merge_transformed_page(pg, Transformation().scale(scale).translate(tx, ty))
            out.append(blank)
        return out

    @staticmethod
    def _image_to_pdf(src: str, out_pdf: str) -> None:
        # 이미지를 RGB로 변환(투명/팔레트 대응) 후 1페이지 PDF로 저장.
        img = Image.open(src)
        if img.mode == "RGBA":
            bg = Image.new("RGB", img.size, (255, 255, 255))
            bg.paste(img, mask=img.split()[3])
            img = bg
        elif img.mode != "RGB":
            img = img.convert("RGB")
        img.save(out_pdf, "PDF", resolution=150.0)

    def _word_to_pdf(self, src: str, out_pdf: str) -> None:
        # Word COM으로 열어 PDF(FileFormat=17, wdFormatPDF)로 저장.
        if self._word is None:
            self._word, self._word_mine = com_app("Word.Application")
            if self._word_mine:          # 내가 띄운 창에만 설정을 건드린다
                self._word.Visible = False
                self._word.DisplayAlerts = 0
            else:
                self.log.emit("워드가 이미 실행 중입니다 - 끝나도 닫지 않습니다.")
        doc = self._word.Documents.Open(os.path.abspath(src), ReadOnly=True)
        try:
            doc.SaveAs(os.path.abspath(out_pdf), FileFormat=17)
        finally:
            doc.Close(False)

    def _ppt_to_pdf(self, src: str, out_pdf: str) -> None:
        # PowerPoint COM으로 열어 PDF(ppSaveAsPDF=32)로 저장.
        if self._ppt is None:
            self._ppt, self._ppt_mine = com_app("PowerPoint.Application")
            if not self._ppt_mine:
                self.log.emit("파워포인트가 이미 실행 중입니다 - 끝나도 닫지 않습니다.")
        # 창 없이 열기(WithWindow=False). 일부 환경은 Visible=False가 막혀 인자로 처리.
        pres = self._ppt.Presentations.Open(
            os.path.abspath(src), ReadOnly=True, WithWindow=False
        )
        try:
            pres.SaveAs(os.path.abspath(out_pdf), 32)
        finally:
            pres.Close()

    def _hwp_to_pdf(self, src: str, ext: str, out_pdf: str) -> None:
        # 한글 COM으로 열어 내장 PDF 엔진으로 저장. (HwpWorker와 동일 방식)
        if self._hwp is None:
            self._stop_popup = threading.Event()
            threading.Thread(
                target=hwp_mod.HwpToPdfApp._popup_watcher,
                args=(self._stop_popup, 0),
                daemon=True,
            ).start()
            self._hwp, self._hwp_mine = com_app("HWPFrame.HwpObject")
            self._hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
            if not self._hwp_mine:
                self.log.emit("한글이 이미 실행 중입니다 - 끝나도 닫지 않습니다.")
            self._hwp_pid = hwp_mod.HwpToPdfApp._get_hwp_pid(self._hwp)
            self._stop_popup.set()
            self._stop_popup = threading.Event()
            threading.Thread(
                target=hwp_mod.HwpToPdfApp._popup_watcher,
                args=(self._stop_popup, self._hwp_pid),
                daemon=True,
            ).start()
            # 한글 '인쇄 중' 진행창을 숨기는 워커(보안 팝업은 건드리지 않음).
            self._stop_hide = threading.Event()
            threading.Thread(
                target=self._hwp_window_hider,
                args=(self._stop_hide, self._hwp_pid),
                daemon=True,
            ).start()
        fmt = "HWPX" if ext == ".hwpx" else "HWP"
        # 한글 경로 문제 회피: 임시 폴더로 복사 후 열고, 임시 PDF로 저장 → 이동.
        tmp_dir = tempfile.mkdtemp(prefix="ranet_hwp_")
        try:
            tmp_hwp = os.path.join(tmp_dir, f"src{ext}")
            tmp_pdf = os.path.join(tmp_dir, "out.pdf")
            shutil.copy2(src, tmp_hwp)
            self._hwp.Open(tmp_hwp, fmt, "forceopen:true")
            # 'Hancom PDF' 프린터로 기본 인쇄(1면 1쪽) → 문서에 저장된 '모아 찍기'를
            # 무시하고 항상 1-up으로 내보낸다. SaveAs("PDF")는 숨은 모아찍기 설정을
            # 따라가 2-up이 되므로 이 방식을 우선 사용한다(실패 시 SaveAs로 폴백).
            if not self._hwp_print_to_pdf(tmp_pdf):
                self._hwp.SaveAs(tmp_pdf, "PDF", "")
                hwp_mod.HwpToPdfApp.wait_for_pdf_saved(tmp_pdf, self.stop_event, timeout=60)
            try:
                self._hwp.Run("FileClose")
            except Exception:
                pass
            # Hancom PDF 스풀러가 파일을 늦게 놓아줄 수 있어 이동을 재시도한다.
            self._move_with_retry(tmp_pdf, out_pdf)
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    def _move_with_retry(self, src_pdf: str, out_pdf: str, attempts: int = 30) -> None:
        # 파일 잠금(PermissionError/OSError)이 풀릴 때까지 짧게 재시도하며 이동한다.
        for _ in range(attempts):
            try:
                shutil.move(src_pdf, out_pdf)
                return
            except (PermissionError, OSError):
                if self.stop_event.is_set():
                    break
                time.sleep(0.5)
        raise RuntimeError("PDF 파일이 잠겨 있어 이동하지 못했습니다 (다른 프로그램이 사용 중).")

    def _hwp_print_to_pdf(self, out_pdf: str) -> bool:
        # CreateAction("Print") + 'Hancom PDF' 프린터로 1면 1쪽(기본 인쇄) PDF 생성.
        # PrintMethod=0(모아찍기 해제)이 먹어서 모아찍기 박힌 문서도 1-up으로 나온다.
        # 'Hancom PDF' 프린터가 없거나 실패하면 False를 돌려 SaveAs로 폴백한다.
        try:
            act = self._hwp.CreateAction("Print")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("PrinterName", "Hancom PDF")
            pset.SetItem("PrintMethod", 0)   # 기본 인쇄(모아 찍기 해제)
            pset.SetItem("PageLayout", 0)
            pset.SetItem("PrintToFile", 1)
            pset.SetItem("FileName", os.path.abspath(out_pdf))
            pset.SetItem("Pause", 0)
            act.Execute(pset)
            # 파일 생성·잠금해제·크기안정까지 확인(스풀러가 다 쓸 때까지 대기).
            hwp_mod.HwpToPdfApp.wait_for_pdf_saved(out_pdf, self.stop_event, timeout=90)
            return True
        except Exception as exc:
            # Hancom PDF 프린터 없음/실패 → False 반환해 SaveAs로 폴백.
            self.log.emit(f"[한글] 인쇄 방식 변환 불가, 내장 저장으로 대체: {exc}")
            return False

    # Hancom PDF 인쇄 시 잠깐 뜨는 '인쇄 출력 저장' 대화창 제목들(FileName이 자동 채워져 곧 닫힘).
    _HWP_SAVE_TITLES = (
        "다음 이름으로 프린터 출력 저장",
        "Save Print Output As",
        "인쇄 출력을 다른 이름으로 저장",
    )

    @staticmethod
    def _hwp_window_hider(stop_event: threading.Event, hwp_pid: int) -> None:
        # 한글 '인쇄 중' 진행창과 '인쇄 출력 저장' 대화창을 즉시 숨긴다.
        # 버튼 있는 보안 팝업(모두 허용)은 건드리지 않는다(팝업 워처가 클릭해야 하므로).
        def cb(hwnd, _param) -> bool:
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                title = win32gui.GetWindowText(hwnd)
                cls = win32gui.GetClassName(hwnd)
                _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
                # Windows가 붙이는 '(응답 없음)'/'(Not Responding)' 꼬리표 제거 후 비교
                norm = (title.replace(" (응답 없음)", "")
                             .replace(" (Not Responding)", "").strip())
                hide = False
                if norm in PdfMergeWorker._HWP_SAVE_TITLES or norm.startswith("인쇄 중"):
                    hide = True   # 인쇄 진행창/저장창(제목 특정적 → pid 무관 안전)
                elif pid == hwp_pid and ("인쇄" in title or cls.startswith("HwndWrapper[Hwp.exe")):
                    hide = True   # 한글 '인쇄 중' 진행창(추가 보강)
                if hide:
                    win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
            except Exception:
                pass
            return True

        while not stop_event.is_set():
            try:
                win32gui.EnumWindows(cb, None)
            except Exception:
                pass
            time.sleep(0.2)   # 0.05→0.2: 일괄(장시간) 작업 중 GIL 경합·CPU 낮춤

    def _cleanup_apps(self) -> None:
        # 띄운 COM 앱들을 정리한다. 실패는 무시.
        if self._stop_popup:
            self._stop_popup.set()
        if self._stop_hide:
            self._stop_hide.set()
        # ⚠️ 내가 띄운 앱만 닫는다. 사용자가 쓰던 워드·파워포인트·한글을 닫으면
        #    저장 안 한 편집분이 그대로 날아간다(자동화 세션이라 확인 창도 안 뜬다).
        for app, mine in ((self._word, getattr(self, "_word_mine", False)),
                          (self._ppt, getattr(self, "_ppt_mine", False))):
            try:
                if app is not None and mine:
                    app.Quit()
            except Exception as exc:
                log_err("office_quit", exc)
        try:
            if self._hwp is not None and getattr(self, "_hwp_mine", False):
                self._hwp.Quit()
        except Exception as exc:
            log_err("hwp_quit", exc)

    def _normalize_page_rotations(self, writer: PdfWriter) -> None:
        # /Rotate 플래그(가로 스캔을 90°로 세로처럼 보이게 한 경우 등)를 콘텐츠에 구워넣어
        # mediabox=시각크기·/Rotate=0 으로 정규화한다. 그래야 본문/주석 쪽수가 항상
        # '보이는 하단'에 바르게 찍힌다(주석 모드 포함 공통 전처리).
        for page in writer.pages:
            try:
                if page.rotation:
                    page.transfer_rotation_to_content()
            except Exception as e:
                log_err("_normalize_page_rotations", e)

    def _annotate_page_numbers(self, path: str) -> None:
        # 쪽수를 '지울 수 있는 주석(FreeText)'으로 넣는다. title=RTK_PAGENUM_TAG 표식을 달아
        # 나중에 그 주석만 찾아 삭제·재번호할 수 있다(본문 손상 0). appearance stream이
        # 생성돼 대부분 뷰어·인쇄에서 동일하게 보인다. fitz는 회전 반영 좌표로 동작한다.
        try:
            doc = fitz.open(path)
        except Exception as exc:
            self.log.emit(f"[번호 경고] 주석 쪽수 열기 실패 → 건너뜀: {exc}")
            return
        try:
            total = doc.page_count
            self.log.emit(f"쪽수 주석 넣는 중... (총 {total}쪽, {self.font_size}pt, 지움 가능)")
            size = float(self.font_size)
            margin = 40.0
            for i in range(total):
                try:
                    page = doc.load_page(i)
                    # 이미 라넷툴킷 주석 쪽수가 있으면(예: 예전에 주석으로 번호 매긴 파일을
                    # 다시 병합) 먼저 떼어내 중복을 막고 새 번호로 교체한다. 본문에 구워진
                    # 옛 번호는 표식이 없어 여기서 못 지운다(주석으로 박은 것만 인식).
                    for an in list(page.annots() or []):
                        if (an.info or {}).get("title") == RTK_PAGENUM_TAG:
                            page.delete_annot(an)
                    text = f"- {i + 1} -"
                    r = page.rect                      # 회전 반영된 표시 영역
                    tw = fitz.get_text_length(text, fontname="hebo", fontsize=size)
                    box_w = tw + size * 1.4
                    box_h = size * 1.8
                    if self.position == "left":
                        cx = r.x0 + margin + box_w / 2.0
                    elif self.position == "right":
                        cx = r.x1 - margin - box_w / 2.0
                    else:
                        cx = (r.x0 + r.x1) / 2.0
                    cy = r.y1 - (size + 9.0)            # 하단 여백(본문판과 동일 느낌)
                    rect = fitz.Rect(cx - box_w / 2, cy - box_h / 2,
                                     cx + box_w / 2, cy + box_h / 2)
                    annot = page.add_freetext_annot(
                        rect, text, fontsize=size, fontname="hebo",
                        text_color=(0.86, 0.10, 0.10), fill_color=(1, 1, 1),
                        align=fitz.TEXT_ALIGN_CENTER)
                    # 표식: 나중에 이 title로 찾아 삭제·재번호. content엔 쪽번호 보관.
                    annot.set_info(title=RTK_PAGENUM_TAG, content=str(i + 1))
                    annot.set_border(width=0)
                    annot.update(text_color=(0.86, 0.10, 0.10), fill_color=(1, 1, 1))
                except Exception as exc:
                    self.log.emit(f"[번호 경고] {i + 1}쪽 주석 실패: {exc}")
            doc.save(path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
            self.log.emit(f"쪽수 주석 처리 완료 (총 {total}쪽, 지움 가능)")
        finally:
            doc.close()

    def _stamp_page_numbers(self, writer: PdfWriter) -> None:
        # 병합된 전체 PDF에 1쪽부터 -1- 형식 번호를 페이지 중앙 하단에 찍는다.
        # 흰 박스 위 붉은 굵은 글자. reportlab으로 페이지별 오버레이를 만들어 합친다.
        from reportlab.pdfgen import canvas   # 지연 임포트

        total = len(writer.pages)
        self.log.emit(f"페이지 번호 매기는 중... (총 {total}쪽, {self.font_size}pt)")
        for i, page in enumerate(writer.pages, 1):
            try:
                # 원본이 /Rotate 플래그를 가진 페이지(가로 스캔을 90°로 세로처럼 보이게 한 경우 등)는
                # mediabox가 '회전 전' 좌표라, 그대로 하단에 찍으면 뷰어가 회전 적용 시 번호가
                # 옆/위로 간다. 회전을 콘텐츠에 구워넣어 /Rotate=0·mediabox=시각크기로 정규화한 뒤
                # 찍으면 항상 '시각적 하단'에 바르게 찍힌다.
                if page.rotation:
                    page.transfer_rotation_to_content()
                w = float(page.mediabox.width)
                h = float(page.mediabox.height)
                # '- 1 -' 처럼 하이픈과 숫자 사이에 공백 한 칸씩.
                overlay = self._make_number_overlay(canvas, w, h, f"- {i} -", self.font_size, self.position)
                page.merge_page(overlay)
            except Exception as exc:
                self.log.emit(f"[번호 경고] {i}쪽 번호 표기 실패: {exc}")
        self.log.emit(f"페이지 번호 처리 완료 (총 {total}쪽)")

    @staticmethod
    def _make_number_overlay(canvas_mod, w: float, h: float, text: str, size: int = 11,
                             position: str = "center"):
        # 페이지 크기에 맞춘 1쪽짜리 오버레이 PDF의 PageObject를 돌려준다.
        buf = io.BytesIO()
        c = canvas_mod.Canvas(buf, pagesize=(w, h))
        font = "Helvetica-Bold"
        c.setFont(font, size)
        tw = c.stringWidth(text, font, size)
        baseline = size + 9.0          # 글자 크기에 맞춰 하단 여백 확보
        pad_x = size * 0.6
        pad_y = size * 0.4
        margin = 40.0                  # 좌/우 정렬 시 페이지 가장자리 여백
        if position == "left":
            cx = margin + pad_x + tw / 2.0
        elif position == "right":
            cx = w - margin - pad_x - tw / 2.0
        else:
            cx = w / 2.0
        # 흰 박스(둥근 모서리)
        c.setFillColorRGB(1, 1, 1)
        c.roundRect(cx - tw / 2 - pad_x, baseline - pad_y,
                    tw + 2 * pad_x, size + 2 * pad_y, 3, fill=1, stroke=0)
        # 붉은 굵은 글자
        c.setFillColorRGB(0.86, 0.10, 0.10)
        c.drawCentredString(cx, baseline, text)
        c.save()
        buf.seek(0)
        return PdfReader(buf).pages[0]


# HWPX 결과보고서 생성부(약 430줄)는 rtk_hwpx_report.py로 옮겼다. 동작은 그대로다.
from rtk_hwpx_report import build_hwpx_report


class PdfBatchMergeWorker(PdfMergeWorker):
    # [폴더명으로 일괄병합] 워커. 최상위 폴더의 각 자식 폴더를 PDF 1개로 병합한다.
    # 변환/모아찍기/페이지번호 로직은 PdfMergeWorker 것을 그대로 재사용하고,
    # COM 앱(한글/Word/PPT)은 전체 폴더에 걸쳐 한 번만 띄워 재사용한다(폴더마다 재기동 X).
    # 추가: 변환 후 이미지 깨짐 검사(fitz) + 결과 보고서(.hwp) 생성.
    folder_progress = Signal(int, int, str)               # 현재 폴더 인덱스, 전체, 상태문구
    batch_done = Signal(str, int, int, str)               # 출력폴더, 성공 폴더수, 전체 폴더수, 보고서경로

    def __init__(self, tasks: list[tuple[str, str, list[str], list[str]]], out_dir: str,
                 page_numbers: bool, font_size: int, position: str,
                 layout_mode: str, special_exts: set[str],
                 make_report: bool, root_dir: str, force_a4: bool = True,
                 bake_numbers: bool = False) -> None:
        # tasks: [(표시명(상대경로), 단위폴더경로, [병합파일...], [미병합(엑셀)파일...]), ...]
        # layout_mode: "all1"=일괄 1페이지 / "all2"=일괄 2페이지 / "special2"=특정 유형만 2페이지
        # special_exts: special2일 때 2페이지로 만들 확장자 집합
        super().__init__([], "", page_numbers, font_size, position, None,
                         force_a4=force_a4, bake_numbers=bake_numbers)
        self.tasks = tasks
        self.out_dir = out_dir
        self.layout_mode = layout_mode
        self.special_exts = special_exts
        self.make_report = make_report
        self.root_dir = root_dir
        # 별도 프로세스(CLI)에서 stdout으로 직접 진행을 흘릴 콜백. 본 스레드가 COM
        # 호출에 묶여 있어도 하트비트 스레드가 이걸로 '살아있음'을 보낸다. None이면
        # 일반 Qt 시그널(folder_progress)로 폴백.
        self._progress_sink = None

    def _heartbeat(self, stop_evt, c: int, t: int, fname: str, bn: str) -> None:
        # 변환이 길어도 멈춘 게 아님을 보이려 2초마다 경과 시간을 진행 상태로 흘린다.
        t0 = time.time()
        while not stop_evt.wait(2.0):
            el = int(time.time() - t0)
            msg = f"[{fname}] {bn} 변환 중… ({el}초 경과)"
            sink = self._progress_sink
            if sink:
                try:
                    sink(c, t, msg)
                except Exception:
                    pass
            else:
                self.folder_progress.emit(c, t, msg)

    def _read_pdf_pages(self, pdf_path: str, tmp_root: str, idx: int) -> list:
        # PDF 페이지 목록을 읽는다. pypdf가 손상(예: 'Trailer cannot be read')으로
        # 못 읽으면 fitz로 재저장해 구조를 복구한 뒤 다시 시도한다.
        try:
            return list(PdfReader(pdf_path).pages)
        except Exception:
            fixed = os.path.join(tmp_root, f"fix{idx}.pdf")
            doc = fitz.open(pdf_path)
            try:
                doc.save(fixed)   # 정상 구조로 재기록(xref/trailer 복구)
            finally:
                doc.close()
            return list(PdfReader(fixed).pages)

    def _reset_hwp(self) -> None:
        # 죽었거나 응답 없는 한글 COM 핸들을 정리한다(다음 변환 때 새로 띄움).
        if self._stop_popup:
            self._stop_popup.set()
        if self._stop_hide:
            self._stop_hide.set()
        try:
            if self._hwp is not None and getattr(self, "_hwp_mine", False):
                self._hwp.Quit()    # 내가 띄운 것만 닫는다
        except Exception as exc:
            log_err("hwp_quit", exc)
        self._hwp = None
        self._hwp_mine = False
        self._hwp_pid = 0

    def _hwp_to_pdf(self, src: str, ext: str, out_pdf: str) -> None:
        # 한글 변환 전략:
        #  - 1차: 인쇄(Hancom PDF) → 모아찍기 문서도 1-up으로 풀어 출력(권장).
        #  - 인쇄가 통째로 실패(예: 큰 문서에서 RPC '서버 프로세스 호출 안됨'으로 한글이
        #    죽는 경우)하면 → 핸들 리셋 후 PDF 저장(SaveAs)으로 1회 재시도.
        #  - SaveAs는 문서에 박힌 모아찍기를 따라가 2-up이 될 수 있음(감수).
        self._cur_hwp_force_saveas = False   # 1차는 인쇄
        try:
            super()._hwp_to_pdf(src, ext, out_pdf)
            return
        except Exception as exc:
            self._reset_hwp()   # 죽었을 수 있는 핸들 정리(새 한글로 재시도)
            self.log.emit(f"[한글] 1차(인쇄) 실패 → PDF 저장(SaveAs)으로 1회 재시도: {exc}")
        # 2차: SaveAs 방식으로 강제 재시도.
        self._cur_hwp_force_saveas = True
        super()._hwp_to_pdf(src, ext, out_pdf)

    def _hwp_print_to_pdf(self, out_pdf: str) -> bool:
        # SaveAs 모드(2차 재시도)면 인쇄를 건너뛰고 한글 내장 PDF 저장(SaveAs)으로 만든다.
        if getattr(self, "_cur_hwp_force_saveas", False):
            self.log.emit("[한글] PDF 저장(SaveAs) 방식으로 변환")
            self._hwp.SaveAs(os.path.abspath(out_pdf), "PDF", "")
            hwp_mod.HwpToPdfApp.wait_for_pdf_saved(out_pdf, self.stop_event, timeout=600)
            return True
        return super()._hwp_print_to_pdf(out_pdf)

    def _nup_for(self, ext: str) -> int:
        # 페이지 배치 옵션에 따라 파일별 모아찍기 수(1/2)를 정한다.
        if self.layout_mode == "all2":
            return 2
        if self.layout_mode == "special2":
            return 2 if ext in self.special_exts else 1
        return 1   # all1: 모두 1면 1페이지(hwp/ppt/word는 변환 시 1-up으로 인쇄됨)

    @staticmethod
    def _check_images(pdf_path: str) -> tuple[int, int]:
        # fitz로 PDF 내 이미지를 디코드해 본다. (깨진 이미지 수, 전체 이미지 수) 반환.
        # 디코드 실패 = 손상/누락 의심(특히 한글 변환 시 그림이 깨지는 경우 잡힌다).
        # 파일 자체를 못 열면 (-1, 0).
        bad = 0
        total = 0
        try:
            doc = fitz.open(pdf_path)
        except Exception:
            return (-1, 0)
        try:
            for page in doc:
                for img in page.get_images(full=True):
                    total += 1
                    xref = img[0]
                    try:
                        info = doc.extract_image(xref)
                        if not info or not info.get("image"):
                            bad += 1
                    except Exception:
                        bad += 1
        finally:
            doc.close()
        return (bad, total)

    @Slot()
    def run(self) -> None:
        pythoncom.CoInitialize()
        tmp_root = tempfile.mkdtemp(prefix="ranet_batch_")
        layout_label = {
            "all1": "일괄 1페이지 (1면 1쪽)",
            "all2": "일괄 2페이지 (1면 2쪽)",
            "special2": "특정 유형만 2페이지 (" + ", ".join(sorted(self.special_exts)) + ")",
        }.get(self.layout_mode, self.layout_mode)
        report_folders: list[dict] = []   # 폴더별 구조화 보고 데이터(표 보고서용)
        idx_global = 0           # 임시 PDF 파일명 충돌 방지용 카운터
        try:
            total = len(self.tasks)
            total_files = sum(len(f) for _, _, f, _ in self.tasks) or 1
            done_files = 0          # 진행바를 파일 단위로 움직여 체감 진행이 부드럽게
            success_folders = 0
            self.log.emit(f"총 {total}개 폴더 / {total_files}개 파일 일괄병합 시작 (배치: {layout_label})")
            for fi, (fname, upath, files, unmerged) in enumerate(self.tasks, 1):
                if self.stop_event.is_set():
                    self.log.emit(f"중지됨: {success_folders}개 폴더 처리 후 중단")
                    break
                self.log.emit(f"── [{fname}] {len(files)}개 파일 ({fi}/{total} 폴더) ──")
                writer = PdfWriter()
                file_recs: list[dict] = []   # 병합 파일별 표 행 데이터
                warns: list[str] = []        # 깨짐/실패 의심 문구
                unmerged_names = [os.path.basename(u) for u in unmerged]
                # 구성 메모: 단독 PDF(통합본)·한글 1개만 등 특이 케이스 안내
                note = ""
                if len(files) == 1:
                    e1 = os.path.splitext(files[0])[1].lower()
                    if e1 == ".pdf":
                        note = "원본이 이미 1개 PDF (통합본 → 병합 없이 그대로 포함)"
                    elif e1 in self.HWP_EXTS:
                        note = "한글 파일 1개만 PDF로 변환 (병합 없음)"
                    else:
                        note = f"{e1[1:].upper()} 파일 1개만 변환 (병합 없음)"
                order = 0
                for src in files:
                    if self.stop_event.is_set():
                        break
                    bn = os.path.basename(src)   # 보고서엔 경로 빼고 파일명+확장자만
                    ext = os.path.splitext(src)[1].lower()
                    nup = self._nup_for(ext)
                    order += 1
                    self.folder_progress.emit(done_files, total_files, f"[{fname}] {bn} 변환 중")
                    # 한글 등 COM 변환은 본 스레드를 통째로 막아 진행 신호를 못 보낸다.
                    # 하트비트 스레드가 stdout으로 경과시간을 흘려 멈춤/작업중을 구분시킨다.
                    hb_stop = threading.Event()
                    threading.Thread(
                        target=self._heartbeat,
                        args=(hb_stop, done_files, total_files, fname, bn),
                        daemon=True).start()
                    try:
                        pdf_path = self._to_pdf(src, ext, tmp_root, idx_global)
                        idx_global += 1
                        if pdf_path is None:
                            file_recs.append({"no": order, "name": bn, "pages": "-", "added_pages": 0,
                                              "type": (ext[1:].upper() or "-"), "status": "건너뜀(미지원 형식)"})
                            continue
                        if not self._is_valid_pdf(pdf_path):
                            raise RuntimeError("정상 PDF 아님 (손상 또는 OneDrive 미동기화 스텁)")
                        pages = self._read_pdf_pages(pdf_path, tmp_root, idx_global)
                        if self.force_a4:
                            pages, fixed = self._normalize_oversized_pages(pages)
                            if fixed:
                                self.log.emit(f"[크기조정] {bn}: A4 초과 {fixed}쪽을 A4로 축소")
                        if nup != 1:
                            out_pages = self._arrange_pages(pages, nup)
                            for pg in out_pages:
                                writer.add_page(pg)
                            pages_disp = f"{len(pages)}→{len(out_pages)}(1면{nup})"
                            added = len(out_pages)   # 통합 PDF에 실제 추가된 쪽수(시작 쪽 누적용)
                        else:
                            for pg in pages:
                                writer.add_page(pg)
                            pages_disp = str(len(pages))
                            added = len(pages)
                        file_recs.append({"no": order, "name": bn, "pages": pages_disp, "added_pages": added,
                                          "type": ext[1:].upper(), "status": ""})
                        bad, timg = self._check_images(pdf_path)
                        if bad > 0:
                            warns.append(f"{bn} : 이미지 {timg}개 중 {bad}개 깨짐 의심")
                        elif bad < 0:
                            warns.append(f"{bn} : PDF 열기 실패(검사 불가)")
                        self.log.emit(f"[완료] {bn} ({pages_disp}쪽)")
                    except Exception as exc:
                        file_recs.append({"no": order, "name": bn, "pages": "-", "added_pages": 0,
                                          "type": (ext[1:].upper() or "-"), "status": f"실패: {exc}"})
                        warns.append(f"{bn} : 변환 실패 - {exc}")
                        self.log.emit(f"[실패] {bn}: {exc}")
                    finally:
                        hb_stop.set()
                        done_files += 1
                        self.folder_progress.emit(done_files, total_files, f"[{fname}] 처리 중")

                folder_rec = {
                    "idx": fi, "name": fname, "note": note,
                    "files": file_recs, "unmerged": unmerged_names, "warns": warns,
                }
                if len(writer.pages) == 0:
                    self.log.emit(f"[건너뜀] {fname}: 병합할 페이지 없음")
                    folder_rec.update(ok=False, out="", out_mb="-", pages=0)
                    report_folders.append(folder_rec)
                    self.folder_progress.emit(done_files, total_files, f"[{fname}] 건너뜀 ({fi}/{total} 폴더)")
                    continue

                if self.page_numbers:
                    if self.bake_numbers:
                        self._stamp_page_numbers(writer)         # 본문에 직접 새김
                    else:
                        self._normalize_page_rotations(writer)   # 주석 위치 정확하게

                # 원본 폴더 계층을 저장 위치에도 그대로 재현한다(평면 저장 방지).
                # 예: 2단계면 out_dir/상위폴더/단위폴더.pdf. 경로 각 구간은 안전화.
                rel = os.path.relpath(upath, self.root_dir)
                parts = [_safe_name(p) for p in rel.split(os.sep) if p not in ("", ".", "..")]
                out_path = _unique_path(os.path.join(self.out_dir, *parts) + ".pdf")
                os.makedirs(os.path.dirname(out_path), exist_ok=True)
                with open(out_path, "wb") as f:
                    writer.write(f)
                if self.page_numbers and not self.bake_numbers:
                    self._annotate_page_numbers(out_path)        # 지울 수 있는 주석 쪽수
                success_folders += 1
                try:
                    out_mb = f"{os.path.getsize(out_path) / (1024 * 1024):.1f}"
                except OSError:
                    out_mb = "-"
                self.log.emit(f"[저장] {os.path.basename(out_path)} (총 {len(writer.pages)}쪽)")
                folder_rec.update(ok=True, out=os.path.basename(out_path),
                                  out_mb=out_mb, pages=len(writer.pages))
                report_folders.append(folder_rec)
                self.folder_progress.emit(done_files, total_files, f"[{fname}] 완료 ({fi}/{total} 폴더)")

            # ── 보고서 파일 생성 (표 형식 .hwp) ──
            report_path = ""
            if self.make_report and report_folders:
                meta = {
                    "root_dir": self.root_dir,
                    "out_dir": self.out_dir,
                    "layout": layout_label,
                    "page_numbers": "예" if self.page_numbers else "아니오",
                    "result": f"성공 {success_folders} / {len(self.tasks)} 폴더",
                }
                self.folder_progress.emit(total_files, total_files, "결과 보고서(.hwpx) 작성 중...")
                self.log.emit("결과 보고서(.hwpx) 작성 중...")
                out_hwpx = _unique_path(os.path.join(self.out_dir, "병합결과보고.hwpx"))
                os.makedirs(self.out_dir, exist_ok=True)
                try:
                    # 템플릿 표를 복제·채워 .hwpx 생성(한글 COM 불필요).
                    report_path = build_hwpx_report(REPORT_TEMPLATE, out_hwpx, meta, report_folders)
                    self.log.emit(f"[보고서] {os.path.basename(report_path)} 생성 완료")
                except Exception as exc:
                    # 실패 시 txt로 폴백(보고 내용은 남긴다).
                    report_path = _unique_path(os.path.join(self.out_dir, "병합결과보고.txt"))
                    with open(report_path, "w", encoding="utf-8") as f:
                        f.write(self._report_as_text(meta, report_folders))
                    self.log.emit(f"[보고서] hwpx 생성 실패({exc}) → txt로 저장: {os.path.basename(report_path)}")

            self.batch_done.emit(self.out_dir, success_folders, len(self.tasks), report_path)
        except Exception as exc:
            self.error.emit(str(exc))
        finally:
            self._cleanup_apps()
            shutil.rmtree(tmp_root, ignore_errors=True)
            pythoncom.CoUninitialize()

    def _report_as_text(self, meta: dict, folders: list) -> str:
        # hwpx 생성 실패 시 폴백용 평문 보고서.
        out = ["[ 병합 결과 보고 ]", "",
               f"최상위 폴더 : {meta['root_dir']}",
               f"저장 위치   : {meta['out_dir']}",
               f"페이지 배치 : {meta['layout']}",
               f"페이지 번호 : {meta['page_numbers']}",
               f"처리 결과   : {meta['result']}", ""]
        for fd in folders:
            out.append("─" * 46)
            head = fd["out"] if fd["ok"] else "(PDF 생성 안 됨)"
            out.append(f"{fd['idx']}. {fd['name']}  →  {head}  (총 {fd.get('pages', 0)}쪽)")
            if fd.get("note"):
                out.append(f"  · 구성 : {fd['note']}")
            for r in fd["files"]:
                st = f"  [{r['status']}]" if r["status"] else ""
                out.append(f"   {r['no']}) {r['name']} ({r['type']}, {r['pages']}쪽){st}")
            if fd["unmerged"]:
                out.append("  · 미병합(엑셀 등): " + ", ".join(fd["unmerged"]))
            out.append("  · 깨짐/실패 의심 : " + (" / ".join(fd["warns"]) if fd["warns"] else "없음"))
            out.append("")
        return "\n".join(out)


# 별도 프로세스(stdout)와 GUI 사이 메시지 구분용 접두문자.
_BATCH_MSG_PREFIX = "\x01RBATCH\x01"


def _run_batch_worker_cli(job_path: str) -> int:
    # '--batch-worker <작업파일>'로 재실행됐을 때 도는 함수.
    # 작업 JSON을 읽어 PdfBatchMergeWorker를 돌리고, 진행/로그/완료를 stdout으로 흘린다.
    import json
    from PySide6.QtCore import QCoreApplication

    with open(job_path, "r", encoding="utf-8") as f:
        job = json.load(f)

    app = QCoreApplication(sys.argv)   # QObject 시그널용(같은 스레드 → 직접 연결, 동기)

    def send(kind: str, *args) -> None:
        # windowed(콘솔없는) exe는 sys.stdout이 None이라 fd 1(파이프)에 직접 바이트로 쓴다.
        try:
            data = (_BATCH_MSG_PREFIX + json.dumps([kind, *args], ensure_ascii=False) + "\n").encode("utf-8")
            os.write(1, data)
        except Exception:
            pass

    # tasks: JSON에서 list[list] → list[tuple]로 복원
    tasks = [(t[0], t[1], list(t[2]), list(t[3])) for t in job["tasks"]]
    worker = PdfBatchMergeWorker(
        tasks, job["out_dir"], job["page_numbers"], job["font_size"], job["position"],
        job["layout_mode"], set(job["special_exts"]), job["make_report"], job["root_dir"],
        force_a4=bool(job.get("force_a4", True)),
        bake_numbers=bool(job.get("bake_numbers", False)),
    )
    worker.log.connect(lambda s: send("log", s))
    worker.folder_progress.connect(lambda c, t, s: send("progress", c, t, s))
    # 하트비트 스레드는 Qt 이벤트루프 없이 stdout으로 직접 진행을 흘려야 한다
    # (큐 연결 시그널은 exec() 없이 전달 안 됨). send를 직접 연결.
    worker._progress_sink = lambda c, t, s: send("progress", c, t, s)
    worker.batch_done.connect(lambda d, sf, tf, r: send("done", d, sf, tf, r))
    worker.error.connect(lambda e: send("error", e))

    # 중지 감시: GUI가 stop 파일을 만들면 우아하게 멈춘다(강제종료 대신).
    # 그래야 finally의 _cleanup_apps()가 돌아 한글/오피스 COM이 정리되고 파일 잠금이 안 남는다.
    stop_path = job.get("stop_path", "")

    def watch_stop() -> None:
        while not worker.stop_event.is_set():
            if stop_path and os.path.exists(stop_path):
                worker.request_stop()
                break
            time.sleep(0.3)

    if stop_path:
        threading.Thread(target=watch_stop, daemon=True).start()

    worker.run()   # 동기 실행(시그널은 같은 스레드 직접연결로 즉시 stdout 전송)
    del app
    return 0


class _BatchProcWorker(QObject):
    # GUI 측: 별도 변환 프로세스를 띄우고 stdout을 읽어 시그널로 중계한다.
    # 무거운 COM 변환은 자식 프로세스가 하므로 GUI 프로세스는 멈추지 않는다.
    log = Signal(str)
    progress = Signal(int, int, str)
    done = Signal(str, int, int, str)
    error = Signal(str)

    def __init__(self, job_path: str) -> None:
        super().__init__()
        self.job_path = job_path
        self.stop_path = job_path + ".stop"
        self._proc = None
        self._stopped = False

    def request_stop(self) -> None:
        # 강제종료(terminate) 대신 stop 파일을 만들어 자식이 '우아하게' 멈추게 한다.
        # 그래야 자식이 한글/오피스 COM(_cleanup_apps)을 정리하고 끝나 파일 잠금이 안 남는다.
        # 단, 자식이 긴 호출에 묶여 안 끝나면 일정 시간 뒤 강제종료로 폴백한다.
        self._stopped = True
        try:
            with open(self.stop_path, "w") as f:
                f.write("stop")
        except Exception as e:
            log_err("_BatchProcWorker.request_stop", e, self.stop_path)

        def _force_kill_later() -> None:
            for _ in range(120):   # 최대 약 60초 우아한 종료 대기
                time.sleep(0.5)
                if not (self._proc and self._proc.poll() is None):
                    return
            try:
                if self._proc and self._proc.poll() is None:
                    self._proc.terminate()
            except Exception:
                pass

        threading.Thread(target=_force_kill_later, daemon=True).start()

    def hard_kill(self) -> None:
        # 앱 종료 직전 마지막 수단: 자식 변환 프로세스를 즉시 종료한다.
        # (GUI가 닫히면 stop 파일을 기다릴 새 없이 고아 프로세스가 남을 수 있음)
        try:
            if self._proc and self._proc.poll() is None:
                self._proc.terminate()
        except Exception:
            pass

    @Slot()
    def run(self) -> None:
        import json
        # 자기 자신을 배치워커 모드로 재실행. frozen(exe)이면 exe, 아니면 python+스크립트.
        if getattr(sys, "frozen", False):
            args = [sys.executable, "--batch-worker", self.job_path]
        else:
            args = [sys.executable, os.path.abspath(__file__), "--batch-worker", self.job_path]
        CREATE_NO_WINDOW = 0x08000000   # 콘솔 창 안 뜨게
        # onefile exe가 자기 자신을 '--batch-worker' 자식으로 띄울 때, 부모가 이미 푼 압축해제
        # 폴더를 '재사용'하게 만든다 → 92MB 재압축해제와 그동안의 부트로더 스플래시 표시를 건너뜀.
        #   · 먼저 상속된 _PYI_*/_MEIPASS2 를 싹 지워 깨끗한 상태로 만들고(과거 'Failed to load
        #     Python DLL' 회피용),
        #   · 그다음 같은 exe 재사용에 필요한 값만 명시 지정한다.
        #   (과거 '제거만' 하던 방식은 자동업데이트 .bat = '다른 exe' 교체용이라 여기와 반대다.)
        env = os.environ.copy()
        for k in list(env):
            if k == "_MEIPASS2" or k.startswith("_PYI"):
                env.pop(k, None)
        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            env["_PYI_APPLICATION_HOME_DIR"] = sys._MEIPASS      # 부모 압축해제 폴더 재사용
            env["_PYI_ARCHIVE_FILE"] = sys.executable           # onefile 아카이브(=exe)
            env["_PYI_PARENT_PROCESS_LEVEL"] = "1"               # '자식' 표시 → 재압축해제·스플래시 생략
        try:
            self._proc = subprocess.Popen(
                args, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                creationflags=CREATE_NO_WINDOW, text=True, encoding="utf-8", bufsize=1,
                env=env,
            )
        except Exception as exc:
            self.error.emit(f"변환 프로세스 시작 실패: {exc}")
            return
        got_done = False
        try:
            for line in self._proc.stdout:
                if not line.startswith(_BATCH_MSG_PREFIX):
                    continue
                try:
                    kind, *a = json.loads(line[len(_BATCH_MSG_PREFIX):])
                except Exception:
                    continue
                if kind == "log":
                    self.log.emit(a[0])
                elif kind == "progress":
                    self.progress.emit(int(a[0]), int(a[1]), a[2])
                elif kind == "done":
                    got_done = True
                    self.done.emit(a[0], int(a[1]), int(a[2]), a[3])
                elif kind == "error":
                    self.error.emit(a[0])
        except Exception as exc:
            self.error.emit(f"변환 진행 읽기 오류: {exc}")
        finally:
            try:
                self._proc.wait(timeout=10)
            except Exception:
                pass
            if self._stopped:
                self.error.emit("사용자 중지로 변환을 멈췄습니다.")
            elif not got_done:
                self.error.emit("변환 프로세스가 예기치 않게 종료되었습니다.")


class _PulseTimer(QObject):
    _STOPS  = [0.0, 0.33, 0.66, 1.0]
    _COLORS = [QColor("#FFFFFF"), QColor("#EEECFC"), QColor("#CFC8F6"), QColor("#FFFFFF")]
    _DURATION_MS = 5000

    def __init__(self, button: QPushButton) -> None:
        super().__init__(button)
        self._btn = button
        self._t = 0.0
        self._timer = QTimer(self)
        self._timer.setInterval(16)
        self._timer.timeout.connect(self._tick)

    def start(self) -> None:
        if not self._timer.isActive():
            self._timer.start()

    def stop(self) -> None:
        self._timer.stop()

    def is_running(self) -> bool:
        return self._timer.isActive()

    @staticmethod
    def _lerp(c1: QColor, c2: QColor, t: float) -> QColor:
        return QColor(
            round(c1.red()   + (c2.red()   - c1.red())   * t),
            round(c1.green() + (c2.green() - c1.green()) * t),
            round(c1.blue()  + (c2.blue()  - c1.blue())  * t),
        )

    def _tick(self) -> None:
        self._t = (self._t + 16 / self._DURATION_MS) % 1.0
        t = self._t
        for i in range(len(self._STOPS) - 1):
            if t <= self._STOPS[i + 1]:
                seg = (t - self._STOPS[i]) / (self._STOPS[i + 1] - self._STOPS[i])
                color = self._lerp(self._COLORS[i], self._COLORS[i + 1], seg)
                self._btn.setStyleSheet(
                    f"background:{color.name()};"
                    f"color:{ACCENT};"
                    f"border:2px solid {ACCENT};"
                    f"border-radius:12px;"
                )
                break


class BaseTool(QWidget):
    TOOL_LABEL = "도구"   # 진단 이슈 제목용. 각 탭에서 덮어쓴다.

    def __init__(self) -> None:
        super().__init__()
        self._pulses: dict[QPushButton, _PulseTimer] = {}
        self._diag_sender: DiagReportSender | None = None

    def reset_state(self) -> None:
        pass

    def _page_header(self, layout: QVBoxLayout, title: str, desc: str) -> None:
        # 각 도구 페이지 상단의 제목 + 한 줄 소개 (홈 헤더와 같은 톤).
        t = QLabel(title)
        t.setFont(_font(13, "bold"))
        t.setStyleSheet(f"color:{TXT};background:transparent;")
        d = QLabel(desc)
        d.setFont(_font(8.5, "regular"))
        d.setStyleSheet(f"color:{TXT_S};background:transparent;")
        layout.addWidget(t)
        layout.addWidget(d)
        layout.addSpacing(2)

    # ── 파일/폴더 드래그 앤 드롭 ──
    # 탭이 self.setAcceptDrops(True)를 켜고 _handle_dropped_paths를 구현하면
    # 페이지 어디에 떨어뜨려도 받는다(자식 위젯이 드롭을 안 받으면 부모로 전파됨).
    def dragEnterEvent(self, e) -> None:
        if e.mimeData().hasUrls():
            # 드래그 진입 시 창을 앞으로 — 다른 창에 가려져 있어도 바로 떨어뜨릴 수 있게
            self.window().raise_()
            self.window().activateWindow()
            e.acceptProposedAction()

    def dropEvent(self, e) -> None:
        paths = [u.toLocalFile() for u in e.mimeData().urls() if u.toLocalFile()]
        if paths:
            self._handle_dropped_paths(paths)
            e.acceptProposedAction()

    def _handle_dropped_paths(self, paths: list[str]) -> None:
        pass

    # ── 오류 진단 로그 전송 (동의 시 GitHub 이슈) ──
    def _offer_diag(self, err_msg: str) -> None:
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Icon.Question)
        msg.setWindowTitle("진단 로그 전송")
        msg.setText("문제 해결을 위해 작업 로그를 개발자에게 보낼까요?")
        msg.setInformativeText(
            "보내는 내용: 오류 메시지 + 작업 로그 (파일명이 포함될 수 있습니다)\n"
            "전송 위치: 라넷 서버(jnranet.com)")
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg.setDefaultButton(QMessageBox.StandardButton.Yes)
        _korean_buttons(msg)
        if msg.exec() != QMessageBox.StandardButton.Yes:
            return
        log_text = ""
        if hasattr(self, "log_box"):
            log_text = self.log_box.toPlainText()[-6000:]   # 이슈 본문 과대화 방지
        self._diag_sender = DiagReportSender(
            f"{self.TOOL_LABEL} 오류", err_msg, log_text, {"tool": self.TOOL_LABEL})
        self._diag_sender.done.connect(self._on_diag_done)
        self._diag_sender.start()

    def _on_diag_done(self, ok: bool, info: str) -> None:
        if hasattr(self, "status_lbl"):
            self.status_lbl.setText(
                "진단 로그 전송 완료 — 감사합니다" if ok
                else f"진단 로그 전송 실패 ({info}) — 무시해도 됩니다")

    def _start_pulse(self, button: QPushButton) -> None:
        if button not in self._pulses:
            self._pulses[button] = _PulseTimer(button)
        self._pulses[button].start()

    def _stop_pulse(self, button: QPushButton) -> None:
        p = self._pulses.get(button)
        if p:
            p.stop()

    def _set_action_ready(self, button: QPushButton, ready: bool) -> None:
        button.setEnabled(ready)
        if ready:
            _apply_glow(button, ACCENT, 34)
            self._start_pulse(button)
        else:
            self._stop_pulse(button)
            button.setStyleSheet(BTN_OFF)
            _clear_effect(button)

    def _set_running_buttons(self, start: QPushButton, running: bool, stop: QPushButton | None = None) -> None:
        if running:
            self._stop_pulse(start)
        start.setEnabled(not running)
        start.setStyleSheet(BTN_OFF if running else BTN_ACTION)
        if running:
            _clear_effect(start)
        else:
            _apply_glow(start, ACCENT, 34)
        if stop:
            stop.setEnabled(running)
            stop.setStyleSheet(BTN_STOP if running else BTN_SOFT)
            if running:
                _apply_glow(stop, DANGER, 20)
            else:
                _clear_effect(stop)


class HwpxTab(BaseTool):
    # HWP → HWPX 일괄 교체 변환 탭. 같은 위치에 .hwpx로 저장하고 원본 .hwp를 삭제한다.
    TOOL_LABEL = "HWP → HWPX 변환"

    def __init__(self) -> None:
        super().__init__()
        self._worker: HwpxWorker | None = None
        self._thread: QThread | None = None
        self._build_ui()
        self.setAcceptDrops(True)   # 폴더/한글 파일 드롭으로 대상 폴더 지정

    def _handle_dropped_paths(self, paths: list[str]) -> None:
        if self.is_running():
            return
        folder = ""
        for p in paths:
            if os.path.isdir(p):
                folder = p
                break
        if not folder:
            for p in paths:
                if os.path.splitext(p)[1].lower() == ".hwp":
                    folder = os.path.dirname(p)
                    break
        if not folder:
            return
        self.src_edit.setText(os.path.abspath(folder))
        self._refresh_files()
        self._update_start_ready()

    def reset_state(self) -> None:
        if self.is_running():
            return
        self.src_edit.clear()
        self.subfolder.setChecked(True)
        self.del_chk.setChecked(True)
        self.select_all.setChecked(True)
        self.file_listw.blockSignals(True)
        self.file_listw.clear()
        self.file_listw.blockSignals(False)
        _placeholder_item(self.file_listw, "폴더를 선택하면 .hwp 파일 목록이 표시됩니다.")
        self.file_count.setText("폴더를 선택해주세요.")
        self.progress.setValue(0)
        self.count_lbl.clear()
        self.status_lbl.setText("대기 중")
        self.log_box.clear()
        self._set_action_ready(self.start_btn, False)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet(BTN_SOFT)
        _clear_effect(self.stop_btn)

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("scroll_content")
        layout = QVBoxLayout(root)
        layout.setContentsMargins(20, 16, 20, 10)
        layout.setSpacing(10)
        self._page_header(
            layout, "HWP → HWPX 변환",
            "폴더 안의 한글(.hwp) 문서를 같은 자리에 .hwpx로 바꿔 저장합니다. 내용·서식은 그대로 유지됩니다.")

        card, vl = _card("HWP 파일이 있는 폴더")
        row = QHBoxLayout()
        self.src_edit = QLineEdit()
        self.src_edit.setReadOnly(True)
        self.src_edit.setPlaceholderText("폴더를 선택하거나 이 창에 끌어다 놓으세요")
        btn = QPushButton("폴더 선택")
        btn.setFont(_font(9, "medium"))
        btn.setFixedHeight(34)
        btn.setStyleSheet(BTN_BROWSE)
        btn.clicked.connect(self._select_src)
        row.addWidget(self.src_edit)
        row.addWidget(btn)
        vl.addLayout(row)
        opt = QHBoxLayout()
        self.subfolder = QCheckBox("하위 폴더까지 검색")
        self.subfolder.setChecked(True)
        self.subfolder.toggled.connect(self._refresh_files)
        opt.addWidget(self.subfolder)
        # 원본 처리 옵션: 체크 시 변환 성공한 .hwp를 휴지통으로(복구 가능). 기본 켬.
        self.del_chk = QCheckBox("변환 후 원본 .hwp 삭제  ( 휴지통으로 이동 — 복구 가능 )")
        self.del_chk.setChecked(True)
        opt.addWidget(self.del_chk)
        opt.addStretch()
        vl.addLayout(opt)
        layout.addWidget(card)

        card, vl = _card("변환할 파일 목록")
        top = QHBoxLayout()
        self.select_all = QCheckBox("전체 선택")
        self.select_all.setChecked(True)
        self.select_all.toggled.connect(self._toggle_all)
        self.file_count = QLabel("폴더를 선택해주세요.")
        self.file_count.setFont(_font(8, "regular"))
        self.file_count.setStyleSheet(f"color:{TXT_S};background:transparent;")
        top.addWidget(self.select_all)
        top.addWidget(self.file_count)
        top.addStretch()
        vl.addLayout(top)
        self.file_listw = _make_check_list(190)
        self.file_listw.itemChanged.connect(lambda _it: self._update_start_ready())
        _placeholder_item(self.file_listw, "폴더를 선택하면 .hwp 파일 목록이 표시됩니다.")
        vl.addWidget(self.file_listw)
        layout.addWidget(card)

        # 진행 상황
        card, vl = _card("진행 상황")
        top = QHBoxLayout()
        self.count_lbl = QLabel("")
        self.count_lbl.setFont(_font(9, "regular"))
        self.count_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        top.addStretch()
        top.addWidget(self.count_lbl)
        vl.addLayout(top)
        self.progress = QProgressBar()
        self.progress.setFixedHeight(6)
        self.progress.setTextVisible(False)
        vl.addWidget(self.progress)
        self.status_lbl = QLabel("대기 중")
        self.status_lbl.setFont(_font(9, "regular"))
        self.status_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        vl.addWidget(self.status_lbl)
        layout.addWidget(card)

        row = QHBoxLayout()
        row.setContentsMargins(6, 10, 6, 10)
        row.setSpacing(100)
        self.start_btn = QPushButton("변환 시작")
        self.stop_btn = QPushButton("변환 중지")
        for b in (self.start_btn, self.stop_btn):
            b.setFont(_font(11, "semibold"))
            b.setFixedHeight(32)
            b.setFixedWidth(140)
        self._set_action_ready(self.start_btn, False)
        self.stop_btn.setStyleSheet(BTN_SOFT)
        self.stop_btn.setEnabled(False)
        self.start_btn.clicked.connect(self._start)
        self.stop_btn.clicked.connect(self._stop)
        row.addStretch()
        row.addWidget(self.start_btn)
        row.addWidget(self.stop_btn)
        row.addStretch()
        layout.addLayout(row)

        card, vl = _card("작업 로그")
        self.log_box = LogoLogTextEdit()
        self.log_box.setObjectName("log_box")
        self.log_box.setReadOnly(True)
        self.log_box.setFont(_font(9, "light"))
        self.log_box.setFixedHeight(120)
        vl.addWidget(self.log_box)
        layout.addWidget(card)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(_make_scroll(root))

    def _select_src(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "HWP 파일이 있는 폴더 선택")
        if folder:
            self.src_edit.setText(os.path.abspath(folder))
            self._refresh_files()
            self._update_start_ready()

    def _find_files(self) -> list[str]:
        # .hwp만 수집(.hwpx는 이미 변환된 것이므로 제외). 숨김/시스템 파일 제외.
        src = self.src_edit.text().strip()
        if not src or not os.path.isdir(src):
            return []
        files: list[str] = []
        if self.subfolder.isChecked():
            for r, _, names in os.walk(src):
                for name in names:
                    path = os.path.join(r, name)
                    if name.lower().endswith(".hwp") and not _is_hidden_or_system(path):
                        files.append(path)
        else:
            for name in os.listdir(src):
                path = os.path.join(src, name)
                if (os.path.isfile(path) and name.lower().endswith(".hwp")
                        and not _is_hidden_or_system(path)):
                    files.append(path)
        return sorted(files, key=lambda p: _natural_key(os.path.relpath(p, src)))

    def _refresh_files(self) -> None:
        self.file_listw.blockSignals(True)
        self.file_listw.clear()
        self.file_listw.blockSignals(False)
        files = self._find_files()
        if not files:
            _placeholder_item(self.file_listw, "변환할 .hwp 파일이 없습니다.")
            self.file_count.setText("0개 파일")
            self._update_start_ready()
            return
        src = self.src_edit.text().strip()
        self.file_listw.blockSignals(True)
        for path in files:
            _add_check_item(self.file_listw, os.path.relpath(path, src), path, True)
        self.file_listw.blockSignals(False)
        self.file_count.setText(f"{len(files)}개 파일")
        self.select_all.blockSignals(True)
        self.select_all.setChecked(True)
        self.select_all.blockSignals(False)
        self._update_start_ready()

    def _toggle_all(self, checked: bool) -> None:
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        self.file_listw.blockSignals(True)
        for i in range(self.file_listw.count()):
            it = self.file_listw.item(i)
            if it.data(Qt.ItemDataRole.UserRole) is not None:
                it.setCheckState(state)
        self.file_listw.blockSignals(False)
        self._update_start_ready()

    def _selected_files(self) -> list[str]:
        res: list[str] = []
        for i in range(self.file_listw.count()):
            it = self.file_listw.item(i)
            path = it.data(Qt.ItemDataRole.UserRole)
            if path and it.checkState() == Qt.CheckState.Checked:
                res.append(path)
        return res

    def _update_start_ready(self) -> None:
        if not hasattr(self, "start_btn"):
            return
        self._set_action_ready(self.start_btn, bool(self._selected_files()))

    def _start(self) -> None:
        files = self._selected_files()
        if not files:
            QMessageBox.warning(self, "경고", "변환할 파일을 선택해주세요.")
            return
        delete_original = self.del_chk.isChecked()
        if delete_original:
            # 원본 삭제(휴지통)가 동반될 때만 시작 전 1회 확인
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Question)
            msg.setWindowTitle("변환 확인")
            msg.setText(f"{len(files)}개 .hwp 파일을 같은 위치에 .hwpx로 변환합니다.")
            msg.setInformativeText("변환에 성공한 원본 .hwp는 휴지통으로 이동됩니다. 계속할까요?")
            msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            msg.setDefaultButton(QMessageBox.StandardButton.Yes)
            _korean_buttons(msg)
            if msg.exec() != QMessageBox.StandardButton.Yes:
                return

        self.log_box.clear()
        self.progress.setValue(0)
        self._set_running_buttons(self.start_btn, True, self.stop_btn)
        self._worker = HwpxWorker(files, delete_original)
        self._thread = QThread()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self.log_box.append)
        self._worker.progress.connect(self._on_progress)
        self._worker.done.connect(self._on_done)
        self._worker.error.connect(self._on_error)
        self._worker.done.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.finished.connect(self._on_finished)
        self._thread.start()

    def _stop(self) -> None:
        if self._worker:
            self._worker.request_stop()
        self.stop_btn.setEnabled(False)
        # 한글이 파일 하나를 여는 중이면 그 호출이 끝나야 멈춘다(느린 PC 는 수십 초).
        # 아무 표시가 없으면 «먹통» 으로 보이므로 상태·로그로 알린다.
        self.status_lbl.setText("중지 요청됨 - 진행 중인 파일을 마치는 중")
        self.log_box.append("[중지 요청] 한글이 처리 중인 파일을 마친 뒤 멈춥니다. "
                            "느린 PC 에서는 수십 초 걸릴 수 있습니다.")

    def _on_progress(self, current: int, total: int, status: str) -> None:
        pct = int(current / total * 100) if total else 0
        self.progress.setValue(pct)
        self.count_lbl.setText(f"{current} / {total} ({pct}%)")
        # 중지를 누른 뒤에는 진행 문구가 «변환 중» 으로 되돌아가지 않게 한다.
        if self._worker and self._worker.stop_event.is_set():
            self.status_lbl.setText("중지 요청됨 - 진행 중인 파일을 마치는 중")
        else:
            self.status_lbl.setText(status)

    def _on_done(self, success: int, fail: int) -> None:
        self.progress.setValue(100)
        self.status_lbl.setText(f"완료 - 성공 {success}개 / 실패 {fail}개")
        _recent_add("HWP → HWPX", f"{success}개 파일 변환",
                    os.path.basename(self.src_edit.text().strip()) or "")
        self._refresh_files()   # 삭제 옵션 사용 시 .hwp가 사라졌으므로 목록 갱신
        tail = ("\n\n성공한 파일의 원본 .hwp는 휴지통으로 이동되었습니다."
                if self.del_chk.isChecked() else "\n\n원본 .hwp는 그대로 유지되었습니다.")
        QMessageBox.information(
            self, "완료",
            f"변환이 완료되었습니다.\n성공 {success}개 / 실패 {fail}개{tail}")

    def _on_error(self, msg: str) -> None:
        self.status_lbl.setText("오류 발생")
        self.log_box.append(f"[오류] {msg}")
        QMessageBox.critical(self, "오류", msg)
        self._offer_diag(msg)   # 동의 시 작업 로그를 GitHub 이슈로 전송

    def _on_finished(self) -> None:
        self._set_running_buttons(self.start_btn, False, self.stop_btn)
        self._update_start_ready()
        self._worker = None
        self._thread = None

    def is_running(self) -> bool:
        return bool(self._thread and self._thread.isRunning())

    def request_stop(self) -> None:
        if self._worker:
            self._worker.request_stop()


class HwpTab(BaseTool):
    TOOL_LABEL = "HWP → PDF 변환"

    def __init__(self) -> None:
        super().__init__()
        self._worker: HwpWorker | None = None
        self._thread: QThread | None = None
        self._build_ui()
        self.setAcceptDrops(True)   # 폴더/한글 파일 드롭으로 대상 폴더 지정

    def _handle_dropped_paths(self, paths: list[str]) -> None:
        # 폴더 드롭 → 변환할 폴더로. .hwp/.hwpx 파일 드롭 → 그 파일이 있는 폴더로.
        if self.is_running():
            return
        folder = ""
        for p in paths:
            if os.path.isdir(p):
                folder = p
                break
        if not folder:
            for p in paths:
                if os.path.splitext(p)[1].lower() in (".hwp", ".hwpx"):
                    folder = os.path.dirname(p)
                    break
        if not folder:
            return
        self.src_edit.setText(self._to_win_path(folder))
        self._sync_dst()
        self._refresh_files()
        self._update_start_ready()

    def reset_state(self) -> None:
        if self.is_running():
            return
        self.src_edit.clear()
        self.dst_edit.clear()
        self.subfolder.setChecked(True)
        self.hwpx.setChecked(True)
        self.same_dir.setChecked(False)
        self.engine_hwp.setChecked(True)
        self.select_all.setChecked(True)
        self._clear_file_list()
        _placeholder_item(self.file_listw, "폴더를 선택하면 파일 목록이 표시됩니다.")
        self.file_count.setText("폴더를 선택하면 파일 목록이 표시됩니다.")
        self.progress.setValue(0)
        self.count_lbl.clear()
        self.status_lbl.setText("대기 중")
        self.log_box.clear()
        self._set_action_ready(self.start_btn, False)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet(BTN_SOFT)
        _clear_effect(self.stop_btn)

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("scroll_content")
        layout = QVBoxLayout(root)
        layout.setContentsMargins(20, 16, 20, 10)
        layout.setSpacing(10)
        self._page_header(
            layout, "HWP → PDF 변환",
            "한글 문서를 PDF로 일괄 변환합니다. 한컴 내장 엔진과 Microsoft Print to PDF 중 선택할 수 있습니다.")

        card, vl = _card("HWP/HWPX 파일이 있는 폴더")
        row = QHBoxLayout()
        self.src_edit = QLineEdit()
        self.src_edit.setReadOnly(True)
        self.src_edit.setPlaceholderText("폴더를 선택하거나 이 창에 끌어다 놓으세요")
        btn = QPushButton("폴더 선택")
        btn.setFont(_font(9, "medium"))
        btn.setFixedHeight(34)
        btn.setStyleSheet(BTN_BROWSE)
        btn.clicked.connect(self._select_src)
        row.addWidget(self.src_edit)
        row.addWidget(btn)
        vl.addLayout(row)
        opt = QHBoxLayout()
        self.subfolder = QCheckBox("하위 폴더까지 검색")
        self.subfolder.setChecked(True)
        self.hwpx = QCheckBox(".hwpx 포함")
        self.hwpx.setChecked(True)
        self.subfolder.toggled.connect(self._refresh_files)
        self.hwpx.toggled.connect(self._refresh_files)
        opt.addWidget(self.subfolder)
        opt.addWidget(self.hwpx)
        hwpx_hint = QLabel("( hwpx는 zip+XML 형태라 PDF 변환 없이 AI가 인식할 수 있습니다 )")
        hwpx_hint.setFont(_font(8, "regular"))
        hwpx_hint.setStyleSheet(f"color:{TXT_T};background:transparent;")
        opt.addWidget(hwpx_hint)
        opt.addStretch()
        vl.addLayout(opt)
        layout.addWidget(card)

        card, vl = _card("변환할 파일 목록")
        top = QHBoxLayout()
        self.select_all = QCheckBox("전체 선택")
        self.select_all.setChecked(True)
        self.select_all.toggled.connect(self._toggle_all)
        self.file_count = QLabel("폴더를 선택하면 파일 목록이 표시됩니다.")
        self.file_count.setFont(_font(8, "regular"))
        self.file_count.setStyleSheet(f"color:{TXT_S};background:transparent;")
        top.addWidget(self.select_all)
        top.addWidget(self.file_count)
        top.addStretch()
        vl.addLayout(top)
        # 수백 개 파일도 가볍게 표시되는 체크 리스트(행마다 위젯 생성 안 함)
        self.file_listw = _make_check_list(190)
        self.file_listw.itemChanged.connect(lambda _it: self._update_start_ready())
        _placeholder_item(self.file_listw, "폴더를 선택하면 파일 목록이 표시됩니다.")
        vl.addWidget(self.file_listw)
        layout.addWidget(card)

        card, vl = _card("PDF 저장 폴더")
        row = QHBoxLayout()
        self.dst_edit = QLineEdit()
        self.dst_edit.setReadOnly(True)
        self.dst_btn = QPushButton("폴더 선택")
        self.dst_btn.setFont(_font(9, "medium"))
        self.dst_btn.setFixedHeight(34)
        self.dst_btn.setStyleSheet(BTN_BROWSE)
        self.dst_btn.clicked.connect(self._select_dst)
        row.addWidget(self.dst_edit)
        row.addWidget(self.dst_btn)
        vl.addLayout(row)
        self.same_dir = QCheckBox("변환할 폴더와 동일")
        self.same_dir.toggled.connect(self._sync_dst)
        vl.addWidget(self.same_dir)
        layout.addWidget(card)

        card, vl = _card("PDF 변환 방식")
        self.engine_group = QButtonGroup(self)
        self.engine_hwp = QRadioButton("한컴오피스 내장 PDF 엔진 (빠르지만 깨질 수 있음)")
        self.engine_ms = QRadioButton("Microsoft Print to PDF (느리지만 깨지지 않음)")
        self.engine_hwp.setChecked(True)
        self.engine_group.addButton(self.engine_hwp)
        self.engine_group.addButton(self.engine_ms)
        vl.addWidget(self.engine_hwp)
        vl.addWidget(self.engine_ms)
        layout.addWidget(card)

        self._build_progress(layout, "대기 중")
        row = QHBoxLayout()
        row.setContentsMargins(6, 10, 6, 10)
        row.setSpacing(100)
        self.start_btn = QPushButton("변환 시작")
        self.stop_btn = QPushButton("변환 중지")
        for b in (self.start_btn, self.stop_btn):
            b.setFont(_font(11, "semibold"))
            b.setFixedHeight(32)
            b.setFixedWidth(140)
        self._set_action_ready(self.start_btn, False)
        self.stop_btn.setStyleSheet(BTN_SOFT)
        self.stop_btn.setEnabled(False)
        self.start_btn.clicked.connect(self._start)
        self.stop_btn.clicked.connect(self._stop)
        row.addStretch()
        row.addWidget(self.start_btn)
        row.addWidget(self.stop_btn)
        row.addStretch()
        layout.addLayout(row)
        self._build_log(layout)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(_make_scroll(root))

    def _build_progress(self, layout: QVBoxLayout, status: str) -> None:
        card, vl = _card("진행 상황")
        top = QHBoxLayout()
        self.count_lbl = QLabel("")
        self.count_lbl.setFont(_font(9, "regular"))
        self.count_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        top.addStretch()
        top.addWidget(self.count_lbl)
        vl.addLayout(top)
        self.progress = QProgressBar()
        self.progress.setFixedHeight(6)
        self.progress.setTextVisible(False)
        vl.addWidget(self.progress)
        self.status_lbl = QLabel(status)
        self.status_lbl.setFont(_font(9, "regular"))
        self.status_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        vl.addWidget(self.status_lbl)
        layout.addWidget(card)

    def _build_log(self, layout: QVBoxLayout) -> None:
        card, vl = _card("작업 로그")
        self.log_box = LogoLogTextEdit()
        self.log_box.setObjectName("log_box")
        self.log_box.setReadOnly(True)
        self.log_box.setFont(_font(9, "light"))
        self.log_box.setFixedHeight(120)
        vl.addWidget(self.log_box)
        layout.addWidget(card)

    @staticmethod
    def _to_win_path(path: str) -> str:
        p = os.path.abspath(path).replace("/", "\\")
        return os.path.realpath(p).replace("/", "\\") if os.path.exists(p) else p

    def _select_src(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "HWP 파일이 있는 폴더 선택")
        if folder:
            self.src_edit.setText(self._to_win_path(folder))
            self._sync_dst()
            self._refresh_files()
            self._update_start_ready()

    def _select_dst(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "PDF 저장 폴더 선택")
        if folder:
            self.dst_edit.setText(self._to_win_path(folder))
            self._update_start_ready()

    def _sync_dst(self) -> None:
        same = self.same_dir.isChecked()
        self.dst_btn.setEnabled(not same)
        self.dst_edit.setEnabled(not same)
        if same:
            self.dst_edit.setText(self.src_edit.text())
        self._update_start_ready()

    def _find_files(self) -> list[str]:
        src = self.src_edit.text().strip()
        if not src or not os.path.isdir(src):
            return []
        exts = [".hwp"]
        if self.hwpx.isChecked():
            exts.append(".hwpx")
        files = []
        # 숨김/시스템 파일(OneDrive 스텁 등 유령 파일)은 제외한다.
        if self.subfolder.isChecked():
            for root, _, names in os.walk(src):
                for name in names:
                    path = os.path.join(root, name)
                    if os.path.splitext(name)[1].lower() in exts and not _is_hidden_or_system(path):
                        files.append(path)
        else:
            for name in os.listdir(src):
                path = os.path.join(src, name)
                if (os.path.isfile(path) and os.path.splitext(name)[1].lower() in exts
                        and not _is_hidden_or_system(path)):
                    files.append(path)
        # 탐색기와 같은 자연 정렬(1, 2, 10 순서. 사전식이면 1, 10, 2가 됨)
        return sorted(files, key=lambda p: _natural_key(os.path.relpath(p, src)))

    def _clear_file_list(self) -> None:
        self.file_listw.blockSignals(True)
        self.file_listw.clear()
        self.file_listw.blockSignals(False)

    def _refresh_files(self) -> None:
        self._clear_file_list()
        files = self._find_files()
        if not files:
            _placeholder_item(self.file_listw, "변환할 HWP/HWPX 파일이 없습니다.")
            self.file_count.setText("0개 파일")
            self._update_start_ready()
            return
        src = self.src_edit.text().strip()
        # 수백 개여도 한 번에 추가(아이템 방식이라 가벼움). 갱신 중 시그널 차단.
        self.file_listw.blockSignals(True)
        for path in files:
            _add_check_item(self.file_listw, os.path.relpath(path, src), path, True)
        self.file_listw.blockSignals(False)
        self.file_count.setText(f"{len(files)}개 파일")
        self.select_all.blockSignals(True)
        self.select_all.setChecked(True)
        self.select_all.blockSignals(False)
        self._update_start_ready()

    def _toggle_all(self, checked: bool) -> None:
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        self.file_listw.blockSignals(True)
        for i in range(self.file_listw.count()):
            it = self.file_listw.item(i)
            if it.data(Qt.ItemDataRole.UserRole) is not None:   # 안내문 행 제외
                it.setCheckState(state)
        self.file_listw.blockSignals(False)
        self._update_start_ready()

    def _selected_files(self) -> list[str]:
        res: list[str] = []
        for i in range(self.file_listw.count()):
            it = self.file_listw.item(i)
            path = it.data(Qt.ItemDataRole.UserRole)
            if path and it.checkState() == Qt.CheckState.Checked:
                res.append(path)
        return res

    def _update_start_ready(self) -> None:
        if not hasattr(self, "start_btn"):
            return
        dst = self.src_edit.text().strip() if self.same_dir.isChecked() else self.dst_edit.text().strip()
        self._set_action_ready(self.start_btn, bool(self._selected_files() and dst))

    def _start(self) -> None:
        files = self._selected_files()
        dst = self.src_edit.text().strip() if self.same_dir.isChecked() else self.dst_edit.text().strip()
        if not files:
            QMessageBox.warning(self, "경고", "변환할 파일을 선택해주세요.")
            return
        if not dst:
            QMessageBox.warning(self, "경고", "PDF 저장 폴더를 선택해주세요.")
            return

        self.log_box.clear()
        self.progress.setValue(0)
        self._set_running_buttons(self.start_btn, True, self.stop_btn)
        engine = "mspdf" if self.engine_ms.isChecked() else "hwp"
        self._worker = HwpWorker(files, dst, engine)
        self._thread = QThread()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self.log_box.append)
        self._worker.progress.connect(self._on_progress)
        self._worker.done.connect(self._on_done)
        self._worker.error.connect(self._on_error)
        self._worker.done.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.finished.connect(self._on_finished)
        self._thread.start()

    def _stop(self) -> None:
        if self._worker:
            self._worker.request_stop()
        self.stop_btn.setEnabled(False)
        self.status_lbl.setText("중지 요청 중")

    def _on_progress(self, current: int, total: int, status: str) -> None:
        pct = int(current / total * 100) if total else 0
        self.progress.setValue(pct)
        self.count_lbl.setText(f"{current} / {total} ({pct}%)")
        self.status_lbl.setText(status)

    def _on_done(self, success: int, fail: int, dst: str) -> None:
        self.progress.setValue(100)
        self.status_lbl.setText(f"완료 - 성공 {success}개 / 실패 {fail}개")
        _recent_add("HWP → PDF", f"{success}개 파일 변환", os.path.basename(dst) or dst)
        QMessageBox.information(self, "완료", f"변환이 완료되었습니다.\n성공 {success}개 / 실패 {fail}개\n\n저장 위치:\n{dst}")

    def _on_error(self, msg: str) -> None:
        self.status_lbl.setText("오류 발생")
        QMessageBox.critical(self, "오류", msg)
        self._offer_diag(msg)   # 동의 시 작업 로그를 GitHub 이슈로 전송

    def _on_finished(self) -> None:
        self._set_running_buttons(self.start_btn, False, self.stop_btn)
        self._update_start_ready()
        self._worker = None
        self._thread = None

    def is_running(self) -> bool:
        return bool(self._thread and self._thread.isRunning())

    def request_stop(self) -> None:
        if self._worker:
            self._worker.request_stop()


class PdfBaseTab(BaseTool):
    def _common_file_card(self, layout: QVBoxLayout, title: str, browse_slot) -> None:
        card, vl = _card(title)
        row = QHBoxLayout()
        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        self.file_edit.setPlaceholderText("PDF를 선택하거나 이 창에 끌어다 놓으세요")
        btn = QPushButton("파일 선택")
        btn.setFont(_font(9, "medium"))
        btn.setFixedHeight(34)
        btn.setStyleSheet(BTN_BROWSE)
        btn.clicked.connect(browse_slot)
        row.addWidget(self.file_edit)
        row.addWidget(btn)
        vl.addLayout(row)
        self.info_lbl = QLabel("-")
        self.info_lbl.setFont(_font(8, "regular"))
        self.info_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        vl.addWidget(self.info_lbl)
        layout.addWidget(card)
        self.browse_btn = btn

    def _progress_card(self, layout: QVBoxLayout) -> None:
        card, vl = _card("진행 상황")
        self.status_lbl = QLabel("파일을 선택해주세요.")
        self.status_lbl.setFont(_font(9, "regular"))
        self.status_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        vl.addWidget(self.status_lbl)
        self.progress = QProgressBar()
        self.progress.setFixedHeight(6)
        self.progress.setTextVisible(False)
        vl.addWidget(self.progress)
        self.count_lbl = QLabel("")
        self.count_lbl.setFont(_font(8, "regular"))
        self.count_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        vl.addWidget(self.count_lbl)
        layout.addWidget(card)

    def _log_card(self, layout: QVBoxLayout) -> None:
        card, vl = _card("작업 로그")
        self.log_box = LogoLogTextEdit()
        self.log_box.setObjectName("log_box")
        self.log_box.setReadOnly(True)
        self.log_box.setFont(_font(9, "light"))
        self.log_box.setFixedHeight(170)
        vl.addWidget(self.log_box)
        layout.addWidget(card)

    def _browse_pdf(self) -> str:
        path, _ = QFileDialog.getOpenFileName(self, "PDF 파일 선택", "", "PDF 파일 (*.pdf);;모든 파일 (*.*)")
        return path


class PdfSplitTab(PdfBaseTab):
    TOOL_LABEL = "PDF 분할"

    def __init__(self) -> None:
        super().__init__()
        self._worker: PdfSplitWorker | None = None
        self._thread: QThread | None = None
        self._build_ui()
        self.setAcceptDrops(True)   # PDF 드롭으로 파일 지정

    def _handle_dropped_paths(self, paths: list[str]) -> None:
        if self.is_running():
            return
        for p in paths:
            if os.path.isfile(p) and p.lower().endswith(".pdf"):
                self._load_file(p)
                return

    def reset_state(self) -> None:
        if self.is_running():
            return
        self.file_edit.clear()
        self.info_lbl.setText("-")
        self.page_radio.setChecked(True)
        self.page_edit.setText("50")
        self.size_edit.setText("10")
        self.progress.setValue(0)
        self.status_lbl.setText("파일을 선택해주세요.")
        self.count_lbl.clear()
        self.log_box.clear()
        self._set_action_ready(self.start_btn, False)

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("scroll_content")
        layout = QVBoxLayout(root)
        layout.setContentsMargins(20, 16, 20, 10)
        layout.setSpacing(10)
        self._page_header(
            layout, "PDF 분할",
            "PDF를 페이지 수 또는 파일 용량 기준으로 여러 개로 나눕니다.")
        self._common_file_card(layout, "분할할 PDF 파일", self._browse)

        card, vl = _card("분할 방식")
        self.mode_group = QButtonGroup(self)
        row1 = QHBoxLayout()
        self.page_radio = QRadioButton("페이지 기준")
        self.page_radio.setChecked(True)
        self.page_edit = QLineEdit("50")
        self.page_edit.setFixedWidth(52)   # 3자리면 충분
        row1.addWidget(self.page_radio)
        row1.addWidget(QLabel("매"))
        row1.addWidget(self.page_edit)
        row1.addWidget(QLabel("페이지마다 분할"))
        row1.addStretch()
        vl.addLayout(row1)
        row2 = QHBoxLayout()
        self.size_radio = QRadioButton("용량 기준")
        self.size_edit = QLineEdit("10")
        self.size_edit.setFixedWidth(52)   # 3자리면 충분
        row2.addWidget(self.size_radio)
        row2.addWidget(QLabel("목표"))
        row2.addWidget(self.size_edit)
        row2.addWidget(QLabel("MB 이하로 분할"))
        row2.addStretch()
        vl.addLayout(row2)
        self.mode_group.addButton(self.page_radio)
        self.mode_group.addButton(self.size_radio)
        layout.addWidget(card)

        self._progress_card(layout)
        self.start_btn = QPushButton("분할 시작")
        self.start_btn.setFont(_font(11, "semibold"))
        self.start_btn.setFixedHeight(32)
        self.start_btn.setFixedWidth(140)
        self.start_btn.setEnabled(False)
        self.start_btn.setStyleSheet(BTN_OFF)
        self.start_btn.clicked.connect(self._start)
        action_row = QHBoxLayout()
        action_row.setContentsMargins(6, 10, 6, 10)
        action_row.addStretch()
        action_row.addWidget(self.start_btn)
        action_row.addStretch()
        layout.addLayout(action_row)
        self._log_card(layout)
        layout.addStretch()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(_make_scroll(root))

    def _browse(self) -> None:
        path = self._browse_pdf()
        if path:
            self._load_file(path)

    def _load_file(self, path: str) -> None:
        # 파일 선택/드롭 공통 진입점.
        self.file_edit.setText(path)
        try:
            reader = PdfReader(path)
            size_mb = os.path.getsize(path) / 1024 / 1024
            self.info_lbl.setText(f"총 {len(reader.pages)}페이지 | {size_mb:.2f} MB")
            self._set_action_ready(self.start_btn, True)
        except Exception as exc:
            QMessageBox.critical(self, "오류", str(exc))

    def _start(self) -> None:
        path = self.file_edit.text()
        if not path:
            return
        try:
            mode = "page" if self.page_radio.isChecked() else "size"
            value = int(self.page_edit.text()) if mode == "page" else float(self.size_edit.text())
            if value <= 0:
                raise ValueError
        except Exception:
            QMessageBox.warning(self, "경고", "분할 기준 값을 올바르게 입력해주세요.")
            return
        self.log_box.clear()
        self.progress.setValue(0)
        self._set_action_ready(self.start_btn, False)
        self._worker = PdfSplitWorker(path, mode, value)
        self._thread = QThread()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self.log_box.append)
        self._worker.progress.connect(self._on_progress)
        self._worker.done.connect(self._on_done)
        self._worker.error.connect(self._on_error)
        self._worker.done.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.finished.connect(self._on_finished)
        self._thread.start()

    def _on_progress(self, current: int, total: int, status: str) -> None:
        pct = int(current / total * 100) if total else 0
        self.progress.setValue(pct)
        self.count_lbl.setText(f"{current} / {total} ({pct}%)")
        self.status_lbl.setText(status)

    def _on_done(self, count: int, folder: str) -> None:
        _recent_add("PDF 분할", f"{count}개 파일 생성", os.path.basename(folder) or folder)
        self.progress.setValue(100)
        self.status_lbl.setText(f"분할 완료 - {count}개 파일 생성")
        QMessageBox.information(self, "완료", f"분할이 완료되었습니다.\n총 {count}개 파일 생성\n\n저장 위치:\n{folder}")

    def _on_error(self, msg: str) -> None:
        self.status_lbl.setText("오류 발생")
        self.log_box.append(f"[오류] {msg}")
        QMessageBox.critical(self, "오류", msg)
        self._offer_diag(msg)   # 동의 시 작업 로그를 GitHub 이슈로 전송

    def _on_finished(self) -> None:
        self._set_action_ready(self.start_btn, bool(self.file_edit.text()))
        self._worker = None
        self._thread = None

    def is_running(self) -> bool:
        return bool(self._thread and self._thread.isRunning())


class PdfMergeTab(BaseTool):
    TOOL_LABEL = "PDF 병합"
    # 폴더에서 파일을 골라 순서대로 하나의 PDF로 병합하는 탭.
    DEFAULT_FONT_SIZE = 11
    MIN_FONT_SIZE = 6
    MAX_FONT_SIZE = 48
    SUPPORTED_EXTS = (
        {".pdf"}
        | PdfMergeWorker.IMG_EXTS
        | PdfMergeWorker.WORD_EXTS
        | PdfMergeWorker.PPT_EXTS
        | PdfMergeWorker.HWP_EXTS
    )
    # 병합하지 않고 보고서 [미병합 파일 목록]에만 표시할 형식(엑셀류).
    EXCEL_EXTS = {".xls", ".xlsx", ".xlsm", ".xlsb"}
    # 확장자 필터 유형(라벨 → 확장자 집합). 목록에서 체크된 유형만 보여준다.
    EXT_CATS = {
        "PDF": {".pdf"},
        "이미지": PdfMergeWorker.IMG_EXTS,
        "Word": PdfMergeWorker.WORD_EXTS,
        "PowerPoint": PdfMergeWorker.PPT_EXTS,
        "한글": PdfMergeWorker.HWP_EXTS,
    }

    def __init__(self) -> None:
        super().__init__()
        self._mode = "single"                 # "single"=단일 병합, "batch"=폴더 일괄병합
        # 모드별 독립 상태(폴더 경로·저장 위치·로그 등)를 따로 보관해 서로 안 섞이게 한다.
        self._snap: dict[str, dict] = {"single": {}, "batch": {}}
        self._folder_files: list[str] = []   # 폴더에서 찾은 전체 파일 경로
        self._merge_list: list[str] = []      # 선택된 병합 대상(체크 순서 유지)
        self._nup: dict[str, int] = {}        # 파일별 모아찍기 수(1/2/4), 기본 1
        self._folder_items: dict[str, QListWidgetItem] = {}   # 폴더 파일 → 체크 아이템
        self._batch_tasks: list[tuple[str, list[str]]] = []   # 일괄: (자식폴더명, 파일목록)
        self._batch_items: dict[str, QListWidgetItem] = {}    # 폴더경로 → 체크 아이템
        self._worker: PdfMergeWorker | None = None
        self._thread: QThread | None = None
        self._build_ui()
        self._set_mode("single")
        self.setAcceptDrops(True)   # 폴더/파일 드롭 지원

    def _handle_dropped_paths(self, paths: list[str]) -> None:
        # 폴더 드롭 → 대상 폴더 교체(단일·일괄 공통).
        # 파일 드롭(단일 모드) → 지원 형식만 병합 목록 끝에 추가.
        if self.is_running():
            return
        dirs = [p for p in paths if os.path.isdir(p)]
        if dirs:
            self.src_edit.setText(self._to_win_path(dirs[0]))
            if not self.dst_edit.text().strip():
                self.dst_edit.setText(self._to_win_path(dirs[0]))
            if self._mode == "single":
                self._refresh_folder_files()
            else:
                self._refresh_batch_tasks()
            self._update_start_ready()
            return
        if self._mode != "single":
            return   # 일괄 모드는 폴더 단위로만 동작
        added = False
        for p in paths:
            if (os.path.isfile(p)
                    and os.path.splitext(p)[1].lower() in self.SUPPORTED_EXTS
                    and p not in self._merge_list):
                self._merge_list.append(p)
                added = True
                it = self._folder_items.get(p)   # 현재 폴더 목록에 있으면 체크 동기화
                if it is not None:
                    self.folder_listw.blockSignals(True)
                    it.setCheckState(Qt.CheckState.Checked)
                    self.folder_listw.blockSignals(False)
        if added:
            self._refresh_merge_list()
            self._reset_progress_display()

    def reset_state(self) -> None:
        if self.is_running():
            return
        self.src_edit.clear()
        self.dst_edit.clear()
        self.name_edit.setText("병합결과")
        self.pagenum_chk.setChecked(False)
        self.size_edit.setText(str(self.DEFAULT_FONT_SIZE))
        self.pos_center.setChecked(True)
        self._on_pagenum_toggled(False)
        self.subfolder.setChecked(False)
        for chk in getattr(self, "_ext_chks", {}).values():
            chk.blockSignals(True)
            chk.setChecked(True)   # 확장자 필터 전체 체크로 복원
            chk.blockSignals(False)
        # 일괄병합 위젯 초기화
        self.depth_combo.setCurrentIndex(0)
        self.batch_recursive.setChecked(True)
        self.report_chk.setChecked(True)
        self.arr_all1.setChecked(True)
        self.special_ext_box.setVisible(False)
        for chk in getattr(self, "_special_chks", {}).values():
            chk.blockSignals(True)
            chk.setChecked(False)
            chk.blockSignals(False)
        self._batch_tasks = []
        self._snap = {"single": {}, "batch": {}}   # 모드별 저장 상태도 초기화
        self._folder_files = []
        self._merge_list = []
        self._nup = {}
        self._folder_items = {}
        self._refresh_folder_list()
        self._refresh_merge_list()
        self._render_batch_preview()
        self.progress.setValue(0)
        self.count_lbl.clear()
        self.status_lbl.setText("폴더를 선택하고 파일을 추가해주세요.")
        self.log_box.clear()
        self._set_action_ready(self.start_btn, False)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet(BTN_SOFT)
        _clear_effect(self.stop_btn)

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("scroll_content")
        layout = QVBoxLayout(root)
        layout.setContentsMargins(20, 16, 20, 10)
        layout.setSpacing(10)
        self._page_header(
            layout, "PDF 병합",
            "여러 문서를 하나의 PDF로 합치거나, 폴더별로 폴더명.pdf 하나씩 일괄 병합합니다.")

        # ── 모드 토글 (단일 병합 / 폴더 일괄병합) ──
        mode_row = QHBoxLayout()
        mode_row.setContentsMargins(0, 0, 0, 0)
        mode_row.setSpacing(8)
        self.mode_single_btn = QPushButton("단일 병합")
        self.mode_batch_btn = QPushButton("폴더 일괄병합")
        for b in (self.mode_single_btn, self.mode_batch_btn):
            b.setFont(_font(9, "semibold"))
            b.setFixedHeight(34)
            b.setCheckable(True)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            mode_row.addWidget(b)
        self.mode_single_btn.clicked.connect(lambda: self._set_mode("single"))
        self.mode_batch_btn.clicked.connect(lambda: self._set_mode("batch"))
        layout.addLayout(mode_row)
        # 일괄 모드 안내(자식 폴더 = PDF). 토글에 따라 보였다 숨겼다 한다.
        self.batch_hint = QLabel(
            "최상위 폴더를 선택하면 그 안의 각 자식 폴더가 폴더명.pdf 하나로 병합됩니다."
        )
        self.batch_hint.setFont(_font(8, "regular"))
        self.batch_hint.setWordWrap(True)
        self.batch_hint.setStyleSheet(f"color:{ACCENT};background:transparent;")
        layout.addWidget(self.batch_hint)

        # ── 폴더 선택 ──
        card, vl = _card("병합할 파일이 있는 폴더")
        self.card_src = card
        row = QHBoxLayout()
        self.src_edit = QLineEdit()
        self.src_edit.setReadOnly(True)
        self.src_edit.setPlaceholderText("폴더를 선택하거나 파일·폴더를 이 창에 끌어다 놓으세요")
        btn = QPushButton("폴더 선택")
        btn.setFont(_font(9, "medium"))
        btn.setFixedHeight(34)
        btn.setStyleSheet(BTN_BROWSE)
        self.src_btn = btn
        btn.clicked.connect(self._select_src)
        row.addWidget(self.src_edit)
        row.addWidget(btn)
        vl.addLayout(row)
        self.subfolder = QCheckBox("하위 폴더까지 검색")     # 단일 모드 전용
        self.subfolder.toggled.connect(self._refresh_folder_files)
        vl.addWidget(self.subfolder)
        # 병합 단위 폴더 단계: 최상위 기준 몇 단계 아래 폴더를 PDF 1개 단위로 볼지.
        # 예) 최상위[샘플] / 1단계=도별 / 2단계=시군별 / 3단계=시군 하위
        self.depth_row_w = QWidget()
        depth_row = QHBoxLayout(self.depth_row_w)
        depth_row.setContentsMargins(0, 0, 0, 0)
        depth_row.setSpacing(8)
        depth_lbl = QLabel("병합 단위 폴더 단계 :")
        depth_lbl.setFont(_font(9, "regular"))
        depth_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        self.depth_combo = QComboBox()
        self.depth_combo.setFont(_font(8, "regular"))
        self.depth_combo.setFixedSize(150, 26)
        self.depth_combo.setCursor(Qt.CursorShape.PointingHandCursor)
        for label, val in (("1단계 (바로 아래)", 1), ("2단계", 2), ("3단계", 3)):
            self.depth_combo.addItem(label, val)
        self.depth_combo.setStyleSheet(
            f"QComboBox {{ background:white; color:{TXT}; border:1px solid {BORDER};"
            f" border-radius:6px; padding:1px 8px; }}"
            f"QComboBox:hover {{ border-color:{ACCENT}; }}"
            "QComboBox::drop-down { border:none; width:18px; }"
            f"QComboBox QAbstractItemView {{ background:white; color:{TXT};"
            f" selection-background-color:{ACCENT_SOFT}; selection-color:{ACCENT};"
            f" border:1px solid {BORDER}; outline:none; }}"
        )
        self.depth_combo.currentIndexChanged.connect(lambda _i: self._refresh_batch_tasks())
        depth_row.addWidget(depth_lbl)
        depth_row.addWidget(self.depth_combo)
        depth_row.addStretch()
        vl.addWidget(self.depth_row_w)
        self.batch_recursive = QCheckBox("하위 폴더 전부 포함  (단위 폴더 속 하위 폴더의 파일까지 재귀 병합)")
        self.batch_recursive.setChecked(True)               # 일괄 모드 전용
        self.batch_recursive.toggled.connect(self._refresh_batch_tasks)
        vl.addWidget(self.batch_recursive)
        hint = QLabel("PDF · 이미지 · Word · PowerPoint · 한글 파일 지원")
        hint.setFont(_font(8, "regular"))
        hint.setStyleSheet(f"color:{TXT_S};background:transparent;")
        vl.addWidget(hint)
        layout.addWidget(card)

        # ── 일괄병합 대상 미리보기 (일괄 모드 전용) ──
        card, vl = _card("일괄병합 대상 (자식 폴더 = PDF 파일)")
        self.card_batch_preview = card
        self.batch_count = QLabel("최상위 폴더를 선택해주세요.")
        self.batch_count.setFont(_font(8, "regular"))
        self.batch_count.setStyleSheet(f"color:{TXT_S};background:transparent;")
        vl.addWidget(self.batch_count)
        # 수백 개 폴더도 가볍게 표시되는 체크 리스트
        self.batch_listw = _make_check_list(220)
        self.batch_listw.itemChanged.connect(lambda _it: self._on_batch_check_toggled())
        vl.addWidget(self.batch_listw)
        layout.addWidget(card)

        # ── 페이지 배치 (일괄 모드 전용) ──
        card, vl = _card("페이지 배치")
        self.card_batch_arrange = card
        self.batch_arr_group = QButtonGroup(self)
        self.arr_all1 = QRadioButton("일괄 1페이지  (1면 1쪽 · 2쪽 모아찍기 문서는 풀어서 출력)")
        self.arr_all2 = QRadioButton("일괄 2페이지  (1면 2쪽)")
        self.arr_special = QRadioButton("특정 유형만 2페이지")
        self.arr_all1.setChecked(True)
        for rb in (self.arr_all1, self.arr_all2, self.arr_special):
            rb.setFont(_font(9, "regular"))
            self.batch_arr_group.addButton(rb)
            vl.addWidget(rb)
        self.arr_special.toggled.connect(lambda c: self.special_ext_box.setVisible(c))
        # special2 선택 시 보일 확장자 체크 줄
        self.special_ext_box = QWidget()
        sx = QHBoxLayout(self.special_ext_box)
        sx.setContentsMargins(24, 0, 0, 0)
        sx.setSpacing(12)
        sx_lbl = QLabel("2페이지로 만들 유형 :")
        sx_lbl.setFont(_font(8, "regular"))
        sx_lbl.setStyleSheet(f"color:{TXT_T};background:transparent;")
        sx.addWidget(sx_lbl)
        self._special_chks: dict[str, QCheckBox] = {}
        for cat in self.EXT_CATS:
            chk = QCheckBox(cat)
            chk.setFont(_font(8, "regular"))
            chk.setCursor(Qt.CursorShape.PointingHandCursor)
            chk.setStyleSheet(
                "QCheckBox { color:%s; background:transparent; spacing:5px; }"
                "QCheckBox::indicator { width:14px; height:14px; border:1px solid #CBCFEA;"
                " border-radius:4px; background:white; }"
                "QCheckBox::indicator:hover { border:1px solid %s; }"
                # 채움 대신 체크표시(채우면 체크가 안 보임)
                "QCheckBox::indicator:checked { border:1px solid %s; background:white;"
                " image:url('%s'); }"
                % (TXT, ACCENT, ACCENT, CHECK_ICON)
            )
            self._special_chks[cat] = chk
            sx.addWidget(chk)
        sx.addStretch()
        self.special_ext_box.setVisible(False)
        vl.addWidget(self.special_ext_box)
        layout.addWidget(card)

        # ── 폴더 내 파일 목록 ──
        card, vl = _card("폴더 내 파일 (체크하여 병합 목록에 추가)")
        self.card_folder_list = card
        self.folder_count = QLabel("폴더를 선택해주세요.")
        self.folder_count.setFont(_font(8, "regular"))
        self.folder_count.setStyleSheet(f"color:{TXT_S};background:transparent;")
        vl.addWidget(self.folder_count)
        # 확장자 필터: 체크된 유형만 목록에 표시(기본 전체 체크). 파일 많을 때 골라 보기.
        filt = QHBoxLayout()
        filt.setSpacing(12)
        filt_lbl = QLabel("유형 :")
        filt_lbl.setFont(_font(8, "regular"))
        filt_lbl.setStyleSheet(f"color:{TXT_T};background:transparent;")
        filt.addWidget(filt_lbl)
        self._ext_chks: dict[str, QCheckBox] = {}
        for cat in self.EXT_CATS:
            chk = QCheckBox(cat)
            chk.setChecked(True)
            chk.setFont(_font(8, "regular"))
            chk.setCursor(Qt.CursorShape.PointingHandCursor)
            chk.setStyleSheet(
                "QCheckBox { color:%s; background:transparent; spacing:5px; }"
                "QCheckBox::indicator { width:14px; height:14px; border:1px solid #CBCFEA;"
                " border-radius:4px; background:white; }"
                "QCheckBox::indicator:hover { border:1px solid %s; }"
                # 채움 대신 체크표시(채우면 체크가 안 보임)
                "QCheckBox::indicator:checked { border:1px solid %s; background:white;"
                " image:url('%s'); }"
                % (TXT, ACCENT, ACCENT, CHECK_ICON)
            )
            chk.toggled.connect(lambda _checked: self._refresh_folder_list())
            self._ext_chks[cat] = chk
            filt.addWidget(chk)
        filt.addStretch()
        vl.addLayout(filt)
        # 수백 개 파일도 가볍게 표시되는 체크 리스트. 체크 = 빨간 ＋(병합 목록에 추가)
        self.folder_listw = _make_check_list(200, accent=DANGER, check_icon=RED_PLUS)
        self.folder_listw.itemChanged.connect(self._on_folder_item_changed)
        vl.addWidget(self.folder_listw)
        add_all = QPushButton("모든 파일 추가")
        add_all.setFont(_font(8, "medium"))
        add_all.setFixedHeight(28)
        add_all.setStyleSheet(BTN_BROWSE)
        add_all.clicked.connect(self._add_all)
        self.add_all_btn = add_all
        vl.addWidget(add_all)
        layout.addWidget(card)

        # ── 병합 순서 ──
        card, vl = _card("병합 순서 (위 → 아래 순서대로 합쳐집니다)")
        self.card_merge_order = card
        top = QHBoxLayout()
        self.merge_count = QLabel("0개 선택됨")
        self.merge_count.setFont(_font(8, "regular"))
        self.merge_count.setStyleSheet(f"color:{TXT_S};background:transparent;")
        clear_btn = QPushButton("전체 비우기")
        clear_btn.setFont(_font(8, "medium"))
        clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        clear_btn.setStyleSheet(
            f"QPushButton {{ background:#E9EAF6; color:{TXT_S}; border:none;"
            f" border-radius:9px; padding:5px 14px; }}"
            f"QPushButton:hover {{ background:{ACCENT_SOFT}; color:{ACCENT}; }}"
        )
        clear_btn.clicked.connect(self._on_clear_clicked)
        self.clear_btn = clear_btn
        top.addWidget(self.merge_count)
        top.addStretch()
        top.addWidget(clear_btn)
        vl.addLayout(top)
        hint2 = QLabel("파일명을 드래그해 순서 변경. 오른쪽에서 파일별 페이지 배치(1면 1·2·4페이지)를 고를 수 있습니다.")
        hint2.setFont(_font(8, "regular"))
        hint2.setWordWrap(True)   # 길어도 줄바꿈 → 카드 가로 폭이 창을 넘지 않게
        hint2.setStyleSheet(f"color:{TXT_T};background:transparent;")
        vl.addWidget(hint2)
        # 내부 드래그(InternalMove)로 행 순서 변경 가능한 리스트.
        self.merge_listw = QListWidget()
        self.merge_listw.setFixedHeight(150)
        self.merge_listw.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.merge_listw.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.merge_listw.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.merge_listw.setStyleSheet(
            f"QListWidget {{ background:#FBFBFE; border:1px solid {BORDER};"
            f" border-radius:8px; outline:none; padding:3px; }}"
            f"QListWidget::item {{ border-radius:6px; margin:1px; }}"
            f"QListWidget::item:selected {{ background:{ACCENT_SOFT}; }}"
            "QScrollBar:vertical { background:transparent; width:9px; margin:4px 2px 4px 2px; }"
            "QScrollBar::handle:vertical { background:#CAD4E4; min-height:28px; border-radius:4px; }"
            "QScrollBar::handle:vertical:hover { background:#AAB8CC; }"
            "QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,"
            "QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background:transparent; border:none; height:0px; }"
        )
        # 드래그로 행을 옮기면 모델 순서가 바뀐다 → 다음 이벤트 루프에서 동기화.
        self.merge_listw.model().rowsMoved.connect(
            lambda *a: QTimer.singleShot(0, self._sync_from_listw)
        )
        vl.addWidget(self.merge_listw)
        layout.addWidget(card)

        # ── 저장 위치 ──
        card, vl = _card("PDF 저장 위치")
        row = QHBoxLayout()
        self.dst_edit = QLineEdit()
        self.dst_edit.setReadOnly(True)
        self.dst_btn = QPushButton("폴더 선택")
        self.dst_btn.setFont(_font(9, "medium"))
        self.dst_btn.setFixedHeight(34)
        self.dst_btn.setStyleSheet(BTN_BROWSE)
        self.dst_btn.clicked.connect(self._select_dst)
        row.addWidget(self.dst_edit)
        row.addWidget(self.dst_btn)
        vl.addLayout(row)
        self.name_row_w = QWidget()       # 단일 모드 전용(일괄은 폴더명 자동)
        name_row = QHBoxLayout(self.name_row_w)
        name_row.setContentsMargins(0, 0, 0, 0)
        name_row.addWidget(QLabel("파일 이름"))
        self.name_edit = QLineEdit("병합결과")
        name_row.addWidget(self.name_edit)
        name_row.addWidget(QLabel(".pdf"))
        vl.addWidget(self.name_row_w)
        self.batch_name_hint = QLabel("저장 파일명은 각 자식 폴더 이름으로 자동 생성됩니다.")
        self.batch_name_hint.setFont(_font(8, "regular"))
        self.batch_name_hint.setStyleSheet(f"color:{TXT_S};background:transparent;")
        vl.addWidget(self.batch_name_hint)
        self.report_chk = QCheckBox("병합 결과 보고서(.hwpx) 생성")   # 일괄 모드 전용
        self.report_chk.setFont(_font(9, "regular"))
        self.report_chk.setChecked(True)
        vl.addWidget(self.report_chk)
        # 사진→PDF 원본처럼 A4보다 큰 페이지를 A4로 축소(비율 유지). 기본 켬.
        self.a4_chk = QCheckBox("A4보다 큰 페이지를 A4 크기로 맞춤  ( 사진 원본 PDF 등 )")
        self.a4_chk.setFont(_font(9, "regular"))
        self.a4_chk.setChecked(True)
        vl.addWidget(self.a4_chk)
        # 페이지 번호는 체크 시 세부옵션이 펼쳐지므로 맨 아래(다른 체크 항목 뒤)에 둔다.
        self.pagenum_chk = QCheckBox("페이지 번호 매기기  ( 중앙 하단, - 1 - 형식 )")
        self.pagenum_chk.setFont(_font(9, "regular"))
        self.pagenum_chk.toggled.connect(self._on_pagenum_toggled)
        vl.addWidget(self.pagenum_chk)
        size_row = QHBoxLayout()
        size_row.setContentsMargins(24, 0, 0, 0)
        self.size_lbl = QLabel("기본 폰트 크기 :")
        self.size_lbl.setFont(_font(9, "regular"))
        self.size_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        self.size_edit = QLineEdit(str(self.DEFAULT_FONT_SIZE))
        self.size_edit.setFixedWidth(56)
        self.size_edit.setValidator(QIntValidator(self.MIN_FONT_SIZE, self.MAX_FONT_SIZE, self))
        self.size_edit.editingFinished.connect(self._clamp_font_size)
        self.size_unit = QLabel(f"pt  ( {self.MIN_FONT_SIZE} ~ {self.MAX_FONT_SIZE} )")
        self.size_unit.setFont(_font(9, "regular"))
        self.size_unit.setStyleSheet(f"color:{TXT_T};background:transparent;")
        size_row.addWidget(self.size_lbl)
        size_row.addWidget(self.size_edit)
        size_row.addWidget(self.size_unit)
        size_row.addStretch()

        pos_row = QHBoxLayout()
        pos_row.setContentsMargins(24, 0, 0, 0)
        pos_row.setSpacing(14)
        self.pos_lbl = QLabel("번호 위치 :")
        self.pos_lbl.setFont(_font(9, "regular"))
        self.pos_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        self.pos_group = QButtonGroup(self)
        self.pos_left = QRadioButton("아래 왼쪽")
        self.pos_center = QRadioButton("아래 중앙")
        self.pos_right = QRadioButton("아래 오른쪽")
        self.pos_center.setChecked(True)
        for rb in (self.pos_left, self.pos_center, self.pos_right):
            rb.setFont(_font(9, "regular"))
            self.pos_group.addButton(rb)
        pos_row.addWidget(self.pos_lbl)
        pos_row.addWidget(self.pos_left)
        pos_row.addWidget(self.pos_center)
        pos_row.addWidget(self.pos_right)
        pos_row.addStretch()

        # 쪽수 표시 방식: 기본=지울 수 있는 주석 / 체크=본문에 직접 새김(영구).
        bake_row = QHBoxLayout()
        bake_row.setContentsMargins(24, 2, 0, 0)
        self.bake_chk = QCheckBox("본문에 직접 새기기 (지우거나 다시 매길 수 없음)")
        self.bake_chk.setFont(_font(9, "regular"))
        self.bake_chk.setToolTip(
            "끄면(기본): 나중에 지우거나 다시 매길 수 있는 '주석' 쪽수로 넣습니다.\n"
            "          (예전 라넷툴킷 쪽수도 이 방식이면 자동 인식해 교체 가능)\n"
            "켜면: 본문에 영구히 새깁니다. 모든 뷰어·인쇄 호환은 최고지만 제거가 불가합니다.")
        bake_row.addWidget(self.bake_chk)
        bake_row.addStretch()
        bake_hint = QLabel("기본은 '주석' 쪽수 — 나중에 지우거나 다시 매길 수 있어요.")
        bake_hint.setContentsMargins(24, 0, 0, 0)
        bake_hint.setFont(_font(8, "regular"))
        bake_hint.setStyleSheet(f"color:{TXT_T};background:transparent;")

        # 체크 시에만 나타나는 옵션 묶음(크기 + 위치 + 표시방식).
        self.pagenum_opts = QWidget()
        opts_v = QVBoxLayout(self.pagenum_opts)
        opts_v.setContentsMargins(0, 4, 0, 0)
        opts_v.setSpacing(6)
        opts_v.addLayout(size_row)
        opts_v.addLayout(pos_row)
        opts_v.addLayout(bake_row)
        opts_v.addWidget(bake_hint)
        vl.addWidget(self.pagenum_opts)

        self._on_pagenum_toggled(False)   # 처음엔 옵션 숨김
        layout.addWidget(card)

        # ── 진행 상황 ──
        card, vl = _card("진행 상황")
        top = QHBoxLayout()
        self.count_lbl = QLabel("")
        self.count_lbl.setFont(_font(9, "regular"))
        self.count_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        top.addStretch()
        top.addWidget(self.count_lbl)
        vl.addLayout(top)
        self.progress = QProgressBar()
        self.progress.setFixedHeight(6)
        self.progress.setTextVisible(False)
        vl.addWidget(self.progress)
        self.status_lbl = QLabel("폴더를 선택하고 파일을 추가해주세요.")
        self.status_lbl.setFont(_font(9, "regular"))
        self.status_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        vl.addWidget(self.status_lbl)
        layout.addWidget(card)

        # ── 실행 버튼 ──
        row = QHBoxLayout()
        row.setContentsMargins(6, 10, 6, 10)
        row.setSpacing(100)
        self.start_btn = QPushButton("병합 시작")
        self.stop_btn = QPushButton("중지")
        for b in (self.start_btn, self.stop_btn):
            b.setFont(_font(11, "semibold"))
            b.setFixedHeight(32)
            b.setFixedWidth(140)
        self._set_action_ready(self.start_btn, False)
        self.stop_btn.setStyleSheet(BTN_SOFT)
        self.stop_btn.setEnabled(False)
        self.start_btn.clicked.connect(self._start)
        self.stop_btn.clicked.connect(self._stop)
        row.addStretch()
        row.addWidget(self.start_btn)
        row.addWidget(self.stop_btn)
        row.addStretch()
        layout.addLayout(row)

        # ── 로그 ──
        card, vl = _card("작업 로그")
        self.log_box = LogoLogTextEdit()
        self.log_box.setObjectName("log_box")
        self.log_box.setReadOnly(True)
        self.log_box.setFont(_font(9, "light"))
        self.log_box.setFixedHeight(120)
        vl.addWidget(self.log_box)
        layout.addWidget(card)
        layout.addStretch()

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(_make_scroll(root))
        self._refresh_folder_list()
        self._refresh_merge_list()

    @staticmethod
    def _to_win_path(path: str) -> str:
        p = os.path.abspath(path).replace("/", "\\")
        return os.path.realpath(p).replace("/", "\\") if os.path.exists(p) else p

    @staticmethod
    def _style_mode_btn(btn: QPushButton, active: bool) -> None:
        # 모드 토글 버튼 모양. 활성=인디고 채움, 비활성=연회색.
        if active:
            btn.setStyleSheet(
                f"QPushButton {{ background:{ACCENT}; color:white; border:none; border-radius:9px; }}"
            )
        else:
            btn.setStyleSheet(
                f"QPushButton {{ background:#E9EAF6; color:{TXT_S}; border:none; border-radius:9px; }}"
                f"QPushButton:hover {{ background:{ACCENT_SOFT}; color:{ACCENT}; }}"
            )

    def _set_mode(self, mode: str) -> None:
        # 단일/일괄 모드 전환. 관련 카드를 보였다 숨겼다 하고 목록을 갱신한다.
        # 두 모드는 폴더 경로·저장 위치·로그를 따로 보관해 서로 섞이지 않는다.
        if self.is_running():
            return
        prev = self._mode
        if mode == prev and self._snap.get(mode):
            return   # 같은 모드 재요청(초기화 후) 무시
        # 이전 모드 상태 저장
        if hasattr(self, "src_edit"):
            self._snap[prev] = {
                "src": self.src_edit.text(),
                "dst": self.dst_edit.text(),
                "log": self.log_box.toHtml(),
                "checked": {up for up, it in self._batch_items.items()
                            if it.checkState() == Qt.CheckState.Checked},
            }
        self._mode = mode
        is_single = (mode == "single")
        self.mode_single_btn.setChecked(is_single)
        self.mode_batch_btn.setChecked(not is_single)
        self._style_mode_btn(self.mode_single_btn, is_single)
        self._style_mode_btn(self.mode_batch_btn, not is_single)
        # 카드/위젯 표시 전환
        self.batch_hint.setVisible(not is_single)
        self.subfolder.setVisible(is_single)
        self.depth_row_w.setVisible(not is_single)
        self.batch_recursive.setVisible(not is_single)
        self.card_folder_list.setVisible(is_single)
        self.card_merge_order.setVisible(is_single)
        self.card_batch_preview.setVisible(not is_single)
        self.card_batch_arrange.setVisible(not is_single)
        self.name_row_w.setVisible(is_single)
        self.batch_name_hint.setVisible(not is_single)
        self.report_chk.setVisible(not is_single)
        # 새 모드의 저장된 상태 복원(없으면 빈 상태)
        snap = self._snap.get(mode, {})
        self.src_edit.blockSignals(True)
        self.src_edit.setText(snap.get("src", ""))
        self.src_edit.blockSignals(False)
        self.dst_edit.setText(snap.get("dst", ""))
        self.log_box.setHtml(snap.get("log", ""))
        self.progress.setValue(0)
        self.count_lbl.clear()
        if is_single:
            self._merge_list = []
            self._nup = {}
            self._refresh_folder_files()
            self._refresh_merge_list()
        else:
            self._batch_check_restore = snap.get("checked")   # 이전 체크 선택 복원용
            self._refresh_batch_tasks()
            self._batch_check_restore = None
        self._reset_progress_display()
        self._update_start_ready()

    def _collect_ext(self, folder: str, recursive: bool, exts: set[str]) -> list[str]:
        # 폴더 내에서 주어진 확장자 파일을 모아 자연정렬한다(숨김/시스템 파일 제외).
        # recursive=True면 하위 폴더까지 os.walk로 재귀 수집한다.
        found: list[str] = []
        if recursive:
            for r, _, names in os.walk(folder):
                for n in names:
                    p = os.path.join(r, n)
                    if (os.path.splitext(n)[1].lower() in exts
                            and not _is_hidden_or_system(p)):
                        found.append(p)
        else:
            for n in os.listdir(folder):
                p = os.path.join(folder, n)
                if (os.path.isfile(p)
                        and os.path.splitext(n)[1].lower() in exts
                        and not _is_hidden_or_system(p)):
                    found.append(p)
        # 단위 폴더 기준 상대경로 자연정렬 → '1. ' '2. ' 접두 폴더/파일 순서 유지
        found.sort(key=lambda p: _natural_key(os.path.relpath(p, folder)))
        return found

    def _collect_files(self, folder: str, recursive: bool) -> list[str]:
        # 병합 대상(지원 형식) 파일 수집.
        return self._collect_ext(folder, recursive, self.SUPPORTED_EXTS)

    def _unit_folders(self, root: str, depth: int) -> list[str]:
        # 최상위(root) 기준 정확히 depth 단계 아래의 폴더들을 병합 단위로 반환한다.
        # 각 단계마다 자연정렬해 도/시군 등 번호·이름 순서를 유지한다.
        level = [root]
        for _ in range(depth):
            nxt: list[str] = []
            for d in level:
                try:
                    subs = [os.path.join(d, x) for x in os.listdir(d)
                            if os.path.isdir(os.path.join(d, x))]
                except OSError:
                    subs = []
                subs.sort(key=lambda p: _natural_key(os.path.basename(p)))
                nxt.extend(subs)
            level = nxt
        return level

    def _refresh_batch_tasks(self) -> None:
        # 선택한 폴더 단계의 각 폴더를 병합 단위로 수집한다.
        # 각 단위: (표시명=상대경로, 단위폴더경로, 병합파일목록, 미병합(엑셀)목록)
        src = self.src_edit.text().strip()
        tasks: list[tuple[str, str, list[str], list[str]]] = []
        if src and os.path.isdir(src):
            depth = self.depth_combo.currentData() or 1
            recursive = self.batch_recursive.isChecked()
            for upath in self._unit_folders(src, depth):
                files = self._collect_files(upath, recursive)
                unmerged = self._collect_ext(upath, recursive, self.EXCEL_EXTS)
                if files or unmerged:
                    name = os.path.relpath(upath, src)
                    tasks.append((name, upath, files, unmerged))
        self._batch_tasks = tasks
        self._render_batch_preview()
        self._reset_progress_display()
        self._update_start_ready()

    def _render_batch_preview(self) -> None:
        # 일괄병합 대상(자식폴더 → pdf) 목록을 체크 아이템으로 표시. 체크한 폴더만 병합.
        lw = self.batch_listw
        lw.blockSignals(True)
        lw.clear()
        self._batch_items = {}
        src = self.src_edit.text().strip()
        if not self._batch_tasks:
            msg = "최상위 폴더를 선택해주세요." if not src else "병합할 파일이 있는 자식 폴더가 없습니다."
            _placeholder_item(lw, msg)
            lw.blockSignals(False)
            self.batch_count.setText("0개 폴더")
            return
        restore = getattr(self, "_batch_check_restore", None)
        for i, (name, upath, files, unmerged) in enumerate(self._batch_tasks, 1):
            leaf = os.path.basename(upath)
            extra = f"  +엑셀 {len(unmerged)}개(미병합)" if unmerged else ""
            text = f"{i}.  {name}  →  {_safe_name(leaf)}.pdf   ({len(files)}개 파일){extra}"
            # 기본값 전부 선택. 모드 복원 시엔 이전에 체크했던 폴더만 체크.
            checked = upath in restore if restore is not None else True
            self._batch_items[upath] = _add_check_item(lw, text, upath, checked)
        lw.blockSignals(False)
        self._update_batch_count()

    def _checked_batch_tasks(self) -> list:
        # 체크된 폴더의 작업만 추린다(아이템 없으면 전부 포함).
        res = []
        for task in self._batch_tasks:
            it = self._batch_items.get(task[1])
            if it is None or it.checkState() == Qt.CheckState.Checked:
                res.append(task)
        return res

    def _update_batch_count(self) -> None:
        # 선택/전체 폴더 수와 선택분 파일 수를 표시.
        sel = self._checked_batch_tasks()
        total = len(self._batch_tasks)
        files = sum(len(t[2]) for t in sel)
        if len(sel) == total:
            self.batch_count.setText(f"{total}개 폴더 · 총 {files}개 파일")
        else:
            self.batch_count.setText(f"선택 {len(sel)} / 전체 {total}개 폴더 · 선택분 {files}개 파일")

    def _on_batch_check_toggled(self) -> None:
        self._update_batch_count()
        self._reset_progress_display()
        self._update_start_ready()

    def _batch_layout_opts(self) -> tuple[str, set[str]]:
        # 페이지 배치 라디오 → (layout_mode, special_exts) 반환.
        if self.arr_all2.isChecked():
            return "all2", set()
        if self.arr_special.isChecked():
            exts: set[str] = set()
            for cat, chk in self._special_chks.items():
                if chk.isChecked():
                    exts |= self.EXT_CATS[cat]
            return "special2", exts
        return "all1", set()

    def _select_src(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "파일이 있는 폴더 선택")
        if folder:
            self.src_edit.setText(self._to_win_path(folder))
            if not self.dst_edit.text().strip():
                self.dst_edit.setText(self._to_win_path(folder))
            if self._mode == "single":
                self._refresh_folder_files()
            else:
                self._refresh_batch_tasks()
            self._update_start_ready()

    def _select_dst(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "PDF 저장 폴더 선택")
        if folder:
            self.dst_edit.setText(self._to_win_path(folder))
            self._update_start_ready()

    def _refresh_folder_files(self) -> None:
        # 폴더(및 옵션 시 하위 폴더)에서 지원 형식 파일을 모아 정렬한다.
        # 숨김/시스템 속성 파일(OneDrive 스텁·OS 부산물 등 유령 파일)은 제외한다.
        src = self.src_edit.text().strip()
        files: list[str] = []
        if src and os.path.isdir(src):
            if self.subfolder.isChecked():
                for r, _, names in os.walk(src):
                    for name in names:
                        path = os.path.join(r, name)
                        if (os.path.splitext(name)[1].lower() in self.SUPPORTED_EXTS
                                and not _is_hidden_or_system(path)):
                            files.append(path)
            else:
                for name in os.listdir(src):
                    path = os.path.join(src, name)
                    if (os.path.isfile(path)
                            and os.path.splitext(name)[1].lower() in self.SUPPORTED_EXTS
                            and not _is_hidden_or_system(path)):
                        files.append(path)
        # 탐색기와 같은 자연 정렬(1, 2, 10 순서. 사전식이면 1, 10, 2가 됨)
        self._folder_files = sorted(files, key=lambda p: _natural_key(os.path.relpath(p, src)))
        self._refresh_folder_list()

    def _shown_files(self) -> list[str]:
        # 확장자 필터에서 체크된 유형의 파일만 추린다(미빌드 시 전체).
        if not hasattr(self, "_ext_chks"):
            return list(self._folder_files)
        enabled: set[str] = set()
        for cat, chk in self._ext_chks.items():
            if chk.isChecked():
                enabled |= self.EXT_CATS[cat]
        return [p for p in self._folder_files if os.path.splitext(p)[1].lower() in enabled]

    def _on_folder_item_changed(self, item: QListWidgetItem) -> None:
        # 체크 아이템 토글 → 병합 목록에 추가/제거.
        path = item.data(Qt.ItemDataRole.UserRole)
        if path:
            self._on_pick(path, item.checkState() == Qt.CheckState.Checked)

    def _refresh_folder_list(self) -> None:
        # 폴더 파일을 체크 아이템 목록으로 표시. 체크하면 병합 목록에 추가(중복 불가),
        # 해제하면 제거. 체크 시 붉은 ＋ 표시. (확장자 필터 적용)
        lw = self.folder_listw
        lw.blockSignals(True)
        lw.clear()
        self._folder_items = {}
        src = self.src_edit.text().strip()
        shown = self._shown_files()
        if not shown:
            if not src:
                msg = "폴더를 선택해주세요."
            elif self._folder_files:
                msg = "선택한 유형의 파일이 없습니다. (위 유형 체크 확인)"
            else:
                msg = "표시할 파일이 없습니다."
            _placeholder_item(lw, msg)
            lw.blockSignals(False)
            total = len(self._folder_files)
            self.folder_count.setText(f"표시 0 / 전체 {total}개" if total else "0개 파일")
            return
        for path in shown:
            rel = os.path.relpath(path, src) if src else os.path.basename(path)
            self._folder_items[path] = _add_check_item(
                lw, rel, path, path in self._merge_list)
        lw.blockSignals(False)
        total = len(self._folder_files)
        if len(shown) == total:
            self.folder_count.setText(f"{total}개 파일")
        else:
            self.folder_count.setText(f"표시 {len(shown)} / 전체 {total}개")

    def _refresh_merge_list(self) -> None:
        # 선택된 파일을 순서대로 리스트에 표시. 행은 드래그로 옮기거나 ▲▼로 이동.
        self.merge_listw.blockSignals(True)
        self.merge_listw.clear()
        if not self._merge_list:
            ph = QListWidgetItem("선택된 파일이 없습니다.")
            ph.setFlags(Qt.ItemFlag.NoItemFlags)   # 드래그/선택 불가
            ph.setForeground(QColor(TXT_S))
            self.merge_listw.addItem(ph)
            self.merge_listw.blockSignals(False)
            self.merge_count.setText("0개 선택됨")
            self._update_start_ready()
            return
        last = len(self._merge_list) - 1
        for i, path in enumerate(self._merge_list):
            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, path)
            item.setSizeHint(QSize(0, 30))
            # 행은 드래그 가능, 자식 위젯에 드롭 받지 않게 함
            item.setFlags(
                Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsDragEnabled
            )
            self.merge_listw.addItem(item)

            row = QHBoxLayout()
            row.setContentsMargins(6, 0, 4, 0)
            row.setSpacing(4)
            num = QLabel(f"{i + 1}.")
            num.setFont(_font(8, "semibold"))
            num.setFixedWidth(22)
            num.setStyleSheet(f"color:{ACCENT};background:transparent;")
            lbl = QLabel(os.path.basename(path))
            lbl.setFont(_font(8, "regular"))
            lbl.setStyleSheet(f"color:{TXT};background:transparent;")
            nup_combo = self._nup_combo(path)
            up = self._arrow_btn("▲", lambda _, p=path: self._move_path(p, -1), i > 0)
            down = self._arrow_btn("▼", lambda _, p=path: self._move_path(p, 1), i < last)
            row.addWidget(num)
            row.addWidget(lbl, stretch=1)
            row.addWidget(nup_combo)
            row.addWidget(up)
            row.addWidget(down)
            holder = QWidget()
            holder.setStyleSheet("background:transparent;")
            holder.setLayout(row)
            self.merge_listw.setItemWidget(item, holder)
        self.merge_listw.blockSignals(False)
        self.merge_count.setText(f"{len(self._merge_list)}개 선택됨")
        self._update_start_ready()

    def _sync_from_listw(self) -> None:
        # 드래그 후 리스트 위젯의 현재 순서를 _merge_list에 반영하고 번호를 다시 매긴다.
        new: list[str] = []
        for i in range(self.merge_listw.count()):
            p = self.merge_listw.item(i).data(Qt.ItemDataRole.UserRole)
            if p:
                new.append(p)
        if new and new != self._merge_list:
            self._merge_list = new
            self._refresh_merge_list()

    # 파일별 목표 페이지 배치. 파일 원래 설정과 무관하게 이 값대로 출력한다.
    # (HWP는 변환 시 Hancom PDF로 1면1쪽 강제 → 여기서 목표대로 모아찍기만 적용)
    NUP_OPTIONS = (
        ("1면 1페이지", 1),
        ("1면 2페이지", 2),
        ("1면 4페이지", 4),
    )

    def _nup_combo(self, path: str) -> QComboBox:
        # 파일별 목표 페이지 배치 선택. 기본 '1면 1페이지'.
        combo = QComboBox()
        combo.setFont(_font(8, "regular"))
        combo.setFixedSize(96, 24)
        combo.setCursor(Qt.CursorShape.PointingHandCursor)
        idx_of = {}
        for i, (label, val) in enumerate(self.NUP_OPTIONS):
            combo.addItem(label, val)
            idx_of[val] = i
        cur = self._nup.get(path, 1)
        combo.setCurrentIndex(idx_of.get(cur, 0))
        combo.setStyleSheet(
            f"QComboBox {{ background:white; color:{TXT}; border:1px solid {BORDER};"
            f" border-radius:6px; padding:1px 6px; }}"
            f"QComboBox:hover {{ border-color:{ACCENT}; }}"
            "QComboBox::drop-down { border:none; width:16px; }"
            f"QComboBox QAbstractItemView {{ background:white; color:{TXT};"
            f" selection-background-color:{ACCENT_SOFT}; selection-color:{ACCENT};"
            f" border:1px solid {BORDER}; outline:none; }}"
        )
        combo.currentIndexChanged.connect(
            lambda _idx, c=combo, p=path: self._on_nup_changed(p, c.currentData())
        )
        return combo

    def _on_nup_changed(self, path: str, value: int) -> None:
        # 모아찍기 선택 변경 → 저장하고 이전 작업 진행 표시 초기화.
        self._nup[path] = int(value)
        self._reset_progress_display()

    @staticmethod
    def _arrow_btn(text: str, slot, enabled: bool) -> QPushButton:
        # 깔끔한 플랫 위/아래 버튼(채움 없음). 비활성 시 흐리게.
        b = QPushButton(text)
        b.setFont(_font(8, "bold"))
        b.setFixedSize(24, 22)
        b.setEnabled(enabled)
        if enabled:
            b.setStyleSheet(
                f"QPushButton {{ background:transparent; color:{ACCENT};"
                f" border:1px solid {BORDER}; border-radius:6px; }}"
                f"QPushButton:hover {{ background:{ACCENT_SOFT}; border-color:{ACCENT}; }}"
            )
            b.setCursor(Qt.CursorShape.PointingHandCursor)
        else:
            b.setStyleSheet(
                f"QPushButton {{ background:transparent; color:#C7CBDD;"
                f" border:1px solid #EEF0F6; border-radius:6px; }}"
            )
        b.clicked.connect(slot)
        return b

    def _on_pick(self, path: str, checked: bool) -> None:
        # 체크 → 순서 끝에 추가(중복 방지), 해제 → 제거.
        if checked:
            if path not in self._merge_list:
                self._merge_list.append(path)
        else:
            if path in self._merge_list:
                self._merge_list.remove(path)
        self._refresh_merge_list()
        self._reset_progress_display()   # 새 선택 → 이전 작업 진행 표시 초기화

    def _add_all(self) -> None:
        # 현재 목록에 보이는(확장자 필터 통과) 파일만 순서대로 미선택분을 추가.
        self.folder_listw.blockSignals(True)
        for path in self._shown_files():
            if path not in self._merge_list:
                self._merge_list.append(path)
            it = self._folder_items.get(path)
            if it is not None:
                it.setCheckState(Qt.CheckState.Checked)
        self.folder_listw.blockSignals(False)
        self._refresh_merge_list()
        self._reset_progress_display()

    def _move_path(self, path: str, delta: int) -> None:
        if path not in self._merge_list:
            return
        idx = self._merge_list.index(path)
        j = idx + delta
        if 0 <= j < len(self._merge_list):
            self._merge_list[idx], self._merge_list[j] = self._merge_list[j], self._merge_list[idx]
            self._refresh_merge_list()

    def _on_clear_clicked(self) -> None:
        # 사용자가 '전체 비우기'를 누른 경우: 선택 해제 + 진행 표시 초기화.
        self._clear_merge()
        self._reset_progress_display()

    def _clear_merge(self) -> None:
        # 선택 전체 해제.
        self._merge_list = []
        self._nup = {}
        self.folder_listw.blockSignals(True)
        for it in self._folder_items.values():
            it.setCheckState(Qt.CheckState.Unchecked)
        self.folder_listw.blockSignals(False)
        self._refresh_merge_list()

    def _reset_progress_display(self) -> None:
        # 새 작업 준비: 직전 작업의 진행 바·카운트·상태문구를 초기화한다.
        # (작업 실행 중에는 건드리지 않는다.)
        if self.is_running():
            return
        self.progress.setValue(0)
        self.count_lbl.clear()
        ready = self._checked_batch_tasks() if self._mode == "batch" else self._merge_list
        if ready:
            self.status_lbl.setText("병합 시작을 눌러주세요.")
        elif self._mode == "batch":
            self.status_lbl.setText("최상위 폴더를 선택해주세요.")
        else:
            self.status_lbl.setText("폴더를 선택하고 파일을 추가해주세요.")

    def _set_controls_enabled(self, enabled: bool) -> None:
        # 병합 실행 중에는 입력·선택 관련 위젯을 모두 잠근다(중지 버튼만 사용 가능).
        widgets = [
            self.src_btn, self.subfolder, self.add_all_btn, self.clear_btn,
            self.folder_listw, self.merge_listw, self.dst_btn, self.name_edit,
            self.pagenum_chk, self.size_edit, self.bake_chk,
            self.pos_left, self.pos_center, self.pos_right,
            # 일괄병합 관련 위젯
            self.mode_single_btn, self.mode_batch_btn, self.batch_recursive,
            self.depth_combo, self.batch_listw, self.report_chk,
            self.arr_all1, self.arr_all2, self.arr_special,
        ]
        widgets += list(getattr(self, "_ext_chks", {}).values())       # 확장자 필터 체크박스
        widgets += list(getattr(self, "_special_chks", {}).values())   # special2 유형 체크박스
        for w in widgets:
            w.setEnabled(enabled)

    def _on_pagenum_toggled(self, checked: bool) -> None:
        # 페이지 번호 옵션을 켤 때만 크기/위치 입력을 보여준다.
        self.pagenum_opts.setVisible(checked)

    def _position(self) -> str:
        if self.pos_left.isChecked():
            return "left"
        if self.pos_right.isChecked():
            return "right"
        return "center"

    def _clamp_font_size(self) -> None:
        # 입력값을 허용 범위로 보정.
        txt = self.size_edit.text().strip()
        try:
            v = int(txt)
        except ValueError:
            v = self.DEFAULT_FONT_SIZE
        v = max(self.MIN_FONT_SIZE, min(self.MAX_FONT_SIZE, v))
        self.size_edit.setText(str(v))

    def _font_size(self) -> int:
        self._clamp_font_size()
        return int(self.size_edit.text())

    def _update_start_ready(self) -> None:
        if not hasattr(self, "start_btn") or self.is_running():
            return
        if self._mode == "batch":
            ready = bool(self._checked_batch_tasks()) and bool(self.dst_edit.text().strip())
        else:
            ready = bool(self._merge_list) and bool(self.dst_edit.text().strip())
        self._set_action_ready(self.start_btn, ready)

    def _start(self) -> None:
        if self._mode == "batch":
            self._start_batch()
            return
        if not self._merge_list:
            QMessageBox.warning(self, "경고", "병합할 파일을 추가해주세요.")
            return
        dst = self.dst_edit.text().strip()
        if not dst:
            QMessageBox.warning(self, "경고", "PDF 저장 폴더를 선택해주세요.")
            return
        name = self.name_edit.text().strip() or "병합결과"
        if not name.lower().endswith(".pdf"):
            name += ".pdf"
        out_path = _unique_path(os.path.join(dst, name))

        self.log_box.clear()
        self.progress.setValue(0)
        self._set_running_buttons(self.start_btn, True, self.stop_btn)
        self._set_controls_enabled(False)   # 실행 중 다른 버튼 잠금
        nups = [self._nup.get(p, 1) for p in self._merge_list]
        self._worker = PdfMergeWorker(
            list(self._merge_list), out_path,
            self.pagenum_chk.isChecked(), self._font_size(), self._position(),
            nups=nups, force_a4=self.a4_chk.isChecked(),
            bake_numbers=self.bake_chk.isChecked(),
        )
        self._thread = QThread()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self.log_box.append)
        self._worker.progress.connect(self._on_progress)
        self._worker.done.connect(self._on_done)
        self._worker.error.connect(self._on_error)
        self._worker.done.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.finished.connect(self._on_finished)
        self._thread.start()

    def _start_batch(self) -> None:
        # [폴더명으로 일괄병합] 실행. 체크된 자식 폴더만 폴더명.pdf로 병합한다.
        tasks = self._checked_batch_tasks()
        if not tasks:
            QMessageBox.warning(self, "경고", "병합할 폴더를 1개 이상 선택해주세요.")
            return
        dst = self.dst_edit.text().strip()
        if not dst:
            QMessageBox.warning(self, "경고", "PDF 저장 폴더를 선택해주세요.")
            return
        layout_mode, special = self._batch_layout_opts()
        if layout_mode == "special2" and not special:
            QMessageBox.warning(self, "경고", "'특정 유형만 2페이지'에서 유형을 1개 이상 선택해주세요.")
            return

        # 작업 명세를 임시 JSON으로 저장 → 별도 변환 프로세스에 넘긴다(GUI 멈춤 방지).
        import json
        fd, job_path = tempfile.mkstemp(prefix="ranet_batchjob_", suffix=".json")
        job = {
            "tasks": [[n, u, list(fl), list(um)] for (n, u, fl, um) in tasks],
            "out_dir": dst,
            "page_numbers": self.pagenum_chk.isChecked(),
            "bake_numbers": self.bake_chk.isChecked(),
            "font_size": self._font_size(),
            "position": self._position(),
            "layout_mode": layout_mode,
            "special_exts": sorted(special),
            "make_report": self.report_chk.isChecked(),
            "root_dir": self.src_edit.text().strip(),
            "force_a4": self.a4_chk.isChecked(),
            "stop_path": job_path + ".stop",   # 중지 신호 파일(우아한 종료용)
        }
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(job, f, ensure_ascii=False)
        self._job_path = job_path

        self.log_box.clear()
        self.progress.setValue(0)
        self._set_running_buttons(self.start_btn, True, self.stop_btn)
        self._set_controls_enabled(False)
        self._worker = _BatchProcWorker(job_path)
        self._thread = QThread()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self.log_box.append)
        self._worker.progress.connect(self._on_progress)
        self._worker.done.connect(self._on_batch_done)
        self._worker.error.connect(self._on_error)
        self._worker.done.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.finished.connect(self._on_finished)
        self._thread.start()

    def _on_batch_done(self, out_dir: str, success: int, total: int, report_path: str) -> None:
        self.progress.setValue(100)
        self.status_lbl.setText(f"일괄병합 완료 - {success}/{total}개 폴더")
        _recent_add("폴더 일괄병합", f"{success}/{total}개 폴더 병합", os.path.basename(out_dir) or out_dir)
        msg = f"일괄병합이 완료되었습니다.\n{success}/{total}개 폴더 병합\n\n저장 위치:\n{out_dir}"
        if report_path:
            msg += f"\n\n보고서:\n{report_path}"
        QMessageBox.information(self, "완료", msg)

    def _stop(self) -> None:
        if self._worker:
            self._worker.request_stop()
        self.stop_btn.setEnabled(False)
        self.status_lbl.setText("중지 요청 중")

    def _on_progress(self, current: int, total: int, status: str) -> None:
        pct = int(current / total * 100) if total else 0
        self.progress.setValue(pct)
        self.count_lbl.setText(f"{current} / {total} ({pct}%)")
        self.status_lbl.setText(status)

    def _on_done(self, out_path: str, success: int, total: int) -> None:
        self.progress.setValue(100)
        self.status_lbl.setText(f"병합 완료 - {success}/{total}개 파일")
        _recent_add("PDF 병합", os.path.basename(out_path), f"{success}/{total}개 파일")
        # 병합 완료 → 선택 자동 초기화(체크 해제 + 순서 목록 비움). 폴더는 유지.
        self._clear_merge()
        QMessageBox.information(
            self, "완료",
            f"병합이 완료되었습니다.\n{success}/{total}개 파일 병합\n\n저장 위치:\n{out_path}",
        )

    def _on_error(self, msg: str) -> None:
        self.status_lbl.setText("오류 발생")
        self.log_box.append(f"[오류] {msg}")
        QMessageBox.critical(self, "오류", msg)
        self._offer_diag(msg)   # 동의 시 작업 로그를 GitHub 이슈로 전송

    def _on_finished(self) -> None:
        self._set_running_buttons(self.start_btn, False, self.stop_btn)
        self._set_controls_enabled(True)   # 실행 종료 → 다시 사용 가능
        self._update_start_ready()
        self._worker = None
        self._thread = None
        # 일괄병합 작업 임시파일 정리(작업명세 + 중지신호 파일)
        jp = getattr(self, "_job_path", "")
        if jp:
            for p in (jp, jp + ".stop"):
                try:
                    os.remove(p)
                except OSError:
                    pass
            self._job_path = ""

    def is_running(self) -> bool:
        return bool(self._thread and self._thread.isRunning())

    def request_stop(self) -> None:
        if self._worker:
            self._worker.request_stop()


class PdfCompressTab(PdfBaseTab):
    TOOL_LABEL = "PDF 압축"
    PRESETS = {"낮음": (85, 150), "보통": (60, 100), "높음": (35, 72)}
    PRESET_TEXT = {
        "낮음": "낮음\nJPEG 85 / 150 DPI",
        "보통": "보통\nJPEG 60 / 100 DPI",
        "높음": "높음\nJPEG 35 / 72 DPI",
    }

    def __init__(self) -> None:
        super().__init__()
        self._worker: PdfCompressWorker | None = None
        self._thread: QThread | None = None
        self._quality = "보통"
        self._quality_btns: dict[str, QPushButton] = {}
        self._build_ui()
        self.setAcceptDrops(True)   # PDF 드롭으로 파일 지정

    def reset_state(self) -> None:
        if self.is_running():
            return
        self.file_edit.clear()
        self.info_lbl.setText("-")
        self._select_quality("보통")
        self.progress.setValue(0)
        self.status_lbl.setText("파일을 선택해주세요.")
        self.count_lbl.clear()
        self.log_box.clear()
        self._set_action_ready(self.start_btn, False)

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("scroll_content")
        layout = QVBoxLayout(root)
        layout.setContentsMargins(20, 16, 20, 10)
        layout.setSpacing(10)
        self._page_header(
            layout, "PDF 압축",
            "이미지가 많은 PDF의 용량을 줄입니다. 압축 강도는 3단계 프리셋 중 선택합니다.")
        self._common_file_card(layout, "압축할 PDF 파일", self._browse)

        card, vl = _card("압축 강도")
        row = QHBoxLayout()
        for name in ("낮음", "보통", "높음"):
            btn = QPushButton(self.PRESET_TEXT[name])
            btn.setFont(_font(9, "medium"))
            btn.setFixedHeight(48)
            btn.setStyleSheet(BTN_ON if name == self._quality else BTN_SOFT)
            btn.clicked.connect(lambda _, n=name: self._select_quality(n))
            row.addWidget(btn)
            self._quality_btns[name] = btn
        vl.addLayout(row)
        layout.addWidget(card)

        self._progress_card(layout)
        self.start_btn = QPushButton("압축 시작")
        self.start_btn.setFont(_font(11, "semibold"))
        self.start_btn.setFixedHeight(32)
        self.start_btn.setFixedWidth(140)
        self.start_btn.setEnabled(False)
        self.start_btn.setStyleSheet(BTN_OFF)
        self.start_btn.clicked.connect(self._start)
        action_row = QHBoxLayout()
        action_row.setContentsMargins(6, 10, 6, 10)
        action_row.addStretch()
        action_row.addWidget(self.start_btn)
        action_row.addStretch()
        layout.addLayout(action_row)
        self._log_card(layout)
        layout.addStretch()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(_make_scroll(root))

    def _select_quality(self, name: str) -> None:
        self._quality = name
        for n, btn in self._quality_btns.items():
            btn.setStyleSheet(BTN_ON if n == name else BTN_SOFT)

    def _browse(self) -> None:
        path = self._browse_pdf()
        if path:
            self._load_file(path)

    def _load_file(self, path: str) -> None:
        # 파일 선택/드롭 공통 진입점.
        self.file_edit.setText(path)
        size_mb = os.path.getsize(path) / 1024 / 1024
        self.info_lbl.setText(f"{os.path.basename(path)} | {size_mb:.2f} MB")
        self._set_action_ready(self.start_btn, True)

    def _handle_dropped_paths(self, paths: list[str]) -> None:
        if self.is_running():
            return
        for p in paths:
            if os.path.isfile(p) and p.lower().endswith(".pdf"):
                self._load_file(p)
                return

    def _start(self) -> None:
        path = self.file_edit.text()
        if not path:
            return
        jpeg_q, dpi = self.PRESETS[self._quality]
        self.log_box.clear()
        self.progress.setValue(0)
        self._set_action_ready(self.start_btn, False)
        self._worker = PdfCompressWorker(path, self._quality, jpeg_q, dpi)
        self._thread = QThread()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self.log_box.append)
        self._worker.progress.connect(self._on_progress)
        self._worker.done.connect(self._on_done)
        self._worker.error.connect(self._on_error)
        self._worker.done.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.finished.connect(self._on_finished)
        self._thread.start()

    def _on_progress(self, current: int, total: int, status: str) -> None:
        pct = int(current / total * 100) if total else 0
        self.progress.setValue(pct)
        self.count_lbl.setText(f"{current} / {total} ({pct}%)")
        self.status_lbl.setText(status)

    def _on_done(self, out_path: str, before: float, after: float, pct: float, done: int, found: int) -> None:
        self.progress.setValue(100)
        self.status_lbl.setText(f"압축 완료 - {before:.1f} MB → {after:.1f} MB ({pct:.0f}% 감소)")
        _recent_add("PDF 압축", os.path.basename(out_path), f"{before:.1f} MB → {after:.1f} MB")
        QMessageBox.information(
            self,
            "완료",
            f"압축이 완료되었습니다.\n\n압축 전 {before:.1f} MB\n압축 후 {after:.1f} MB\n감소율 {pct:.0f}%\n이미지 {done}/{found}개 압축\n\n저장 위치:\n{out_path}",
        )

    def _on_error(self, msg: str) -> None:
        self.status_lbl.setText("오류 발생")
        self.log_box.append(f"[오류] {msg}")
        QMessageBox.critical(self, "오류", msg)
        self._offer_diag(msg)   # 동의 시 작업 로그를 GitHub 이슈로 전송

    def _on_finished(self) -> None:
        self._set_action_ready(self.start_btn, bool(self.file_edit.text()))
        self._worker = None
        self._thread = None

    def is_running(self) -> bool:
        return bool(self._thread and self._thread.isRunning())


# ─────────────────── PDF → SVG / EMF / PNG 내보내기 ───────────────────
# EMF 벡터 기록부(약 540줄)는 rtk_emf.py로 옮겼다. 동작은 그대로다.
from rtk_emf import pdf_page_to_emf, _emf_ref_device


class PdfExportWorker(QObject):
    # PDF 각 페이지를 SVG / EMF / PNG 파일로 하나씩 내보낸다.
    log = Signal(str)
    progress = Signal(int, int, str)
    done = Signal(int, str)
    error = Signal(str)
    cancelled = Signal()

    # 이미지 렌더 상한(픽셀 수). 넘으면 DPI를 자동으로 낮춘다.
    MAX_PNG_PIXELS = 120_000_000
    # 배경 투명은 채널 연산까지 하므로 메모리를 더 쓴다 → 상한을 따로 둔다.
    MAX_ALPHA_PIXELS = 60_000_000

    def __init__(self, input_path: str, fmt: str, opts: dict) -> None:
        super().__init__()
        self.input_path = input_path
        self.fmt = fmt                  # "svg" | "emf" | "png" | "jpg"
        self.opts = opts
        self._stop = False

    def request_stop(self) -> None:
        self._stop = True

    @Slot()
    def run(self) -> None:
        doc = None
        try:
            doc = fitz.open(self.input_path)
            total = len(doc)
            if total == 0:
                raise ValueError("PDF에 페이지가 없습니다.")
            stem = _safe_name(os.path.splitext(os.path.basename(self.input_path))[0])
            out_dir = _unique_path(os.path.join(os.path.dirname(self.input_path),
                                                f"{stem}_{self.fmt.upper()}"))
            os.makedirs(out_dir, exist_ok=True)
            if self.fmt == "emf":
                _hdc, rw, rh = _emf_ref_device()
                self.log.emit(f"기준 장치 해상도: {rw} x {rh}")

            made = 0
            for i in range(total):
                if self._stop:
                    self.log.emit("사용자 중지 요청")
                    self.cancelled.emit()
                    return
                page = doc[i]
                out = os.path.join(out_dir, f"{stem}_p{i + 1:03d}.{self.fmt}")
                if self.fmt == "svg":
                    self._write_svg(page, out)
                elif self.fmt == "emf":
                    st = pdf_page_to_emf(page, out)
                    self.log.emit(
                        f"{i + 1}쪽: 도형 {st['paths']} · 글자 {st['spans']} · 이미지 {st['images']}"
                        + (f" (건너뜀 {st['skipped']})" if st["skipped"] else ""))
                else:
                    self._write_image(page, out, i + 1)
                made += 1
                size_mb = os.path.getsize(out) / 1024 / 1024 if os.path.exists(out) else 0
                if self.fmt != "emf":
                    self.log.emit(f"{i + 1}쪽: {os.path.basename(out)} ({size_mb:.2f} MB)")
                if self.fmt == "png" and size_mb > 30:
                    self.log.emit(
                        f"    └ 사진이 들어간 쪽이라 PNG로는 용량이 큽니다. "
                        f"JPG로 저장하면 크게 줄어듭니다.")
                self.progress.emit(i + 1, total, f"{i + 1} / {total} 페이지 변환 완료")
            self.done.emit(made, out_dir)
        except Exception as exc:
            self.error.emit(str(exc))
        finally:
            if doc is not None:
                try:
                    doc.close()
                except Exception:
                    pass

    def _write_svg(self, page, out: str) -> None:
        svg = page.get_svg_image(text_as_path=bool(self.opts.get("text_as_path", True)))
        with open(out, "w", encoding="utf-8") as f:
            f.write(svg)

    def _write_image(self, page, out: str, page_no: int) -> None:
        # PNG는 무손실(선 도면에 유리), JPG는 손실 압축(사진·조감도에 유리).
        dpi = int(self.opts.get("dpi", 300))
        alpha = bool(self.opts.get("transparent", False)) and self.fmt == "png"
        # 메모리 보호: 픽셀 수 상한을 넘으면 DPI를 자동으로 낮춘다.
        w_in = page.rect.width / 72.0
        h_in = page.rect.height / 72.0
        px = (w_in * dpi) * (h_in * dpi)
        cap = self.MAX_ALPHA_PIXELS if alpha else self.MAX_PNG_PIXELS
        if px > cap:
            new_dpi = max(72, int((cap / max(w_in * h_in, 0.01)) ** 0.5))
            self.log.emit(f"{page_no}쪽: 해상도가 너무 높아 {dpi} → {new_dpi} DPI로 낮춥니다.")
            dpi = new_dpi
        # 투명 배경도 일단 흰 배경으로 렌더한다. MuPDF의 투명 배경 렌더는 PDF 투명
        # 그룹·불투명도(/ca) 요소가 합성 바탕(흰 종이)을 잃어 색이 옅고 탁해진다.
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        if alpha:
            self._white_to_alpha(pix).save(out)
            return
        if self.fmt == "jpg":
            pix.pil_save(out, format="JPEG", quality=int(self.opts.get("quality", 92)),
                         optimize=True, subsampling=0)
        else:
            pix.save(out)

    @staticmethod
    def _white_to_alpha(pix):
        # 흰 배경 위 합성 결과를 역산해 RGBA로. 흰색에서 멀수록 불투명하다고 보고
        #   a = 255 - min(r,g,b),  원색 = (c - min) * 255 / a
        # 로 되돌리면 흰 배경에 다시 얹었을 때 원본과 오차 1/255 이내로 일치한다.
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        r, g, b = img.split()
        mn = ImageChops.darker(ImageChops.darker(r, g), b)      # min(r,g,b)
        a = ImageChops.invert(mn)                               # 255 - min
        den = ImageChops.lighter(a, Image.new("L", img.size, 1))  # 0 나눗셈 방지
        # Pillow 10부터 ImageMath.eval이 빠져 lambda_eval을 쓴다(구버전 폴백 포함).
        _ev = getattr(ImageMath, "lambda_eval", None)
        if _ev is not None:
            chans = [_ev(lambda a: a["convert"](255 * (a["c"] - a["m"]) / a["d"], "L"),
                         c=c, m=mn, d=den) for c in (r, g, b)]
        else:
            chans = [ImageMath.eval("convert(255*(c-m)/d, 'L')", c=c, m=mn, d=den)
                     for c in (r, g, b)]
        return Image.merge("RGBA", chans + [a])


class PdfThumbWorker(QObject):
    """PDF 페이지를 썸네일 이미지로 백그라운드 렌더한다(큰 PDF에서 UI 멈춤 방지).

    jobs: [(token, src_path, page_index)] 목록. 페이지당 QImage 하나를 token과 함께
    emit하면 GUI 스레드가 그걸 QPixmap으로 만들어 해당 아이템 아이콘에 넣는다.
    (QPixmap은 GUI 스레드에서만 생성 가능하므로 워커는 QImage까지만 만든다.)
    """

    thumb = Signal(str, QImage)   # token, 렌더된 이미지
    done = Signal()

    # 150 DPI(중간 품질) 렌더 후 썸네일로 축소 → 글자 또렷하면서 메모리는 작게.
    RENDER_DPI = 150

    def __init__(self, jobs: list[tuple[str, str, int]]) -> None:
        super().__init__()
        self.jobs = jobs
        self._abort = False

    def stop(self) -> None:
        self._abort = True

    @Slot()
    def run(self) -> None:
        docs: dict[str, "fitz.Document"] = {}   # src별 fitz 문서 캐시(반복 open 방지)
        try:
            zoom = self.RENDER_DPI / 72.0
            mat = fitz.Matrix(zoom, zoom)
            for token, src, pno in self.jobs:
                if self._abort:
                    break
                doc = docs.get(src)
                if doc is None:
                    doc = fitz.open(src)
                    docs[src] = doc
                try:
                    page = doc.load_page(pno)
                    pix = page.get_pixmap(matrix=mat, alpha=False)
                    img = QImage(pix.samples, pix.width, pix.height,
                                 pix.stride, QImage.Format.Format_RGB888).copy()
                    self.thumb.emit(token, img)
                except Exception:
                    # 한 페이지 렌더 실패해도 전체는 계속 진행(빈 썸네일로 남음).
                    continue
        finally:
            for doc in docs.values():
                try:
                    doc.close()
                except Exception:
                    pass
            self.done.emit()


class PdfEditSaveWorker(QObject):
    """현재 페이지 순서·회전대로 새 PDF를 만들어 저장한다(백그라운드).

    descriptors: [(src_path, page_index, rotation_deg)] — 화면 리스트 순서 그대로.
    각 페이지를 원본 reader에서 가져와 writer에 추가하고, 회전 delta가 있으면 적용한다.
    """

    progress = Signal(int, int)   # current, total
    done = Signal(str)            # 저장 경로
    error = Signal(str)

    def __init__(self, descriptors: list[tuple[str, int, int]], out_path: str) -> None:
        super().__init__()
        self.descriptors = descriptors
        self.out_path = out_path

    @Slot()
    def run(self) -> None:
        try:
            writer = PdfWriter()
            readers: dict[str, PdfReader] = {}
            total = len(self.descriptors)
            for i, (src, pno, rotation) in enumerate(self.descriptors, 1):
                reader = readers.get(src)
                if reader is None:
                    reader = PdfReader(src)
                    readers[src] = reader
                writer.add_page(reader.pages[pno])
                if rotation % 360:
                    # writer에 복사된 페이지에만 회전 적용(원본 reader 객체 오염 방지).
                    writer.pages[-1].rotate(rotation % 360)
                self.progress.emit(i, total)
            with open(self.out_path, "wb") as f:
                writer.write(f)
            self.done.emit(self.out_path)
        except Exception as exc:
            self.error.emit(str(exc))


class _VResizeHandle(QWidget):
    """수직 드래그로 대상 위젯(썸네일 격자) 높이를 조절하는 손잡이 막대.

    PDF 편집 탭은 전체가 QScrollArea 안에 있어 QSplitter를 쓰기 어렵다.
    그래서 리스트 아래에 얇은 그립 바를 두고, 드래그한 만큼 대상의 고정
    높이를 늘리거나 줄인다(min_h~max_h 범위로 제한).
    """

    def __init__(self, target: QWidget, min_h: int, max_h: int, parent=None) -> None:
        super().__init__(parent)
        self._target = target
        self._min_h = min_h
        self._max_h = max_h
        self._drag_y: float | None = None   # 드래그 시작 시점의 전역 Y 좌표
        self._start_h = 0                    # 드래그 시작 시점의 대상 높이
        self.setFixedHeight(14)
        self.setCursor(Qt.CursorShape.SizeVerCursor)
        self.setToolTip("드래그하여 미리보기 영역 크기 조절")

    def paintEvent(self, e) -> None:
        # 가운데 그립 점 3개 — 드래그 가능하다는 시각적 힌트.
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor("#C2C8DA"))
        cx = self.width() // 2
        cy = self.height() // 2
        for dx in (-9, 0, 9):
            p.drawEllipse(cx + dx - 2, cy - 2, 4, 4)

    def mousePressEvent(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton:
            self._drag_y = e.globalPosition().y()
            self._start_h = self._target.height()

    def mouseMoveEvent(self, e) -> None:
        if self._drag_y is None:
            return
        # 시작점 대비 이동량만큼 대상 높이를 가감(범위 제한).
        dy = e.globalPosition().y() - self._drag_y
        new_h = int(max(self._min_h, min(self._max_h, self._start_h + dy)))
        self._target.setFixedHeight(new_h)

    def mouseReleaseEvent(self, e) -> None:
        self._drag_y = None


class PdfEditTab(BaseTool):
    """PDF 페이지 편집: 썸네일로 보면서 삭제·회전, 다른 PDF 끼워넣기, 드래그로 순서변경.

    - 원본 PDF를 열면 페이지가 격자 썸네일로 펼쳐진다.
    - [PDF 추가]로 다른 PDF의 페이지들을 낱장 썸네일로 끼워넣는다(각 페이지 개별 이동).
    - 썸네일을 드래그해 원하는 위치로 재배치(InternalMove).
    - 선택 후 Delete(또는 버튼)로 삭제, 좌/우 회전 버튼으로 90° 회전.
    - [새 PDF로 저장]은 현재 화면 순서·회전대로 _편집.pdf를 만든다(원본 보존).
    """

    TOOL_LABEL = "PDF 편집"

    # 썸네일/격자 칸 기본 크기(px). 아이콘 + 아래 캡션 공간 확보. Ctrl+휠로 가변.
    THUMB_W, THUMB_H = 124, 175
    THUMB_ASPECT = THUMB_H / THUMB_W   # 세로/가로 비(A4)
    THUMB_W_MIN = 84                   # 칸 최소 너비(=최대 열수 한도)
    # 썸네일을 한 번만 이 기준 해상도로 렌더해 캐시. 줌은 iconSize만 바꿔 뷰가 자동 축소표시
    # → 휠마다 재렌더 없음(렉 제거). 크게 확대 시(1열) 약간 부드러워지나 읽을 만함.
    BASE_CACHE_W = 440

    def __init__(self) -> None:
        super().__init__()
        self._thumb_worker: PdfThumbWorker | None = None
        self._thumb_thread: QThread | None = None
        self._save_worker: PdfEditSaveWorker | None = None
        self._save_thread: QThread | None = None
        self._pix_cache: dict[tuple[str, int], QPixmap] = {}   # (src,page)→기준해상도 썸네일 캐시
        self._token_items: dict[str, QListWidgetItem] = {}     # 렌더 토큰→아이템
        self._token_seq = 0
        # 줌은 인스턴스 속성으로 가변(클래스 기본값을 복사). 모든 self.THUMB_* 참조가 이 값을 씀.
        self.THUMB_W = PdfEditTab.THUMB_W
        self.THUMB_H = PdfEditTab.THUMB_H
        self._cols: int | None = None   # 현재 한 줄 열 수(Ctrl+휠로 ±1 스냅)
        self._build_ui()
        self.setAcceptDrops(True)   # PDF 드롭으로 열기/추가

    # ── 드래그 앤 드롭(파일) ──
    def _handle_dropped_paths(self, paths: list[str]) -> None:
        pdfs = [p for p in paths if os.path.isfile(p) and p.lower().endswith(".pdf")]
        if not pdfs:
            return
        if not self._has_pages():
            # 비어 있으면 첫 PDF를 원본으로 연다(나머지는 이어서 추가).
            self._load_pdf(pdfs[0], as_base=True)
            for p in pdfs[1:]:
                self._load_pdf(p, as_base=False)
        else:
            for p in pdfs:
                self._load_pdf(p, as_base=False)

    def reset_state(self) -> None:
        if self.is_running():
            return
        self._stop_thumb_render()
        self.listw.clear()
        self._pix_cache.clear()
        self._token_items.clear()
        self.info_lbl.setText("PDF를 선택하거나 이 창에 끌어다 놓으세요.")
        self.status_lbl.setText("대기 중")
        self.progress.setValue(0)
        self._refresh_buttons()

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("scroll_content")
        layout = QVBoxLayout(root)
        layout.setContentsMargins(20, 16, 20, 10)
        layout.setSpacing(10)
        self._page_header(
            layout, "PDF 편집",
            "페이지를 썸네일로 보며 삭제·회전하고, 다른 PDF를 끼워넣어 드래그로 순서를 바꿉니다.")

        # ── 파일 열기 / 추가 ──
        card, vl = _card("PDF 열기")
        row = QHBoxLayout()
        self.open_btn = QPushButton("PDF 열기")
        self.open_btn.setFont(_font(9, "medium"))
        self.open_btn.setFixedHeight(34)
        self.open_btn.setStyleSheet(BTN_BROWSE)
        self.open_btn.clicked.connect(self._browse_open)
        self.add_btn = QPushButton("＋ PDF 추가(끼워넣기)")
        self.add_btn.setFont(_font(9, "medium"))
        self.add_btn.setFixedHeight(34)
        self.add_btn.setStyleSheet(BTN_BROWSE)
        self.add_btn.clicked.connect(self._browse_add)
        row.addWidget(self.open_btn)
        row.addWidget(self.add_btn)
        row.addStretch()
        vl.addLayout(row)
        self.info_lbl = QLabel("PDF를 선택하거나 이 창에 끌어다 놓으세요.")
        self.info_lbl.setFont(_font(8, "regular"))
        self.info_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        vl.addWidget(self.info_lbl)
        layout.addWidget(card)

        # ── 페이지 썸네일 격자 ──
        card, vl = _card("페이지 (드래그로 순서변경 · 더블클릭 크게보기)")
        edit_row = QHBoxLayout()
        edit_row.setSpacing(8)
        self.del_btn = QPushButton("선택 삭제")
        self.rotl_btn = QPushButton("⟲ 왼쪽 회전")
        self.rotr_btn = QPushButton("⟳ 오른쪽 회전")
        # 좌우 패딩이 있는 알약 버튼(글자가 박스 밖으로 삐져나오던 문제 보정).
        edit_btn_qss = (
            "QPushButton{background:#E9EAF6;color:#5A6075;border:none;"
            "border-radius:8px;padding:5px 16px;}"
            f"QPushButton:hover{{background:{ACCENT_SOFT};color:{ACCENT};}}"
            "QPushButton:disabled{color:#B9BECF;}"
        )
        for b in (self.del_btn, self.rotl_btn, self.rotr_btn):
            b.setFont(_font(9, "medium"))
            b.setMinimumHeight(30)
            b.setSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)
            b.setStyleSheet(edit_btn_qss)
            edit_row.addWidget(b)
        edit_row.addStretch()
        self.sel_lbl = QLabel("")
        self.sel_lbl.setFont(_font(8, "regular"))
        self.sel_lbl.setStyleSheet(f"color:{TXT_T};background:transparent;")
        edit_row.addWidget(self.sel_lbl)
        self.del_btn.clicked.connect(self._delete_selected)
        self.rotl_btn.clicked.connect(lambda: self._rotate_selected(-90))
        self.rotr_btn.clicked.connect(lambda: self._rotate_selected(90))
        vl.addLayout(edit_row)

        # 격자(IconMode) + 내부 드래그 재배치. Static movement라 자유배치 대신 순서이동만 됨.
        # MoveAction을 기본 드롭동작으로 지정해야 드래그 시 '금지' 커서 없이 재배치된다.
        self.listw = ThumbGridList()
        self.listw.setViewMode(QListView.ViewMode.IconMode)
        self.listw.setFlow(QListView.Flow.LeftToRight)
        self.listw.setWrapping(True)
        self.listw.setResizeMode(QListView.ResizeMode.Adjust)
        self.listw.setMovement(QListView.Movement.Static)
        self.listw.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.listw.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.listw.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.listw.setDragEnabled(True)
        self.listw.viewport().setAcceptDrops(True)
        self.listw.setDropIndicatorShown(True)
        # 드래그를 격자 가장자리로 가져가면 자동 스크롤(위/아래로 따라 스크롤).
        self.listw.setAutoScroll(True)
        self.listw.setAutoScrollMargin(56)
        self.listw.setIconSize(QSize(self.THUMB_W - 16, self.THUMB_H - 34))
        self.listw.setGridSize(QSize(self.THUMB_W, self.THUMB_H))
        self.listw.setUniformItemSizes(True)
        self.listw.setSpacing(6)
        self.listw.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        # 초기 높이 = 3줄(썸네일 높이×3 + 여백). 아래 드래그 손잡이로 사용자가 키우거나 줄임.
        self.listw.setFixedHeight(self.THUMB_H * 3 + 40)
        self.listw.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        # ::item에 background를 지정하지 않아 끼워넣은 페이지의 연한 빨강 배경(BackgroundRole)이 보이게 함.
        self.listw.setStyleSheet(
            f"QListWidget {{ background:#FBFBFE; border:1px solid {BORDER};"
            f" border-radius:8px; outline:none; padding:6px; }}"
            f"QListWidget::item {{ border:1px solid {BORDER}; border-radius:6px;"
            f" color:{TXT_S}; padding-bottom:3px; }}"
            f"QListWidget::item:selected {{ background:{ACCENT_SOFT};"
            f" border:1px solid {ACCENT}; color:{TXT}; }}"
            "QScrollBar:vertical { background:transparent; width:9px; margin:4px 2px 4px 2px; }"
            "QScrollBar::handle:vertical { background:#CAD4E4; min-height:28px; border-radius:4px; }"
            "QScrollBar::handle:vertical:hover { background:#AAB8CC; }"
            "QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,"
            "QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background:transparent; border:none; height:0px; }"
        )
        self.listw.setItemDelegate(ThumbDelegate(self.listw))
        self.listw.itemSelectionChanged.connect(self._on_sel_changed)
        self.listw.itemDoubleClicked.connect(self._preview_item)
        self.listw.reordered.connect(self._on_reordered)
        self.listw.zoomRequested.connect(self._on_zoom)
        vl.addWidget(self.listw)
        # 미리보기(썸네일 격자) 영역 크기를 세로 드래그로 조절하는 손잡이.
        self.list_resize = _VResizeHandle(
            self.listw, min_h=self.THUMB_H + 60, max_h=self.THUMB_H * 10)
        vl.addWidget(self.list_resize)
        hint = QLabel("Tip: 여러 장은 Ctrl/Shift로 다중선택. Delete 키로 삭제. 아래 ⋯ 손잡이로 영역 크기, Ctrl+휠로 페이지 크게/작게.")
        hint.setFont(_font(8, "regular"))
        hint.setStyleSheet(f"color:{TXT_T};background:transparent;")
        vl.addWidget(hint)
        layout.addWidget(card)

        # ── 진행 + 저장 ──
        prog_card, pvl = _card("진행 상황")
        self.status_lbl = QLabel("대기 중")
        self.status_lbl.setFont(_font(9, "regular"))
        self.status_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        pvl.addWidget(self.status_lbl)
        self.progress = QProgressBar()
        self.progress.setFixedHeight(6)
        self.progress.setTextVisible(False)
        pvl.addWidget(self.progress)
        layout.addWidget(prog_card)

        self.save_btn = QPushButton("새 PDF로 저장")
        self.save_btn.setFont(_font(11, "semibold"))
        self.save_btn.setFixedHeight(32)
        self.save_btn.setFixedWidth(160)
        self.save_btn.setEnabled(False)
        self.save_btn.setStyleSheet(BTN_OFF)
        self.save_btn.clicked.connect(self._save)
        action_row = QHBoxLayout()
        action_row.setContentsMargins(6, 10, 6, 10)
        action_row.addStretch()
        action_row.addWidget(self.save_btn)
        action_row.addStretch()
        layout.addLayout(action_row)

        self._log_card(layout)
        layout.addStretch()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(_make_scroll(root))

    def _log_card(self, layout: QVBoxLayout) -> None:
        card, vl = _card("작업 로그")
        self.log_box = LogoLogTextEdit()
        self.log_box.setObjectName("log_box")
        self.log_box.setReadOnly(True)
        self.log_box.setFont(_font(9, "light"))
        self.log_box.setFixedHeight(130)
        vl.addWidget(self.log_box)
        layout.addWidget(card)

    # ── 키보드: Delete로 선택 삭제 ──
    def keyPressEvent(self, e) -> None:
        if e.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace) and self.listw.hasFocus():
            self._delete_selected()
            return
        super().keyPressEvent(e)

    # ── 파일 열기/추가 ──
    def _browse_open(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "PDF 열기", "", "PDF 파일 (*.pdf)")
        if path:
            self._stop_thumb_render()
            self.listw.clear()
            self._token_items.clear()
            self._load_pdf(path, as_base=True)

    def _browse_add(self) -> None:
        if not self._has_pages():
            QMessageBox.information(self, "안내", "먼저 PDF를 연 다음 추가하세요.")
            return
        paths, _ = QFileDialog.getOpenFileNames(self, "끼워넣을 PDF 선택", "", "PDF 파일 (*.pdf)")
        if not paths:
            return
        insert_at = self._ask_insert_position()
        if insert_at is None:   # 취소
            return
        # 여러 PDF를 고르면 고른 순서대로 같은 위치에 이어서 끼워넣는다.
        pos = insert_at
        for path in paths:
            before = self.listw.count()
            self._load_pdf(path, as_base=False, insert_at=pos)
            pos += self.listw.count() - before

    def _ask_insert_position(self) -> int | None:
        """끼워넣을 위치를 묻는 팝업. 반환: 삽입할 행 인덱스, 취소 시 None.

        - '맨 뒤에 추가' → 현재 페이지 수(맨 끝).
        - 'N쪽 뒤에 삽입' → 인덱스 N (N=0이면 맨 앞). N은 0~현재페이지수.
        """
        total = self.listw.count()
        dlg = QDialog(self)
        dlg.setWindowTitle("끼워넣을 위치")
        dlg.setMinimumWidth(380)
        dv = QVBoxLayout(dlg)
        dv.setContentsMargins(20, 18, 20, 16)
        dv.setSpacing(12)
        info = QLabel(f"현재 총 {total}쪽입니다. 새 PDF를 어디에 끼워넣을까요?")
        info.setFont(_font(9, "regular"))
        dv.addWidget(info)

        grp = QButtonGroup(dlg)
        rb_end = QRadioButton("맨 뒤에 추가")
        rb_end.setChecked(True)
        rb_at = QRadioButton("특정 쪽 뒤에 삽입")
        grp.addButton(rb_end)
        grp.addButton(rb_at)
        dv.addWidget(rb_end)

        at_row = QHBoxLayout()
        at_row.addWidget(rb_at)
        page_edit = QLineEdit(str(total))
        page_edit.setFixedWidth(60)
        page_edit.setValidator(QIntValidator(0, total, dlg))
        at_row.addWidget(page_edit)
        at_row.addWidget(QLabel("쪽 뒤  (0 = 맨 앞)"))
        at_row.addStretch()
        dv.addLayout(at_row)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        btn_row.addStretch()
        cancel = QPushButton("취소")
        # 알약 버튼에 좌우 패딩 + 최소 폭(글자가 박스 밖으로 삐져나오던 문제 보정).
        cancel.setStyleSheet(
            "QPushButton{background:#E9EAF6;color:#5A6075;border:none;"
            "border-radius:9px;padding:6px 22px;}"
            f"QPushButton:hover{{background:{ACCENT_SOFT};color:{ACCENT};}}"
        )
        cancel.setMinimumSize(96, 34)
        ok = QPushButton("확인")
        ok.setStyleSheet(BTN_BROWSE)
        ok.setMinimumSize(96, 34)
        ok.setFont(_font(9, "semibold"))
        cancel.clicked.connect(dlg.reject)
        ok.clicked.connect(dlg.accept)
        # 쪽수 칸을 만지면 자동으로 '특정 쪽 뒤' 라디오 선택.
        page_edit.textEdited.connect(lambda *_: rb_at.setChecked(True))
        btn_row.addWidget(cancel)
        btn_row.addWidget(ok)
        dv.addLayout(btn_row)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return None
        if rb_end.isChecked():
            return total
        try:
            n = max(0, min(int(page_edit.text() or 0), total))
        except ValueError:
            n = total
        return n

    def _load_pdf(self, path: str, as_base: bool, insert_at: int | None = None) -> None:
        """PDF의 전체 페이지를 썸네일 아이템으로 리스트에 추가/삽입한다.

        insert_at=None이면 맨 뒤에 추가, 정수면 그 위치(행 인덱스)부터 끼워넣는다.
        """
        try:
            doc = fitz.open(path)
            n = doc.page_count
            doc.close()
        except Exception as exc:
            QMessageBox.critical(self, "오류", f"PDF를 열 수 없습니다.\n{exc}")
            return
        if n == 0:
            return
        name = os.path.basename(path)
        short = (name[:14] + "…") if len(name) > 15 else name
        jobs: list[tuple[str, str, int]] = []
        pos = self.listw.count() if insert_at is None else max(0, min(insert_at, self.listw.count()))
        for pno in range(n):
            self._token_seq += 1
            token = str(self._token_seq)
            it = QListWidgetItem()
            it.setTextAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignBottom)
            it.setData(Qt.ItemDataRole.UserRole,
                       {"src": path, "page": pno, "rotation": 0, "token": token,
                        "inserted": not as_base, "short": short})
            # 드래그 재배치를 위해 드래그/드롭 둘 다 활성(이전엔 드롭을 꺼서 '금지' 커서가 떴음).
            it.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled
                        | Qt.ItemFlag.ItemIsDragEnabled | Qt.ItemFlag.ItemIsDropEnabled)
            cached = self._pix_cache.get((path, pno))
            if cached is not None:
                it.setIcon(QIcon(cached))
            else:
                it.setIcon(self._placeholder_icon())
                jobs.append((token, path, pno))
            self.listw.insertItem(pos, it)
            pos += 1
            self._token_items[token] = it
        where = "맨 뒤" if insert_at is None else f"{insert_at}번 위치"
        self.log_box.append(f"[{'추가' if as_base else '끼워넣기'}] {name} — {n}쪽 → {where}")
        self._renumber()
        self._update_info()
        self._refresh_buttons()
        if jobs:
            self._start_thumb_render(jobs)

    def _renumber(self) -> None:
        """현재 배치 순서대로 페이지 번호(1쪽, 2쪽…)를 다시 매겨 캡션을 갱신한다."""
        for i in range(self.listw.count()):
            it = self.listw.item(i)
            d = it.data(Qt.ItemDataRole.UserRole)
            it.setText(f"{i + 1}쪽\n{d.get('short', '')}")

    def _on_reordered(self) -> None:
        # 드래그로 순서가 바뀌면 번호를 다시 매긴다.
        self._renumber()

    def _placeholder_icon(self) -> QIcon:
        # 렌더 전 임시 흰 카드 아이콘.
        pm = QPixmap(self.THUMB_W - 16, self.THUMB_H - 34)
        pm.fill(QColor("#F0F1F8"))
        return QIcon(pm)

    # ── 썸네일 백그라운드 렌더 ──
    def _start_thumb_render(self, jobs: list[tuple[str, str, int]]) -> None:
        # 이미 렌더 중이면 새 작업을 잇기 위해 멈추고 남은 미렌더 페이지까지 다시 모은다.
        if self._thumb_thread and self._thumb_thread.isRunning():
            self._stop_thumb_render()
            jobs = self._pending_thumb_jobs()
        if not jobs:
            return
        self.status_lbl.setText("썸네일 생성 중…")
        self._thumb_worker = PdfThumbWorker(jobs)
        self._thumb_thread = QThread()
        self._thumb_worker.moveToThread(self._thumb_thread)
        self._thumb_thread.started.connect(self._thumb_worker.run)
        self._thumb_worker.thumb.connect(self._on_thumb)
        self._thumb_worker.done.connect(self._on_thumb_done)
        self._thumb_worker.done.connect(self._thumb_thread.quit)
        self._thumb_thread.start()

    def _pending_thumb_jobs(self) -> list[tuple[str, str, int]]:
        # 아직 캐시에 없는(=플레이스홀더) 아이템들의 렌더 작업 목록.
        jobs = []
        for i in range(self.listw.count()):
            d = self.listw.item(i).data(Qt.ItemDataRole.UserRole)
            if (d["src"], d["page"]) not in self._pix_cache:
                jobs.append((d["token"], d["src"], d["page"]))
        return jobs

    def _stop_thumb_render(self) -> None:
        if self._thumb_worker:
            self._thumb_worker.stop()
        if self._thumb_thread and self._thumb_thread.isRunning():
            self._thumb_thread.quit()
            self._thumb_thread.wait(3000)
        self._thumb_worker = None
        self._thumb_thread = None

    @Slot(str, QImage)
    def _on_thumb(self, token: str, img: QImage) -> None:
        it = self._token_items.get(token)
        if it is None:
            return
        d = it.data(Qt.ItemDataRole.UserRole)
        # 기준 해상도(BASE_CACHE_W)로 한 번만 축소 저장. 화면 표시 크기는 뷰(iconSize)가
        # 이 캐시본을 그때그때 스케일 → 줌해도 재렌더 불필요.
        pm = QPixmap.fromImage(img).scaled(
            self.BASE_CACHE_W, round(self.BASE_CACHE_W * self.THUMB_ASPECT),
            Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self._pix_cache[(d["src"], d["page"])] = pm
        self._apply_icon(it, pm, d.get("rotation", 0))

    def _apply_icon(self, it: QListWidgetItem, base_pm: QPixmap, rotation: int) -> None:
        # 회전 각도를 반영해 아이콘을 그린다(원본 썸네일은 캐시 그대로 두고 표시만 회전).
        pm = base_pm
        if rotation % 360:
            pm = base_pm.transformed(QTransform().rotate(rotation), Qt.TransformationMode.SmoothTransformation)
        it.setIcon(QIcon(pm))

    def _on_thumb_done(self) -> None:
        self.status_lbl.setText(f"준비 완료 — 총 {self.listw.count()}쪽")

    # ── Ctrl+휠 줌(한 줄 열 수 ±1 스냅) ──
    def _grid_inner_w(self) -> int:
        # 아이템이 깔리는 실제 가로 폭(뷰포트 - 좌우 패딩 - 스크롤바 여유).
        return max(self.THUMB_W_MIN, self.listw.viewport().width() - 12)

    def _max_cols(self) -> int:
        # 칸 최소 너비 기준 한 줄에 들어갈 수 있는 최대 열 수(=최대 축소).
        s = self.listw.spacing()
        return max(1, (self._grid_inner_w() - s) // (self.THUMB_W_MIN + s))

    def _current_cols(self) -> int:
        # 현재 칸 너비로부터 역산한 열 수.
        s = self.listw.spacing()
        return max(1, (self._grid_inner_w() - s) // max(1, self.THUMB_W + s))

    def _apply_cols(self, cols: int) -> None:
        # cols개가 한 줄에 딱 들어가도록 칸 크기를 정해 격자에 적용(재렌더 없음).
        cols = max(1, min(cols, self._max_cols()))
        inner = self._grid_inner_w()
        s = self.listw.spacing()
        # cols*w + (cols+1)*spacing <= inner 가 되도록 칸 너비 산출(스페이싱만큼 여유).
        w = max(self.THUMB_W_MIN, (inner - (cols + 1) * s) // cols)
        self._cols = cols
        self.THUMB_W = w
        self.THUMB_H = round(w * self.THUMB_ASPECT)
        # gridSize=칸 전체, iconSize=칸 안 이미지(좌우16·하단34는 테두리/캡션 공간).
        self.listw.setIconSize(QSize(self.THUMB_W - 16, self.THUMB_H - 34))
        self.listw.setGridSize(QSize(self.THUMB_W, self.THUMB_H))

    @Slot(int)
    def _on_zoom(self, direction: int) -> None:
        # 확대(휠 위)=열 1개 줄여 칸 키움 / 축소(휠 아래)=열 1개 늘림.
        cur = self._cols if self._cols is not None else self._current_cols()
        self._apply_cols(cur - 1 if direction > 0 else cur + 1)

    # ── 편집 동작 ──
    def _delete_selected(self) -> None:
        rows = sorted((self.listw.row(it) for it in self.listw.selectedItems()), reverse=True)
        if not rows:
            return
        for r in rows:
            it = self.listw.takeItem(r)
            d = it.data(Qt.ItemDataRole.UserRole)
            self._token_items.pop(d.get("token", ""), None)
        self.log_box.append(f"[삭제] {len(rows)}쪽 제거")
        self._renumber()
        self._update_info()
        self._refresh_buttons()

    def _rotate_selected(self, deg: int) -> None:
        items = self.listw.selectedItems()
        if not items:
            return
        for it in items:
            d = it.data(Qt.ItemDataRole.UserRole)
            d["rotation"] = (d.get("rotation", 0) + deg) % 360
            it.setData(Qt.ItemDataRole.UserRole, d)
            cached = self._pix_cache.get((d["src"], d["page"]))
            if cached is not None:
                self._apply_icon(it, cached, d["rotation"])
        self.log_box.append(f"[회전] {len(items)}쪽 {deg:+d}°")

    def _preview_item(self, it: QListWidgetItem) -> None:
        # 더블클릭: 해당 페이지를 크게 렌더해 팝업으로 보여준다.
        d = it.data(Qt.ItemDataRole.UserRole)
        try:
            doc = fitz.open(d["src"])
            page = doc.load_page(d["page"])
            pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888).copy()
            doc.close()
        except Exception as exc:
            QMessageBox.warning(self, "미리보기 실패", str(exc))
            return
        pm = QPixmap.fromImage(img)
        if d.get("rotation", 0) % 360:
            pm = pm.transformed(QTransform().rotate(d["rotation"]), Qt.TransformationMode.SmoothTransformation)
        dlg = QDialog(self)
        dlg.setWindowTitle(f"{os.path.basename(d['src'])} — {d['page'] + 1}쪽")
        dv = QVBoxLayout(dlg)
        lbl = QLabel()
        lbl.setPixmap(pm.scaled(720, 960, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        dv.addWidget(lbl)
        dlg.exec()

    def _on_sel_changed(self) -> None:
        sel = self.listw.selectedItems()
        # 선택(클릭)된 끼워넣기 페이지는 '새 페이지' 표시(연빨강)를 해제 → 원본과 동일 처리.
        changed = False
        for it in sel:
            d = it.data(Qt.ItemDataRole.UserRole)
            if d.get("inserted"):
                d["inserted"] = False
                it.setData(Qt.ItemDataRole.UserRole, d)
                changed = True
        if changed:
            self.listw.viewport().update()
        self.sel_lbl.setText(f"{len(sel)}쪽 선택" if sel else "")

    # ── 상태/버튼 ──
    def _has_pages(self) -> bool:
        return self.listw.count() > 0

    def _update_info(self) -> None:
        n = self.listw.count()
        srcs = set()
        for i in range(n):
            srcs.add(self.listw.item(i).data(Qt.ItemDataRole.UserRole)["src"])
        self.info_lbl.setText(f"총 {n}쪽 · 원본 {len(srcs)}개 PDF" if n else "PDF를 선택하거나 이 창에 끌어다 놓으세요.")

    def _refresh_buttons(self) -> None:
        has = self._has_pages() and not self.is_running()
        for b in (self.del_btn, self.rotl_btn, self.rotr_btn):
            b.setEnabled(has)
        if has:
            self.save_btn.setEnabled(True)
            self._set_action_ready(self.save_btn, True)
        else:
            self.save_btn.setEnabled(False)
            self.save_btn.setStyleSheet(BTN_OFF)
            self._stop_pulse(self.save_btn)

    # ── 저장 ──
    def _save(self) -> None:
        if not self._has_pages() or self.is_running():
            return
        # 기본 저장명: 첫 원본 이름 + _편집.pdf
        first_src = self.listw.item(0).data(Qt.ItemDataRole.UserRole)["src"]
        base = os.path.splitext(os.path.basename(first_src))[0]
        default = os.path.join(os.path.dirname(first_src), f"{base}_편집.pdf")
        out, _ = QFileDialog.getSaveFileName(self, "새 PDF로 저장", default, "PDF 파일 (*.pdf)")
        if not out:
            return
        descriptors = []
        for i in range(self.listw.count()):
            d = self.listw.item(i).data(Qt.ItemDataRole.UserRole)
            descriptors.append((d["src"], d["page"], d.get("rotation", 0)))
        self.progress.setValue(0)
        self.status_lbl.setText("저장 중…")
        self._refresh_buttons()
        self._set_running_buttons(self.save_btn, True)
        self._save_worker = PdfEditSaveWorker(descriptors, out)
        self._save_thread = QThread()
        self._save_worker.moveToThread(self._save_thread)
        self._save_thread.started.connect(self._save_worker.run)
        self._save_worker.progress.connect(self._on_save_progress)
        self._save_worker.done.connect(self._on_save_done)
        self._save_worker.error.connect(self._on_save_error)
        self._save_worker.done.connect(self._save_thread.quit)
        self._save_worker.error.connect(self._save_thread.quit)
        self._save_thread.finished.connect(self._on_save_finished)
        self._save_thread.start()

    def _on_save_progress(self, cur: int, total: int) -> None:
        self.progress.setValue(int(cur / total * 100) if total else 0)

    def _on_save_done(self, out: str) -> None:
        self.progress.setValue(100)
        self.status_lbl.setText("저장 완료")
        _recent_add("PDF 편집", f"{self.listw.count()}쪽 저장", os.path.basename(out))
        self.log_box.append(f"[완료] {out}")
        QMessageBox.information(self, "완료", f"저장되었습니다.\n\n{out}")

    def _on_save_error(self, msg: str) -> None:
        self.status_lbl.setText("오류 발생")
        self.log_box.append(f"[오류] {msg}")
        QMessageBox.critical(self, "오류", msg)
        self._offer_diag(msg)

    def _on_save_finished(self) -> None:
        self._save_worker = None
        self._save_thread = None
        self._set_running_buttons(self.save_btn, False)
        self._refresh_buttons()

    def is_running(self) -> bool:
        return bool(self._save_thread and self._save_thread.isRunning())


class PdfExportTab(PdfBaseTab):
    # PDF 페이지를 SVG / EMF(벡터) 또는 고해상도 PNG로 한 장씩 내보내는 탭.
    TOOL_LABEL = "PDF 벡터 변환"
    DPI_CHOICES = (150, 300, 600, 900, 1200)

    def __init__(self) -> None:
        super().__init__()
        self._worker: PdfExportWorker | None = None
        self._thread: QThread | None = None
        self._fmt = "emf"
        self._dpi = 600
        self._img_kind = "png"
        self._dpi_btns: dict = {}
        self._kind_btns: dict = {}
        self._build_ui()
        self.setAcceptDrops(True)   # PDF 드롭으로 파일 지정

    def _handle_dropped_paths(self, paths: list) -> None:
        if self.is_running():
            return
        for p in paths:
            if os.path.isfile(p) and p.lower().endswith(".pdf"):
                self._load_file(p)
                return

    def reset_state(self) -> None:
        if self.is_running():
            return
        self.file_edit.clear()
        self.info_lbl.setText("-")
        self._select_fmt("emf")
        self._select_dpi(600)
        self._select_img_kind("png")
        self.path_chk.setChecked(True)
        self.alpha_chk.setChecked(False)
        self.progress.setValue(0)
        self.status_lbl.setText("파일을 선택해주세요.")
        self.count_lbl.clear()
        self.log_box.clear()
        self._set_action_ready(self.start_btn, False)

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("scroll_content")
        layout = QVBoxLayout(root)
        layout.setContentsMargins(20, 16, 20, 10)
        layout.setSpacing(10)
        self._page_header(
            layout, "PDF 벡터 변환",
            "도면·그림 PDF를 한 쪽씩 SVG·EMF(벡터) 또는 고해상도 PNG로 뽑아 "
            "한글·PPT에 깨지지 않게 넣습니다.")
        self._common_file_card(layout, "내보낼 PDF 파일", self._browse)

        # ── 출력 형식 ──
        card, vl = _card("출력 형식")
        self.fmt_group = QButtonGroup(self)
        self._fmt_radios = {}
        for key, title, desc in (
            ("emf", "EMF (벡터)",
             "한글·오피스 표준 벡터 형식. 확대해도 안 깨지고 크기 조절이 자유롭습니다."),
            ("svg", "SVG (벡터)",
             "웹·일러스트레이터용 표준 벡터 형식. 원본 그대로 무손실로 뽑습니다."),
            ("png", "PNG (고해상도 이미지)",
             "선이 아주 많은 도면처럼 벡터로 넣으면 PPT가 버거워질 때 쓰는 방법입니다."),
        ):
            row = QVBoxLayout()
            row.setSpacing(1)
            rb = QRadioButton(title)
            rb.setFont(_font(9, "medium"))
            rb.toggled.connect(lambda on, k=key: self._select_fmt(k) if on else None)
            sub = QLabel("      " + desc)
            sub.setFont(_font(8, "regular"))
            sub.setStyleSheet(f"color:{TXT_S};background:transparent;")
            row.addWidget(rb)
            row.addWidget(sub)
            vl.addLayout(row)
            self.fmt_group.addButton(rb)
            self._fmt_radios[key] = rb
        layout.addWidget(card)

        # ── 형식별 세부 설정 ──
        card, vl = _card("세부 설정")
        # QStackedWidget은 항상 '가장 큰 페이지' 높이로 카드를 고정해(EMF·SVG를
        # 골라도 PNG 페이지만큼 빈칸이 남음) 형식별 페이지를 형제 위젯으로 두고
        # show/hide 한다. 숨긴 위젯은 레이아웃 크기 계산에서 제외되므로 카드
        # 높이가 선택된 페이지에 딱 맞게 줄어든다.
        self._opt_pages: dict = {}

        emf_page = QWidget()
        emf_l = QVBoxLayout(emf_page)
        emf_l.setContentsMargins(0, 0, 0, 0)
        emf_note = QLabel(
            "도형·글자·이미지를 그대로 벡터로 다시 그립니다. 글꼴은 이 PC에 설치된 "
            "글꼴 중 가장 비슷한 것으로 바뀌며, 글자 위치는 원본 그대로 유지됩니다.\n"
            "그라데이션·반투명은 근사값으로 처리되고, 잘라내기(클리핑) 영역은 반영되지 않습니다.")
        emf_note.setWordWrap(True)
        emf_note.setFont(_font(8.5, "regular"))
        emf_note.setStyleSheet(f"color:{TXT_S};background:transparent;")
        emf_l.addWidget(emf_note)
        self._opt_pages["emf"] = emf_page
        vl.addWidget(emf_page)

        svg_page = QWidget()
        svg_l = QVBoxLayout(svg_page)
        svg_l.setContentsMargins(0, 0, 0, 0)
        self.path_chk = QCheckBox("글자를 도형(패스)으로 변환 — 글꼴이 없는 PC에서도 똑같이 보입니다")
        self.path_chk.setFont(_font(9, "regular"))
        self.path_chk.setChecked(True)
        svg_l.addWidget(self.path_chk)
        svg_note = QLabel("끄면 글자가 텍스트로 남아 편집·검색이 되지만, 같은 글꼴이 없으면 모양이 달라집니다.")
        svg_note.setWordWrap(True)
        svg_note.setFont(_font(8.5, "regular"))
        svg_note.setStyleSheet(f"color:{TXT_S};background:transparent;")
        svg_l.addWidget(svg_note)
        self._opt_pages["svg"] = svg_page
        vl.addWidget(svg_page)

        png_page = QWidget()
        png_l = QVBoxLayout(png_page)
        png_l.setContentsMargins(0, 0, 0, 0)
        png_l.setSpacing(6)
        dpi_lbl = QLabel("해상도(DPI)")
        dpi_lbl.setFont(_font(9, "medium"))
        dpi_lbl.setStyleSheet(f"color:{TXT};background:transparent;")
        png_l.addWidget(dpi_lbl)
        dpi_row = QHBoxLayout()
        for dpi in self.DPI_CHOICES:
            btn = QPushButton(str(dpi))
            btn.setFont(_font(9, "medium"))
            btn.setFixedHeight(32)
            btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            btn.clicked.connect(lambda _, d=dpi: self._select_dpi(d))
            dpi_row.addWidget(btn)
            self._dpi_btns[dpi] = btn
        png_l.addLayout(dpi_row)
        kind_lbl = QLabel("파일 형식")
        kind_lbl.setFont(_font(9, "medium"))
        kind_lbl.setStyleSheet(f"color:{TXT};background:transparent;")
        png_l.addWidget(kind_lbl)
        kind_row = QHBoxLayout()
        for key, text in (("png", "PNG — 무손실, 선 도면에 유리"),
                          ("jpg", "JPG — 사진·조감도 쪽 용량 대폭 감소")):
            btn = QPushButton(text)
            btn.setFont(_font(9, "medium"))
            btn.setFixedHeight(32)
            btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            btn.clicked.connect(lambda _, k=key: self._select_img_kind(k))
            kind_row.addWidget(btn)
            self._kind_btns[key] = btn
        png_l.addLayout(kind_row)
        self.alpha_chk = QCheckBox("배경을 투명하게 — 흰 바탕만 지우고 색·농도는 그대로")
        self.alpha_chk.setFont(_font(9, "regular"))
        png_l.addWidget(self.alpha_chk)
        png_note = QLabel(
            "투명 배경은 흰색 부분이 모두 비워지므로, 흰색으로 칠한 도형도 함께 투명해집니다.\n"
            "해상도가 너무 높아 메모리를 넘기면 자동으로 낮춰서 저장합니다(투명은 상한이 더 낮음).")
        png_note.setWordWrap(True)
        png_note.setFont(_font(8.5, "regular"))
        png_note.setStyleSheet(f"color:{TXT_S};background:transparent;")
        png_l.addWidget(png_note)
        self._opt_pages["png"] = png_page
        vl.addWidget(png_page)

        layout.addWidget(card)

        self._progress_card(layout)

        row = QHBoxLayout()
        row.setContentsMargins(6, 10, 6, 10)
        row.setSpacing(100)
        self.start_btn = QPushButton("변환 시작")
        self.stop_btn = QPushButton("변환 중지")
        for b in (self.start_btn, self.stop_btn):
            b.setFont(_font(11, "semibold"))
            b.setFixedHeight(32)
            b.setFixedWidth(140)
        self.start_btn.setEnabled(False)
        self.start_btn.setStyleSheet(BTN_OFF)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet(BTN_SOFT)
        self.start_btn.clicked.connect(self._start)
        self.stop_btn.clicked.connect(self._stop)
        row.addStretch()
        row.addWidget(self.start_btn)
        row.addWidget(self.stop_btn)
        row.addStretch()
        layout.addLayout(row)

        self._log_card(layout)
        layout.addStretch()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(_make_scroll(root))

        self._fmt_radios["emf"].setChecked(True)
        self._select_dpi(self._dpi)
        self._select_img_kind(self._img_kind)

    # ── 선택 상태 ──
    def _select_fmt(self, key: str) -> None:
        self._fmt = key
        # 선택된 형식의 페이지만 표시 → '세부 설정' 카드가 그 높이에 맞게 줄어듦
        for k, w in self._opt_pages.items():
            w.setVisible(k == key)
        if not self._fmt_radios[key].isChecked():
            self._fmt_radios[key].setChecked(True)

    def _select_dpi(self, dpi: int) -> None:
        self._dpi = dpi
        for d, btn in self._dpi_btns.items():
            btn.setStyleSheet(BTN_ON if d == dpi else BTN_SOFT)

    def _select_img_kind(self, kind: str) -> None:
        self._img_kind = kind
        for k, btn in self._kind_btns.items():
            btn.setStyleSheet(BTN_ON if k == kind else BTN_SOFT)
        # JPG는 투명 배경을 담을 수 없다.
        self.alpha_chk.setEnabled(kind == "png")
        if kind != "png":
            self.alpha_chk.setChecked(False)

    def _browse(self) -> None:
        path = self._browse_pdf()
        if path:
            self._load_file(path)

    def _load_file(self, path: str) -> None:
        self.file_edit.setText(path)
        try:
            with fitz.open(path) as doc:
                n = len(doc)
                r = doc[0].rect if n else None
            size_mb = os.path.getsize(path) / 1024 / 1024
            info = f"총 {n}페이지 | {size_mb:.2f} MB"
            if r is not None:
                info += f" | 첫 쪽 {r.width * 25.4 / 72:.0f} x {r.height * 25.4 / 72:.0f} mm"
            self.info_lbl.setText(info)
            self.status_lbl.setText("변환 시작 버튼을 눌러주세요.")
            self.log_box.clear()
            self.log_box.append(f"선택됨: {os.path.basename(path)}")
            self._set_action_ready(self.start_btn, True)
        except Exception as exc:
            QMessageBox.critical(self, "오류", str(exc))

    # ── 실행 ──
    def _start(self) -> None:
        path = self.file_edit.text()
        if not path:
            return
        fmt = self._img_kind if self._fmt == "png" else self._fmt
        opts = {"text_as_path": self.path_chk.isChecked(),
                "dpi": self._dpi,
                "transparent": self.alpha_chk.isChecked()}
        self.log_box.clear()
        self.progress.setValue(0)
        self.status_lbl.setText("변환 중")
        self._set_running_buttons(self.start_btn, True, self.stop_btn)
        self._worker = PdfExportWorker(path, fmt, opts)
        self._thread = QThread()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self.log_box.append)
        self._worker.progress.connect(self._on_progress)
        self._worker.done.connect(self._on_done)
        self._worker.error.connect(self._on_error)
        self._worker.cancelled.connect(self._on_cancelled)
        self._worker.done.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._worker.cancelled.connect(self._thread.quit)
        self._thread.finished.connect(self._on_finished)
        self._thread.start()

    def _stop(self) -> None:
        if self._worker:
            self._worker.request_stop()
        self.stop_btn.setEnabled(False)
        self.status_lbl.setText("중지 요청 중")

    def _on_progress(self, current: int, total: int, status: str) -> None:
        pct = int(current / total * 100) if total else 0
        self.progress.setValue(pct)
        self.count_lbl.setText(f"{current} / {total} ({pct}%)")
        self.status_lbl.setText(status)

    def _on_done(self, count: int, folder: str) -> None:
        self.progress.setValue(100)
        self.status_lbl.setText(f"변환 완료 - {count}개 파일 생성")
        kind = self._img_kind if self._fmt == "png" else self._fmt
        _recent_add("PDF 벡터 변환", f"{count}개 {kind.upper()} 생성",
                    os.path.basename(folder) or folder)
        QMessageBox.information(
            self, "완료",
            f"변환이 완료되었습니다.\n총 {count}개 파일 생성\n\n저장 위치:\n{folder}")

    def _on_error(self, msg: str) -> None:
        self.status_lbl.setText("오류 발생")
        self.log_box.append(f"[오류] {msg}")
        QMessageBox.critical(self, "오류", msg)
        self._offer_diag(msg)   # 동의 시 작업 로그를 GitHub 이슈로 전송

    def _on_cancelled(self) -> None:
        self.status_lbl.setText("변환 중지됨")

    def _on_finished(self) -> None:
        self._set_running_buttons(self.start_btn, False, self.stop_btn)
        self._set_action_ready(self.start_btn, bool(self.file_edit.text()))
        self._worker = None
        self._thread = None

    def is_running(self) -> bool:
        return bool(self._thread and self._thread.isRunning())

    def request_stop(self) -> None:
        if self._worker:
            self._worker.request_stop()


class PptTab(BaseTool):
    TOOL_LABEL = "PPT 이미지 변환"

    def __init__(self) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self._file_path = ""
        self._dpi = 300
        self._dpi_btns: dict[int, QPushButton] = {}
        self._worker: ConvertWorker | None = None
        self._thread: QThread | None = None
        self._build_ui()

    def reset_state(self) -> None:
        if self.is_running():
            return
        self._file_path = ""
        self.file_label.setText("")
        self.preview_widget.clearPixmap()
        self.drop_stack.setCurrentIndex(0)
        self.progress.setValue(0)
        self.count_lbl.clear()
        self.status_lbl.setText("파일을 선택해주세요.")
        self.log_box.clear()
        self._set_action_ready(self.start_btn, False)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet(BTN_SOFT)
        _clear_effect(self.stop_btn)

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("scroll_content")
        root.setAcceptDrops(True)
        root.installEventFilter(self)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(20, 16, 20, 10)
        layout.setSpacing(10)
        self._page_header(
            layout, "PPT 이미지 변환",
            "PPT 슬라이드를 고해상도 이미지로 바꿔, 어디서 열어도 깨지지 않는 자료를 만듭니다.")

        self.drop_card = QFrame()
        self.drop_card.setObjectName("drop_card")
        self.drop_card.setAcceptDrops(True)
        self.drop_card.installEventFilter(self)
        self.drop_card.setFixedHeight(240)
        card_l = QVBoxLayout(self.drop_card)
        card_l.setContentsMargins(0, 0, 0, 0)
        card_l.setSpacing(0)
        self.drop_stack = QStackedWidget()
        card_l.addWidget(self.drop_stack)

        empty_w = QWidget()
        empty_w.setStyleSheet("background: transparent;")
        empty_l = QVBoxLayout(empty_w)
        empty_l.setAlignment(Qt.AlignmentFlag.AlignCenter)
        empty_l.setSpacing(6)
        empty_l.setContentsMargins(20, 10, 20, 10)
        icon = QLabel("📂")
        icon.setFont(QFont("Segoe UI Emoji", 30))
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title = QLabel("여기에 PPTX 파일을 드래그")
        title.setFont(_font(10, "medium"))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet(f"color:{TXT};background:transparent;")
        sub = QLabel("또는 아래 버튼으로 파일을 선택하세요")
        sub.setFont(_font(9, "regular"))
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub.setStyleSheet(f"color:{TXT_S};background:transparent;")
        self.browse_btn = QPushButton("파일 선택하기")
        self.browse_btn.setFont(_font(9, "medium"))
        self.browse_btn.setFixedHeight(34)
        self.browse_btn.setStyleSheet(BTN_BROWSE)
        self.browse_btn.clicked.connect(self._browse)
        for w in (icon, title, sub, self.browse_btn):
            empty_l.addWidget(w, alignment=Qt.AlignmentFlag.AlignCenter)
        self.drop_stack.addWidget(empty_w)

        _preview_page = QWidget()
        _preview_page.setStyleSheet("background: transparent;")
        _preview_layout = QVBoxLayout(_preview_page)
        _preview_layout.setContentsMargins(10, 8, 10, 6)
        _preview_layout.setSpacing(4)
        self.preview_widget = RoundedPreviewWidget()
        self.file_label = ElideLabel("")
        self.file_label.setFont(_font(9, "medium"))
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_label.setStyleSheet(f"color:{TXT_S};background:transparent;")
        _preview_layout.addWidget(self.preview_widget, stretch=1)
        _preview_layout.addWidget(self.file_label)
        self.drop_stack.addWidget(_preview_page)
        self.drop_stack.setCurrentIndex(0)
        layout.addWidget(self.drop_card)

        card, vl = _card("내보내기 해상도(DPI)")
        row = QHBoxLayout()
        for dpi in (96, 150, 200, 300, 600, 900):
            btn = QPushButton(str(dpi))
            btn.setFont(_font(9, "medium"))
            btn.setFixedHeight(32)
            btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            btn.setStyleSheet(BTN_ON if dpi == self._dpi else BTN_SOFT)
            btn.clicked.connect(lambda _, d=dpi: self._select_dpi(d))
            row.addWidget(btn)
            self._dpi_btns[dpi] = btn
        vl.addLayout(row)
        layout.addWidget(card)

        card, vl = _card("진행 상황")
        top = QHBoxLayout()
        self.count_lbl = QLabel("")
        self.count_lbl.setFont(_font(9, "regular"))
        self.count_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        top.addStretch()
        top.addWidget(self.count_lbl)
        vl.addLayout(top)
        self.progress = QProgressBar()
        self.progress.setFixedHeight(6)
        self.progress.setTextVisible(False)
        vl.addWidget(self.progress)
        self.status_lbl = QLabel("파일을 선택해주세요.")
        self.status_lbl.setFont(_font(9, "regular"))
        self.status_lbl.setStyleSheet(f"color:{TXT_S};background:transparent;")
        vl.addWidget(self.status_lbl)
        layout.addWidget(card)

        row = QHBoxLayout()
        row.setContentsMargins(6, 10, 6, 10)
        row.setSpacing(100)
        self.start_btn = QPushButton("변환 시작")
        self.stop_btn = QPushButton("변환 중지")
        for b in (self.start_btn, self.stop_btn):
            b.setFont(_font(11, "semibold"))
            b.setFixedHeight(32)
            b.setFixedWidth(140)
        self.start_btn.setEnabled(False)
        self.start_btn.setStyleSheet(BTN_OFF)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet(BTN_SOFT)
        self.start_btn.clicked.connect(self._start)
        self.stop_btn.clicked.connect(self._stop)
        row.addStretch()
        row.addWidget(self.start_btn)
        row.addWidget(self.stop_btn)
        row.addStretch()
        layout.addLayout(row)

        card, vl = _card("작업 로그")
        self.log_box = LogoLogTextEdit()
        self.log_box.setObjectName("log_box")
        self.log_box.setReadOnly(True)
        self.log_box.setFont(_font(9, "light"))
        self.log_box.setFixedHeight(130)
        vl.addWidget(self.log_box)
        layout.addWidget(card)
        layout.addStretch()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        scroll = _make_scroll(root)
        scroll.setAcceptDrops(True)
        scroll.installEventFilter(self)
        outer.addWidget(scroll)

    def eventFilter(self, watched, event) -> bool:
        if event.type() == QEvent.Type.DragEnter:
            self.dragEnterEvent(event)
            return event.isAccepted()
        if event.type() == QEvent.Type.Drop:
            self.dropEvent(event)
            return True
        return super().eventFilter(watched, event)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls() and any(u.toLocalFile().lower().endswith(".pptx") for u in event.mimeData().urls()):
            # 드래그 진입 시 창을 앞으로 (다른 탭들과 동일 동작)
            self.window().raise_()
            self.window().activateWindow()
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".pptx"):
                self._set_file(path)
                break

    def _browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "PPTX 파일 선택", "", "PowerPoint 파일 (*.pptx);;모든 파일 (*.*)")
        if path:
            self._set_file(path)

    def _set_file(self, path: str) -> None:
        self._file_path = path
        self.file_label.setText(os.path.basename(path))
        self._update_preview(path)
        self._set_action_ready(self.start_btn, True)
        self.progress.setValue(0)
        self.count_lbl.setText("")
        self.status_lbl.setText("변환 시작 버튼을 눌러주세요.")
        self.log_box.clear()
        self.log_box.append(f"선택됨: {os.path.basename(path)}")

    def _update_preview(self, path: str) -> None:
        self.preview_widget.clearPixmap()
        self.drop_stack.setCurrentIndex(1)
        try:
            with zipfile.ZipFile(path, "r") as z:
                names = z.namelist()
                for cand in (
                    "docProps/thumbnail.jpeg",
                    "docProps/thumbnail.jpg",
                    "docProps/thumbnail.png",
                ):
                    if cand in names:
                        data = z.read(cand)
                        img = Image.open(io.BytesIO(data)).convert("RGB")
                        out_buf = io.BytesIO()
                        img.save(out_buf, format="PNG")
                        px = QPixmap()
                        px.loadFromData(out_buf.getvalue())
                        self.preview_widget.setPixmap(px)
                        return
        except Exception as exc:
            self.log_box.append(f"[미리보기] 썸네일을 읽지 못했습니다: {exc}")

    def _select_dpi(self, dpi: int) -> None:
        self._dpi = dpi
        for d, btn in self._dpi_btns.items():
            btn.setStyleSheet(BTN_ON if d == dpi else BTN_SOFT)

    def _start(self) -> None:
        if not self._file_path:
            return
        self._set_running_buttons(self.start_btn, True, self.stop_btn)
        self.progress.setValue(0)
        self.status_lbl.setText("변환 중")
        self._worker = ConvertWorker(self._file_path, self._dpi)
        self._thread = QThread()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self.log_box.append)
        self._worker.progress.connect(self._on_progress)
        self._worker.done.connect(self._on_done)
        self._worker.error.connect(self._on_error)
        self._worker.cancelled.connect(self._on_cancelled)
        self._worker.done.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._worker.cancelled.connect(self._thread.quit)
        self._thread.finished.connect(self._on_finished)
        self._thread.start()

    def _stop(self) -> None:
        if self._worker:
            self._worker.request_stop()
        self.stop_btn.setEnabled(False)
        self.status_lbl.setText("중지 요청 중")

    def _on_progress(self, current: int, total: int, status: str) -> None:
        pct = int(current / total * 100) if total else 0
        self.progress.setValue(pct)
        self.count_lbl.setText(f"{current} / {total} ({pct}%)")
        self.status_lbl.setText(status)

    def _on_done(self, out_path: str) -> None:
        self.progress.setValue(100)
        self.status_lbl.setText("변환 완료")
        _recent_add("PPT 이미지 변환", os.path.basename(out_path) or out_path, "")
        QMessageBox.information(self, "완료", f"변환이 완료되었습니다.\n\n저장 위치:\n{out_path}")

    def _on_error(self, msg: str) -> None:
        self.status_lbl.setText("오류 발생")
        QMessageBox.critical(self, "오류", msg)
        self._offer_diag(msg)   # 동의 시 작업 로그를 GitHub 이슈로 전송

    def _on_cancelled(self) -> None:
        self.status_lbl.setText("변환 중지됨")

    def _on_finished(self) -> None:
        self._set_running_buttons(self.start_btn, False, self.stop_btn)
        self._set_action_ready(self.start_btn, bool(self._file_path))
        self._worker = None
        self._thread = None

    def is_running(self) -> bool:
        return bool(self._thread and self._thread.isRunning())

    def request_stop(self) -> None:
        if self._worker:
            self._worker.request_stop()


# ─────────────────────────── Auto Update ─────────────────────────────────
class UpdateChecker(QThread):
    # 시작 시 백그라운드로 version.txt를 받아 현재 버전과 비교. 새 버전이면 시그널.
    update_available = Signal(str, str)   # 원격 버전, 변경사항 텍스트

    def run(self) -> None:
        try:
            r = _session.get(
                UPDATE_VERSION_URL,
                params={"ref": "main"},
                headers={
                    "Accept": "application/vnd.github.raw",
                    "Cache-Control": "no-cache",
                    "Pragma": "no-cache",
                },
                timeout=5,
            )
            if r.status_code != 200:
                return
            remote = (r.text or "").strip().split()[0] if r.text.strip() else ""
            if remote and is_newer_version(remote, APP_VERSION):
                # 변경사항도 같은 백그라운드 스레드에서 받아 UI를 막지 않는다.
                notes = changelog_notes(fetch_changelog_text(), newer_than=APP_VERSION)
                self.update_available.emit(remote, notes)
        except Exception as e:
            # 인터넷 없음/타임아웃/파싱 실패 → 실행은 방해하지 않고 기록만 남긴다
            log_err("UpdateChecker.run", e)


class ChangelogFetcher(QThread):
    # 업데이트 완료 팝업에 띄울 현재 버전 변경사항을 백그라운드로 받아온다.
    notes_ready = Signal(str)

    def run(self) -> None:
        try:
            self.notes_ready.emit(changelog_notes(fetch_changelog_text(), only=APP_VERSION))
        except Exception:
            self.notes_ready.emit("")   # 실패해도 팝업 자체는 띄운다


class UpdateDownloader(QThread):
    # 새 exe를 스트리밍 다운로드하며 진행률(0~100)을 전달하고, 다 받으면 검증한다.
    # 검증에 실패하면 받은 파일을 지우고 failed 를 보낸다(교체로 넘어가지 않는다).
    progress = Signal(int)
    finished_ok = Signal(str)
    failed = Signal(str)

    def __init__(self, dest_path: str) -> None:
        super().__init__()
        self.dest_path = dest_path

    def run(self) -> None:
        try:
            # 이전 시도가 남긴(숨김 속성 포함) 파일은 open("wb")이 실패하므로 먼저 제거.
            self._remove_partial_file()
            expected = _fetch_release_asset()
            hasher = hashlib.sha256()
            head = b""
            with _session.get(expected["url"], stream=True, timeout=30) as r:
                r.raise_for_status()
                total = int(r.headers.get("Content-Length", 0) or 0)
                done = 0
                with open(self.dest_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=256 * 1024):
                        if not chunk:
                            continue
                        if not head:
                            head = chunk[:len(EXE_SIGNATURE)]
                        f.write(chunk)
                        hasher.update(chunk)
                        done += len(chunk)
                        if total:
                            self.progress.emit(min(100, int(done * 100 / total)))
                if not total:
                    self.progress.emit(100)
            received = {"head": head, "size": done, "sha256": hasher.hexdigest()}
            problem = describe_download_problem(received, total, expected, _minimum_update_size())
            if problem:
                self._remove_partial_file()
                log_err("UpdateDownloader.verify", RuntimeError(problem), f"size={done} expected={expected}")
                self.failed.emit(problem)
                return
            self.finished_ok.emit(self.dest_path)
        except Exception as e:
            self._remove_partial_file()
            self.failed.emit(str(e))

    def _remove_partial_file(self) -> None:
        try:
            if os.path.exists(self.dest_path):
                os.remove(self.dest_path)
        except OSError:  # noqa: T2 — 못 지워도 다음 시도가 먼저 지우고, 교체는 이미 막혔다
            pass


class UpdateProgressDialog(QDialog):
    # 업데이트 다운로드 동안 띄우는 진행창. 닫기/Esc/X를 막는다.
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("업데이트")
        self._closable = False
        self.setWindowFlags(
            Qt.WindowType.Dialog
            | Qt.WindowType.CustomizeWindowHint
            | Qt.WindowType.WindowTitleHint
        )
        self.setFixedSize(380, 120)
        self.setStyleSheet(
            "QDialog { background-color: #FFFFFF; }"
            "QLabel { color: #1A2C3D; font-size: 12px; }"
            "QProgressBar { border: 1px solid #CFD8DC; border-radius: 4px; "
            "background-color: #ECEFF1; text-align: center; height: 20px; "
            "color: #1A237E; font-weight: bold; }"
            "QProgressBar::chunk { background-color: #1565C0; border-radius: 3px; }"
        )
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 18, 20, 18)
        lay.setSpacing(12)
        self.lbl = QLabel("새 버전을 내려받는 중입니다...")
        self.lbl.setWordWrap(True)
        self.bar = QProgressBar()
        self.bar.setRange(0, 100)
        self.bar.setValue(0)
        self.bar.setTextVisible(True)
        lay.addWidget(self.lbl)
        lay.addWidget(self.bar)

    def set_progress(self, pct: int) -> None:
        self.bar.setValue(pct)

    def set_status(self, text: str) -> None:
        self.lbl.setText(text)

    def allow_close(self) -> None:
        self._closable = True

    def closeEvent(self, event) -> None:
        if self._closable:
            event.accept()
        else:
            event.ignore()

    def keyPressEvent(self, event) -> None:
        if event.key() == Qt.Key.Key_Escape:
            return
        super().keyPressEvent(event)


class CaptionButton(QPushButton):
    # 창 제어 버튼(최소화/최대화/복원/닫기)을 폰트 글리프 없이 직접 그린다.
    # 어떤 폰트 환경에서도 또렷하고 동일하게 보인다.
    def __init__(self, kind: str, parent=None) -> None:
        super().__init__(parent)
        self.kind = kind
        self.setFixedSize(44, 30)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def set_kind(self, kind: str) -> None:
        self.kind = kind
        self.update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)   # QSS 배경/hover 먼저 그림
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        # 흰 타이틀바 → 어두운 글리프. 닫기 버튼은 hover(빨강 배경) 시 흰색.
        glyph = "#5B6172"
        if self.kind == "close" and self.underMouse():
            glyph = "#FFFFFF"
        pen = QPen(QColor(glyph))
        pen.setWidthF(1.3)
        p.setPen(pen)
        cx = self.width() / 2
        cy = self.height() / 2
        if self.kind == "min":
            p.drawLine(int(cx - 5), int(cy), int(cx + 5), int(cy))
        elif self.kind == "max":
            p.setRenderHint(QPainter.RenderHint.Antialiasing, False)
            p.drawRect(int(cx - 5), int(cy - 5), 10, 10)
        elif self.kind == "restore":
            p.setRenderHint(QPainter.RenderHint.Antialiasing, False)
            p.drawRect(int(cx - 6), int(cy - 3), 9, 9)
            p.drawLine(int(cx - 3), int(cy - 3), int(cx - 3), int(cy - 6))
            p.drawLine(int(cx - 3), int(cy - 6), int(cx + 6), int(cy - 6))
            p.drawLine(int(cx + 6), int(cy - 6), int(cx + 6), int(cy + 3))
            p.drawLine(int(cx + 6), int(cy + 3), int(cx + 3), int(cy + 3))
        elif self.kind == "close":
            p.drawLine(int(cx - 5), int(cy - 5), int(cx + 5), int(cy + 5))
            p.drawLine(int(cx - 5), int(cy + 5), int(cx + 5), int(cy - 5))
        p.end()


class RanetToolKit(QMainWindow):
    # 사이드바 내비 항목: (라벨, 아이콘 이름). index 0 = 홈 대시보드, 1~ = stack 페이지.
    NAV_ITEMS = [
        ("홈", "tab_home"),
        ("HWP → HWPX 변환", "tab_hwpx"),
        ("HWP → PDF 변환", "tab_hwp"),
        ("PDF 분할", "tab_split"),
        ("PDF 병합", "tab_merge"),
        ("PDF 편집", "tab_edit"),
        ("PDF 압축", "tab_compress"),
        ("PDF 벡터 변환", "tab_export"),
        ("PPT 이미지 변환", "tab_ppt"),
    ]
    # 홈 대시보드 카드 설명문 (NAV_ITEMS[1:] 순서)
    TOOL_DESCS = [
        "한글 문서를 hwpx로 일괄 교체",
        "한글 문서를 PDF로 일괄 변환",
        "페이지 범위·구간별로 나누기",
        "여러 PDF를 하나로, 폴더 일괄병합",
        "썸네일로 삭제·회전·끼워넣기",
        "용량 줄이기 · 3단계 프리셋",
        "SVG·EMF 벡터 · 초고해상도 PNG",
        "슬라이드를 고해상도 이미지로",
    ]

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"Ranet Toolkit v{APP_VERSION}")
        self.update_checker: UpdateChecker | None = None
        self.update_downloader: UpdateDownloader | None = None
        self.update_dialog: UpdateProgressDialog | None = None
        # 프레임리스 + 불투명 창. 반투명(WA_TranslucentBackground)은 ClearType이
        # 꺼져 글자가 뭉개지므로 쓰지 않는다. 둥근 모서리는 showEvent에서
        # Win11 DWM으로 깎는다(Win10은 직각 폴백). 그림자도 DWM이 그려준다.
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        # 축소되며 레이아웃 깨지는 것 방지: 최소 크기 = 기본 크기(확대만 허용)
        self.setMinimumSize(1120, 720)
        self.resize(1120, 720)
        self.setStyleSheet(QSS)
        self._nav_btns: list[QPushButton] = []
        self._nav_icons: list[tuple[QIcon, QIcon]] = []

        for name in ("RTK.ico", "RTK.png"):
            p = os.path.join(BASE_DIR, "assets", name)
            if os.path.exists(p):
                self.setWindowIcon(QIcon(p))
                break

        # ── 중앙: 좌 사이드바 + 우(타이틀바·페이지) ──
        central = QWidget()
        central.setObjectName("Central")
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        root.addWidget(self._build_sidebar())

        right = QWidget()
        right.setObjectName("Content")
        rv = QVBoxLayout(right)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(0)
        rv.addWidget(self._build_titlebar())

        # ── 페이지: 0=홈 대시보드, 1~6=도구 (NAV_ITEMS 순서와 동일) ──
        self.stack = QStackedWidget()
        self.hwpx_tab = HwpxTab()
        self.hwp_tab = HwpTab()
        self.split_tab = PdfSplitTab()
        self.merge_tab = PdfMergeTab()
        self.edit_tab = PdfEditTab()
        self.compress_tab = PdfCompressTab()
        self.export_tab = PdfExportTab()
        self.ppt_tab = PptTab()
        self.home_page = self._build_home()
        for w in (self.home_page, self.hwpx_tab, self.hwp_tab, self.split_tab,
                  self.merge_tab, self.edit_tab, self.compress_tab,
                  self.export_tab, self.ppt_tab):
            self.stack.addWidget(w)
        rv.addWidget(self.stack, 1)
        root.addWidget(right, 1)

        # 우하단 리사이즈 그립(프레임리스라 기본 리사이즈 없음). 떠 있는 오버레이로
        # 두고 resizeEvent에서 위치를 잡는다.
        self._grip = QSizeGrip(self)
        self._grip.setFixedSize(16, 16)
        self._grip.setStyleSheet("background: transparent;")

        self._select_page(0)

        # 모든 버튼에 손가락(포인팅) 커서. 비활성 버튼은 Qt가 자동으로 화살표로 보여줌.
        for btn in self.findChildren(QPushButton):
            btn.setCursor(Qt.CursorShape.PointingHandCursor)

        # 업데이트로 재실행된 경우(.bat가 --updated 인자로 실행) 완료 팝업.
        if "--updated" in sys.argv:
            QTimer.singleShot(500, self._notify_updated)
        self._check_update()   # 시작 시 백그라운드 업데이트 확인

    # ── 사이드바 ──
    def _build_sidebar(self) -> QFrame:
        side = QFrame()
        side.setObjectName("Sidebar")
        side.setFixedWidth(216)
        lay = QVBoxLayout(side)
        lay.setContentsMargins(14, 16, 14, 14)
        lay.setSpacing(4)

        # 로고 + 앱명 (드래그 핸들 겸용)
        brand = QHBoxLayout()
        brand.setSpacing(9)
        logo = QLabel()
        for name in ("RTK.png", "RTK.ico"):
            p = os.path.join(BASE_DIR, "assets", name)
            if os.path.exists(p):
                logo.setPixmap(QPixmap(p).scaled(
                    32, 32, Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation))
                break
        logo.setStyleSheet("background:transparent;")
        brand.addWidget(logo)
        name_col = QVBoxLayout()
        name_col.setSpacing(0)
        t1 = QLabel("Ranet Toolkit")
        t1.setObjectName("BrandName")
        t1.setFont(_font(11, "bold"))
        t2 = QLabel("업무 문서 도구")
        t2.setObjectName("BrandSub")
        t2.setFont(_font(7, "medium"))
        name_col.addWidget(t1)
        name_col.addWidget(t2)
        brand.addLayout(name_col)
        brand.addStretch()
        lay.addLayout(brand)
        lay.addSpacing(16)

        sec = QLabel("도구")
        sec.setObjectName("NavSection")
        sec.setFont(_font(8, "semibold"))
        lay.addWidget(sec)
        lay.addSpacing(2)

        for idx, (label, icon_name) in enumerate(self.NAV_ITEMS):
            btn = QPushButton("  " + label)
            btn.setFont(_font(9, "semibold"))
            btn.setFixedHeight(38)
            btn.setIconSize(QSize(17, 17))
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(lambda _, i=idx: self._select_page(i))
            lay.addWidget(btn)
            self._nav_btns.append(btn)
            self._nav_icons.append((
                QIcon(os.path.join(BASE_DIR, "assets", f"{icon_name}_on.svg")),
                QIcon(os.path.join(BASE_DIR, "assets", f"{icon_name}_off.svg")),
            ))
        lay.addStretch()

        # 하단 버전 칩: 평소 '최신 버전', 새 버전 발견 시 안내문으로 교체
        self.ver_chip = QLabel()
        self.ver_chip.setObjectName("VersionChip")
        self.ver_chip.setFont(_font(8, "semibold"))
        self._set_ver_chip_latest()
        lay.addWidget(self.ver_chip)

        # 라이선스 고지: PyMuPDF(AGPL-3.0)를 쓰므로 프로그램 전체를 AGPL-3.0 으로 배포하고
        # 화면에 라이선스·소스 위치를 알린다(AGPL 5조 «적절한 법적 고지»).
        license_lbl = QLabel(
            f'<a href="{SOURCE_URL}" style="color:{TXT_S};text-decoration:none;">'
            'AGPL-3.0 오픈소스 · 소스 보기</a>')
        license_lbl.setFont(_font(8, "regular"))
        license_lbl.setOpenExternalLinks(True)
        license_lbl.setStyleSheet("background:transparent;padding:2px 4px;")
        license_lbl.setToolTip("이 프로그램은 GNU AGPL-3.0 에 따라 배포됩니다. 보증 없음. 누르면 소스 저장소가 열립니다.")
        lay.addWidget(license_lbl)

        # 사이드바 빈 영역으로도 창 드래그 가능
        side.mousePressEvent = self._title_press
        side.mouseMoveEvent = self._title_move
        return side

    def _set_ver_chip_latest(self) -> None:
        self.ver_chip.setText(
            f'<span style="color:{OK};">●</span>&nbsp;&nbsp;'
            f'<span style="color:{TXT_S};">v{APP_VERSION} · 최신 버전</span>')

    # ── 홈 대시보드 ──
    def _build_home(self) -> QWidget:
        page = QWidget()
        page.setObjectName("scroll_content")
        lay = QVBoxLayout(page)
        lay.setContentsMargins(28, 22, 28, 18)
        lay.setSpacing(8)

        head = QLabel("무엇을 도와드릴까요?")
        head.setFont(_font(16, "bold"))
        lay.addWidget(head)
        sub = QLabel("자주 쓰는 문서 작업을 골라 바로 시작하세요.")
        sub.setFont(_font(9, "regular"))
        sub.setStyleSheet(f"color:{TXT_S};background:transparent;")
        lay.addWidget(sub)
        lay.addSpacing(10)

        grid = QGridLayout()
        grid.setSpacing(12)
        for i, (label, icon_name) in enumerate(self.NAV_ITEMS[1:]):
            card = self._tool_card(i + 1, icon_name, label, self.TOOL_DESCS[i])
            grid.addWidget(card, i // 3, i % 3)
        lay.addLayout(grid)
        lay.addSpacing(14)

        recent_lbl = QLabel("최근 작업")
        recent_lbl.setFont(_font(10, "bold"))
        lay.addWidget(recent_lbl)
        self._recent_holder = QVBoxLayout()
        self._recent_holder.setSpacing(6)
        lay.addLayout(self._recent_holder)
        lay.addStretch()

        outer_page = QWidget()
        ov = QVBoxLayout(outer_page)
        ov.setContentsMargins(0, 0, 0, 0)
        ov.addWidget(_make_scroll(page))
        return outer_page

    def _tool_card(self, page_idx: int, icon_name: str, title: str, desc: str) -> QFrame:
        card = QFrame()
        card.setFixedHeight(112)
        card.setCursor(Qt.CursorShape.PointingHandCursor)
        card.setStyleSheet(
            f"QFrame {{ background:{SURFACE}; border:1px solid {BORDER}; border-radius:14px; }}"
            f"QFrame:hover {{ border:1px solid {ACCENT}; }}")
        v = QVBoxLayout(card)
        v.setContentsMargins(16, 13, 16, 11)
        v.setSpacing(5)
        ic = QLabel()
        ic.setPixmap(_tinted_pixmap(icon_name, ACCENT, 22))
        ic.setFixedSize(36, 36)
        ic.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ic.setStyleSheet(f"background:{ACCENT_SOFT};border:none;border-radius:10px;")
        v.addWidget(ic)
        t = QLabel(title)
        t.setFont(_font(9.5, "bold"))
        t.setStyleSheet(f"color:{TXT};border:none;background:transparent;")
        d = QLabel(desc)
        d.setFont(_font(8, "regular"))
        d.setStyleSheet(f"color:{TXT_S};border:none;background:transparent;")
        v.addWidget(t)
        v.addWidget(d)
        # 카드 클릭 → 해당 도구 페이지로 이동
        card.mouseReleaseEvent = lambda e, i=page_idx: self._select_page(i)
        return card

    def _refresh_recent(self) -> None:
        # 홈 진입 시마다 최근 작업 목록을 다시 그린다.
        while self._recent_holder.count():
            item = self._recent_holder.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        items = _recent_load()[:6]
        if not items:
            lbl = QLabel("아직 기록이 없습니다. 작업을 완료하면 여기에 표시됩니다.")
            lbl.setFont(_font(8.5, "regular"))
            lbl.setStyleSheet(f"color:{TXT_T};background:transparent;")
            self._recent_holder.addWidget(lbl)
            return
        for idx, it in enumerate(items):
            row = QFrame()
            row.setFixedHeight(42)
            row.setStyleSheet(
                f"background:{SURFACE}; border:1px solid {BORDER}; border-radius:11px;")
            h = QHBoxLayout(row)
            h.setContentsMargins(14, 0, 8, 0)
            h.setSpacing(10)
            tool = QLabel(str(it.get("tool", "")))
            tool.setFont(_font(8, "semibold"))
            tool.setFixedWidth(100)
            tool.setAlignment(Qt.AlignmentFlag.AlignCenter)
            tool.setStyleSheet(
                f"color:{ACCENT}; background:{ACCENT_SOFT}; border:none;"
                f"border-radius:9px; padding:3px 0;")
            h.addWidget(tool)
            title = QLabel(str(it.get("title", "")))
            title.setFont(_font(8.5, "medium"))
            title.setStyleSheet(f"color:{TXT};border:none;background:transparent;")
            h.addWidget(title)
            h.addStretch()
            meta_txt = str(it.get("meta", ""))
            when = str(it.get("when", ""))
            meta = QLabel(f"{meta_txt}   ·   {when}" if meta_txt else when)
            meta.setFont(_font(8, "regular"))
            meta.setStyleSheet(f"color:{TXT_T};border:none;background:transparent;")
            h.addWidget(meta)
            # 행 삭제 버튼 (× — 라틴 곱셈 기호라 어떤 폰트에서도 안전)
            close = QPushButton("×")
            close.setFixedSize(24, 24)
            close.setFont(_font(11, "medium"))
            close.setCursor(Qt.CursorShape.PointingHandCursor)
            close.setToolTip("기록 삭제")
            close.setStyleSheet(
                f"QPushButton {{ background:transparent; color:{TXT_T}; border:none;"
                f" border-radius:7px; }}"
                f"QPushButton:hover {{ background:#F1F2F9; color:{DANGER}; }}")
            close.clicked.connect(lambda _, i=idx: self._remove_recent(i))
            h.addWidget(close)
            self._recent_holder.addWidget(row)

    def _remove_recent(self, idx: int) -> None:
        # 최근 작업 기록 1건 삭제 후 목록 다시 그림.
        items = _recent_load()
        if 0 <= idx < len(items):
            del items[idx]
            _recent_save(items)
        self._refresh_recent()

    # ── 커스텀 타이틀바 (콘텐츠 상단, 흰 배경 + 브레드크럼) ──
    def _build_titlebar(self) -> QFrame:
        bar = QFrame()
        bar.setObjectName("TitleBar")
        bar.setFixedHeight(44)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(20, 0, 8, 0)
        lay.setSpacing(8)
        self.crumb = QLabel("홈")
        self.crumb.setObjectName("Crumb")
        self.crumb.setFont(_font(9, "semibold"))
        lay.addWidget(self.crumb)
        lay.addStretch()
        self.min_btn = CaptionButton("min")
        self.min_btn.setObjectName("WinMin")
        self.min_btn.clicked.connect(self.showMinimized)
        self.max_btn = CaptionButton("max")
        self.max_btn.setObjectName("WinMax")
        self.max_btn.clicked.connect(self._toggle_max)
        self.close_btn = CaptionButton("close")
        self.close_btn.setObjectName("WinClose")
        self.close_btn.clicked.connect(self.close)
        lay.addWidget(self.min_btn)
        lay.addWidget(self.max_btn)
        lay.addWidget(self.close_btn)
        bar.mousePressEvent = self._title_press
        bar.mouseMoveEvent = self._title_move
        bar.mouseDoubleClickEvent = lambda e: self._toggle_max()
        return bar

    def _toggle_max(self) -> None:
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()

    def showEvent(self, e) -> None:
        super().showEvent(e)
        # Windows 11 네이티브 둥근 모서리 + 그림자.
        # (불투명 창이라 QSS radius 대신 DWM이 모서리를 깎는다. Win10은 직각 폴백)
        try:
            hwnd = int(self.winId())
            pref = ctypes.c_int(2)   # DWMWCP_ROUND
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 33, ctypes.byref(pref), ctypes.sizeof(pref))   # 33 = CORNER_PREFERENCE
        except Exception:
            pass

    def resizeEvent(self, e) -> None:
        super().resizeEvent(e)
        # 리사이즈 그립을 항상 우하단 모서리에 띄워 둔다.
        if hasattr(self, "_grip"):
            self._grip.move(self.width() - 18, self.height() - 18)
            self._grip.raise_()

    def changeEvent(self, e) -> None:
        # 최대화/복원 시 최대화 버튼 모양만 갈아끼운다(여백·라운드는 DWM 몫).
        if e.type() == QEvent.Type.WindowStateChange and hasattr(self, "max_btn"):
            self.max_btn.set_kind("restore" if self.isMaximized() else "max")
        super().changeEvent(e)

    def _title_press(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton:
            self._drag_pos = e.globalPosition().toPoint() - self.frameGeometry().topLeft()
            e.accept()

    def _title_move(self, e) -> None:
        if self.isMaximized():
            return
        if (e.buttons() & Qt.MouseButton.LeftButton) and hasattr(self, "_drag_pos"):
            self.move(e.globalPosition().toPoint() - self._drag_pos)
            e.accept()

    # ── 자동업데이트 ──
    def _notify_updated(self) -> None:
        # 변경사항을 백그라운드로 받은 뒤 팝업 (네트워크 실패 시 변경사항 없이 표시).
        self._changelog_fetcher = ChangelogFetcher()
        self._changelog_fetcher.notes_ready.connect(self._show_updated_popup)
        self._changelog_fetcher.start()

    def _show_updated_popup(self, notes: str) -> None:
        msg = f"최신 버전 v{APP_VERSION} 으로 업데이트되었습니다."
        if notes:
            msg += f"\n\n[변경사항]\n{notes}"
        QMessageBox.information(self, "업데이트 완료", msg)

    def _check_update(self) -> None:
        # exe(frozen)로 실행될 때만 확인. 스크립트 개발 모드에선 건너뛴다.
        if not getattr(sys, "frozen", False):
            return
        self.update_checker = UpdateChecker()
        self.update_checker.update_available.connect(self._on_update_available)
        self.update_checker.start()

    def _on_update_available(self, remote_version: str, notes: str = "") -> None:
        # 사이드바 버전 칩도 새 버전 안내로 교체
        self.ver_chip.setText(
            f'<span style="color:{ACCENT};">●</span>&nbsp;&nbsp;'
            f'<span style="color:{ACCENT};">새 버전 v{remote_version} 있음</span>')
        # 변경사항은 여기서 보여주지 않는다(긴 목록으로 창이 커짐) —
        # 업데이트 완료 후 팝업에서만 표시. (란맵과 동일 UX)
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Icon.Question)
        msg.setWindowTitle("업데이트 확인")
        msg.setText(f"새 버전 v{remote_version} 이(가) 있습니다. (현재 v{APP_VERSION})")
        msg.setInformativeText(
            "지금 업데이트할까요?\n예를 누르면 새 버전을 내려받아 자동으로 교체·재실행합니다.")
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg.setDefaultButton(QMessageBox.StandardButton.Yes)
        _korean_buttons(msg)
        if msg.exec() != QMessageBox.StandardButton.Yes:
            return
        self._start_download()

    def _start_download(self) -> None:
        # 새 파일은 사용자 눈에 안 띄도록 임시폴더에 받는다. (숨김 속성 주지 않음)
        dest = os.path.join(tempfile.gettempdir(), "ranet_toolkit_new.exe")
        self.hide()
        self.update_dialog = UpdateProgressDialog(None)
        self.update_dialog.show()
        scr = QApplication.primaryScreen().availableGeometry()
        self.update_dialog.move(
            scr.center().x() - self.update_dialog.width() // 2,
            scr.center().y() - self.update_dialog.height() // 2,
        )
        self.update_downloader = UpdateDownloader(dest)
        self.update_downloader.progress.connect(self._on_dl_progress)
        self.update_downloader.finished_ok.connect(self._on_dl_done)
        self.update_downloader.failed.connect(self._on_dl_failed)
        self.update_downloader.start()

    def _on_dl_progress(self, pct: int) -> None:
        if self.update_dialog:
            self.update_dialog.set_progress(pct)

    def _on_dl_failed(self, err: str) -> None:
        if self.update_dialog:
            self.update_dialog.allow_close()
            self.update_dialog.close()
            self.update_dialog = None
        self.show()
        QMessageBox.warning(self, "업데이트 실패", f"업데이트 다운로드에 실패했습니다.\n기존 버전으로 계속 사용합니다.\n\n{err}")

    def _on_dl_done(self, new_path: str) -> None:
        if self.update_dialog:
            self.update_dialog.set_progress(100)
            self.update_dialog.set_status("업데이트 적용 중... 잠시 후 자동 재시작됩니다")
        self._apply_update(new_path)

    def _apply_update(self, new_path: str) -> None:
        # 실행 중 exe는 자기 자신을 덮어쓸 수 없으므로, 종료를 기다렸다 교체·재실행하는
        # .bat를 만들고 창을 완전히 숨기는 VBScript로 감싸 백그라운드 실행한다.
        # PyInstaller onefile 환경변수(_MEIPASS2 등)를 비워야 'Failed to load Python DLL'을 막는다.
        tmp = tempfile.gettempdir()
        cur_exe = sys.executable
        bat_path = os.path.join(tmp, "ranet_toolkit_update.bat")
        vbs_path = os.path.join(tmp, "ranet_toolkit_update.vbs")
        # move 실패(파일 잠김) 시 재시도하되, 쓰기권한 없는 폴더(Program Files 등)에서
        # 영구 실패하면 무한 루프가 되므로 약 40초(40회)에서 포기하고 구버전을 다시 띄운다.
        script = (
            "@echo off\r\n"
            "ping -n 3 127.0.0.1 >nul\r\n"
            "set /a tries=0\r\n"
            ":retry\r\n"
            f'move /y "{new_path}" "{cur_exe}" >nul 2>&1\r\n'
            "if not errorlevel 1 goto ok\r\n"
            "set /a tries+=1\r\n"
            "if %tries% GEQ 40 goto giveup\r\n"
            "ping -n 2 127.0.0.1 >nul\r\n"
            "goto retry\r\n"
            ":ok\r\n"
            "set _MEIPASS2=\r\n"
            "set _PYI_ONEFILE_PARENT_PID=\r\n"
            "set _PYI_ARCHIVE_FILE=\r\n"
            "set _PYI_APPLICATION_HOME_DIR=\r\n"
            f'start "" "{cur_exe}" --updated\r\n'
            "goto cleanup\r\n"
            ":giveup\r\n"
            f'start "" "{cur_exe}"\r\n'   # 교체 실패 → 기존 버전 그대로 재실행
            ":cleanup\r\n"
            f'del "{vbs_path}" >nul 2>&1\r\n'
            'del "%~f0"\r\n'
        )
        vbs = f'CreateObject("WScript.Shell").Run "cmd /c ""{bat_path}""", 0, False\r\n'
        try:
            with open(bat_path, "w", encoding="cp949") as f:
                f.write(script)
            with open(vbs_path, "w", encoding="cp949") as f:
                f.write(vbs)
            child_env = os.environ.copy()
            for k in ("_MEIPASS2", "_PYI_ONEFILE_PARENT_PID", "_PYI_ARCHIVE_FILE",
                      "_PYI_APPLICATION_HOME_DIR"):
                child_env.pop(k, None)
            subprocess.Popen(["wscript.exe", vbs_path],
                             creationflags=subprocess.CREATE_NO_WINDOW,
                             close_fds=True, env=child_env)
        except Exception:
            if self.update_dialog:
                self.update_dialog.allow_close()
                self.update_dialog.close()
                self.update_dialog = None
            self.show()
            QMessageBox.warning(self, "업데이트 실패", "업데이트 적용에 실패했습니다.\n기존 버전으로 계속 사용합니다.")
            return
        if self.update_dialog:
            self.update_dialog.allow_close()
        QApplication.quit()

    def _select_page(self, index: int) -> None:
        # 사이드바/홈 카드에서 페이지 전환. 0=홈, 1~=도구.
        previous = self.stack.currentIndex() if hasattr(self, "stack") else index
        if previous != index and previous >= 0:
            previous_widget = self.stack.widget(previous)
            if hasattr(previous_widget, "is_running") and previous_widget.is_running():
                QMessageBox.warning(self, "작업 진행 중", "현재 도구의 작업이 진행 중입니다.\n작업을 중지하거나 완료한 뒤 이동해주세요.")
                return
            if hasattr(previous_widget, "reset_state"):
                previous_widget.reset_state()
        self.stack.setCurrentIndex(index)
        label = self.NAV_ITEMS[index][0]
        self.crumb.setText("홈" if index == 0 else f"도구  ›  {label}")
        for i, btn in enumerate(self._nav_btns):
            active = i == index
            btn.setStyleSheet(NAV_ON if active else NAV_OFF)
            on_icon, off_icon = self._nav_icons[i]
            btn.setIcon(on_icon if active else off_icon)
        if index == 0:
            self._refresh_recent()   # 홈 진입 시 최근 작업 갱신

    def closeEvent(self, event) -> None:
        running_tabs = [self.hwpx_tab, self.hwp_tab, self.split_tab, self.merge_tab,
                        self.edit_tab, self.compress_tab, self.ppt_tab]
        if any(tab.is_running() for tab in running_tabs):
            reply = QMessageBox.question(
                self,
                "종료 확인",
                "작업이 진행 중입니다.\n종료하시겠습니까?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
            for tab in running_tabs:
                if hasattr(tab, "request_stop"):
                    tab.request_stop()
            # 진행 중 워커 스레드가 끝나길 잠시 기다린다. 안 기다리면 종료 시
            # 'QThread destroyed while running' 경고/비정상 종료 + Office·한글 COM
            # 작업 중이면 고아 프로세스(WINWORD/POWERPNT/Hwp)가 남을 수 있다.
            for tab in running_tabs:
                th = getattr(tab, "_thread", None)
                if th is not None:
                    try:
                        if th.isRunning():
                            th.quit()
                            th.wait(5000)   # 파일 사이 중지 신호 반영까지 최대 5초
                    except RuntimeError:
                        pass   # 이미 파괴된 QThread
        # 일괄병합 자식 프로세스가 아직 살아 있으면 즉시 종료(고아 프로세스 방지).
        w = getattr(self.merge_tab, "_worker", None)
        if isinstance(w, _BatchProcWorker):
            w.hard_kill()
        # 진단 전송 스레드도 잠시 대기(실행 중 파괴 크래시 방지).
        for tab in running_tabs:
            ds = getattr(tab, "_diag_sender", None)
            if ds is not None and ds.isRunning():
                ds.wait(3000)
        # 업데이트 스레드도 끝까지 대기해 실행 중 파괴 크래시를 막는다.
        # 무제한 wait() 는 네트워크가 느리면 창이 «응답 없음» 인 채로 멈춘다.
        # 상한을 두고 못 끝내면 넘어간다(둘 다 데몬 성격의 작업).
        if self.update_checker and self.update_checker.isRunning():
            self.update_checker.wait(3000)
        if self.update_downloader and self.update_downloader.isRunning():
            self.update_downloader.wait(3000)
        event.accept()


# 단일 실행 보장용 전역 뮤텍스 핸들(프로세스 수명 동안 유지되어야 함).
_SINGLE_INSTANCE_HANDLE = None
_ERROR_ALREADY_EXISTS = 183


def _acquire_single_instance() -> bool:
    # 같은 이름의 named mutex를 만든다. 이미 있으면(다른 인스턴스 실행 중) False.
    # 업데이트 다운로드 도중 사용자가 exe를 다시 클릭해도 2번째 인스턴스가 떠
    # 같은 임시파일에 동시 접근·교체하는 충돌을 막는다.
    global _SINGLE_INSTANCE_HANDLE
    try:
        h = ctypes.windll.kernel32.CreateMutexW(None, False, "ranet_toolkit_single_instance_mutex")
        if ctypes.windll.kernel32.GetLastError() == _ERROR_ALREADY_EXISTS:
            return False
        _SINGLE_INSTANCE_HANDLE = h
        return True
    except Exception:
        return True


def _start_parent_watchdog() -> None:
    """부트로더(부모)가 사라지면 실앱(자식)도 스스로 끝낸다.

    PyInstaller onefile 은 «부트로더(부모) + 실앱(자식)» 두 프로세스로 뜬다.
    작업관리자에서 부모만 끝내면 자식이 살아남아 단일실행 뮤텍스를 계속 쥐고,
    다음 실행이 «이미 실행 중» 으로 막힌다(실측 2026-09-07: 부모 강제종료 3초 뒤
    자식 생존·뮤텍스 유지). 한글 COM 호출로 자식이 굳어 있으면 더 오래 못 켠다.
    스크립트로 실행할 때는 부모가 셸이라 감시하지 않는다.
    """
    if not (getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS")):
        return
    try:
        ppid = os.getppid()
    except OSError:
        return
    if not ppid:
        return

    def _watch() -> None:
        SYNCHRONIZE = 0x00100000
        INFINITE = 0xFFFFFFFF
        h = ctypes.windll.kernel32.OpenProcess(SYNCHRONIZE, False, ppid)
        if not h:
            return
        ctypes.windll.kernel32.WaitForSingleObject(h, INFINITE)
        ctypes.windll.kernel32.CloseHandle(h)
        # 부모가 죽었다 = 정상 종료 경로가 아니다. Qt 를 건드리면 오히려 멈추므로
        # 정리 없이 즉시 끝낸다.
        os._exit(0)

    threading.Thread(target=_watch, daemon=True).start()


def _cleanup_stale_temp() -> None:
    # 강제종료 등 비정상 종료로 남은 일괄병합 임시파일(작업명세 JSON + stop 신호)을
    # 시작 시 정리한다. 단일 실행 뮤텍스 획득 후에만 호출되므로(다른 인스턴스 없음)
    # 사용 중 파일을 지울 위험이 없다.
    import glob
    tmp = tempfile.gettempdir()
    for pat in ("ranet_batchjob_*.json", "ranet_batchjob_*.json.stop"):
        for p in glob.glob(os.path.join(tmp, pat)):
            try:
                os.remove(p)
            except OSError:
                pass


def _splash_text(msg: str) -> None:
    # 시작 스플래시(빌드 시 spec의 Splash) 하단 문구 갱신. 스플래시 없는 환경(스크립트 실행 등)은 무시.
    try:
        import pyi_splash
        pyi_splash.update_text(msg)
    except Exception:
        pass


def _splash_close() -> None:
    try:
        import pyi_splash
        pyi_splash.close()
    except Exception:
        pass


def main() -> int:
    # 임베드 폰트(Pretendard)를 작은 한글에서도 또렷하게: FreeType 엔진 + 풀힌팅.
    # 기본 DirectWrite는 힌팅 없는 폰트를 뭉갠다. QApplication 생성 전에 설정해야 한다.
    os.environ.setdefault("QT_QPA_PLATFORM", "windows:fontengine=freetype")
    _splash_text("프로그램 기동 중...")
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    # 중복 실행 차단: 이미 실행 중이면 안내 후 종료.
    if not _acquire_single_instance():
        _splash_close()
        QMessageBox.information(
            None, "이미 실행 중",
            "프로그램이 이미 실행 중입니다.\n업데이트 중일 수 있으니 잠시 기다려 주세요.\n\n"
            "창이 안 보이는데 계속 이 안내가 뜨면, 작업관리자에서 "
            "'Ranet Toolkit.exe' 를 모두 끝낸 뒤 다시 실행해 주세요.",
        )
        return 0
    _start_parent_watchdog()   # 부트로더만 강제종료됐을 때 자식이 남지 않게
    _cleanup_stale_temp()   # 비정상 종료가 남긴 임시파일 정리
    _splash_text("폰트 불러오는 중...")
    _load_fonts()
    app.setPalette(_make_light_palette())
    _splash_text("화면 구성 중...")
    win = RanetToolKit()
    win.show()
    _splash_close()   # 메인 창이 떴으니 스플래시 종료
    return app.exec()


if __name__ == "__main__":
    # 일괄병합 변환은 별도 프로세스에서 돈다(GUI가 한글/오피스 COM 때문에 멈추지 않게).
    # GUI가 자기 자신(exe/스크립트)을 '--batch-worker <작업파일>'로 재실행한다.
    if len(sys.argv) >= 3 and sys.argv[1] == "--batch-worker":
        _splash_close()   # 자식 배치워커는 GUI 없음 → 부트로더 스플래시 즉시 닫기
        raise SystemExit(_run_batch_worker_cli(sys.argv[2]))
    raise SystemExit(main())
