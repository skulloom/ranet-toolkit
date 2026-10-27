# Ranet Toolkit - Copyright (C) 2026 Region & Agriculture NETwork Cooperative, Jeonnam
# SPDX-License-Identifier: AGPL-3.0-only  (see LICENSE)
"""PDF 한 쪽을 EMF(벡터)로 직접 기록.

GDI를 ctypes로 직접 불러 도형·글자·이미지를 다시 그린다(외부 변환기 없음).
`Ranet ToolKit.py`에서 분리했다(동작 변경 없음).
"""

import ctypes
from ctypes import wintypes
import math
import struct

import fitz
from PIL import Image

from rtk_common import log_err


# EMF는 외부 프로그램 없이 GDI로 직접 기록한다(벡터 유지). 핵심 제약 두 가지:
#  ① EMF 재생기는 그린 좌표(장치 단위)를 헤더의 szlDevice/szlMillimeters 비율로
#     물리치수로 환산한 뒤 rclFrame에 맞춰 배치한다.
#  ② 기록 시점에 기준 장치(reference device)의 픽셀 범위를 넘는 좌표는 잘려 나간다.
# 그래서 고해상도 프린터 IC를 기준 장치로 잡고, 페이지가 그 범위에 들어가도록 배율을
# 정한 뒤, 저장 후 헤더의 szlDevice/szlMillimeters를 그 배율에 맞게 고쳐 쓴다.
_EMF_GDI = ctypes.WinDLL("gdi32", use_last_error=True)
_EMF_USER = ctypes.WinDLL("user32", use_last_error=True)
_EMF_P = ctypes.c_void_p

_EMF_ALTERNATE, _EMF_WINDING = 1, 2
_EMF_TRANSPARENT = 1
_EMF_TA_LEFT_BASELINE = 0 | 24 | 0          # TA_LEFT | TA_BASELINE | TA_NOUPDATECP
_EMF_PS_GEOMETRIC, _EMF_PS_SOLID = 0x00010000, 0
_EMF_CAPS = {0: 0x00000200, 1: 0x00000000, 2: 0x00000100}   # flat / round / square
_EMF_JOINS = {0: 0x00002000, 1: 0x00000000, 2: 0x00001000}  # miter / round / bevel
_EMF_BS_SOLID = 0
_EMF_DIB_RGB_COLORS, _EMF_SRCCOPY = 0, 0x00CC0020
_EMF_FW_NORMAL, _EMF_FW_BOLD = 400, 700
_EMF_DEFAULT_CHARSET, _EMF_OUT_TT_PRECIS, _EMF_ANTIALIASED = 1, 4, 4
_EMF_HALFTONE = 4
_EMF_NULL_BRUSH, _EMF_NULL_PEN = 5, 8
_EMF_HORZSIZE, _EMF_VERTSIZE, _EMF_HORZRES, _EMF_VERTRES = 4, 6, 8, 10
# 목표 배율(1pt당 논리단위). 기준 장치 범위에 맞춰 이보다 낮아질 수 있다.
_EMF_TARGET_UPP = 20.0
# 기준 장치로 우선 시도할 프린터(윈도우 기본 설치분).
_EMF_PREF_PRINTERS = ("Microsoft Print to PDF", "Microsoft XPS Document Writer")
_EMF_REF_CACHE: dict = {}


class _EMF_RECT(ctypes.Structure):
    _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG),
                ("right", wintypes.LONG), ("bottom", wintypes.LONG)]


class _EMF_LOGBRUSH(ctypes.Structure):
    _fields_ = [("lbStyle", wintypes.UINT), ("lbColor", wintypes.DWORD),
                ("lbHatch", ctypes.c_void_p)]


class _EMF_LOGFONTW(ctypes.Structure):
    _fields_ = [("lfHeight", wintypes.LONG), ("lfWidth", wintypes.LONG),
                ("lfEscapement", wintypes.LONG), ("lfOrientation", wintypes.LONG),
                ("lfWeight", wintypes.LONG), ("lfItalic", wintypes.BYTE),
                ("lfUnderline", wintypes.BYTE), ("lfStrikeOut", wintypes.BYTE),
                ("lfCharSet", wintypes.BYTE), ("lfOutPrecision", wintypes.BYTE),
                ("lfClipPrecision", wintypes.BYTE), ("lfQuality", wintypes.BYTE),
                ("lfPitchAndFamily", wintypes.BYTE), ("lfFaceName", wintypes.WCHAR * 32)]


class _EMF_BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
                ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
                ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
                ("biClrImportant", wintypes.DWORD)]


def _emf_bind() -> None:
    # ctypes 기본 int 변환은 64비트 핸들에서 넘치므로 argtypes를 명시한다.
    sig = {
        "CreateEnhMetaFileW": ([_EMF_P, wintypes.LPCWSTR, _EMF_P, wintypes.LPCWSTR], _EMF_P),
        "CloseEnhMetaFile": ([_EMF_P], _EMF_P),
        "DeleteEnhMetaFile": ([_EMF_P], wintypes.BOOL),
        "CreateICW": ([wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, _EMF_P], _EMF_P),
        "DeleteDC": ([_EMF_P], wintypes.BOOL),
        "GetDeviceCaps": ([_EMF_P, ctypes.c_int], ctypes.c_int),
        "SetPolyFillMode": ([_EMF_P, ctypes.c_int], ctypes.c_int),
        "SetBkMode": ([_EMF_P, ctypes.c_int], ctypes.c_int),
        "SetTextAlign": ([_EMF_P, wintypes.UINT], wintypes.UINT),
        "SetTextColor": ([_EMF_P, wintypes.DWORD], wintypes.DWORD),
        "SetStretchBltMode": ([_EMF_P, ctypes.c_int], ctypes.c_int),
        "CreateSolidBrush": ([wintypes.DWORD], _EMF_P),
        "CreatePen": ([ctypes.c_int, ctypes.c_int, wintypes.DWORD], _EMF_P),
        "ExtCreatePen": ([wintypes.DWORD, wintypes.DWORD, _EMF_P, wintypes.DWORD, _EMF_P], _EMF_P),
        "CreateFontIndirectW": ([_EMF_P], _EMF_P),
        "SelectObject": ([_EMF_P, _EMF_P], _EMF_P),
        "DeleteObject": ([_EMF_P], wintypes.BOOL),
        "GetStockObject": ([ctypes.c_int], _EMF_P),
        "BeginPath": ([_EMF_P], wintypes.BOOL),
        "EndPath": ([_EMF_P], wintypes.BOOL),
        "CloseFigure": ([_EMF_P], wintypes.BOOL),
        "FillPath": ([_EMF_P], wintypes.BOOL),
        "StrokePath": ([_EMF_P], wintypes.BOOL),
        "StrokeAndFillPath": ([_EMF_P], wintypes.BOOL),
        "MoveToEx": ([_EMF_P, ctypes.c_int, ctypes.c_int, _EMF_P], wintypes.BOOL),
        "LineTo": ([_EMF_P, ctypes.c_int, ctypes.c_int], wintypes.BOOL),
        "PolyBezierTo": ([_EMF_P, _EMF_P, wintypes.DWORD], wintypes.BOOL),
        "Polygon": ([_EMF_P, _EMF_P, ctypes.c_int], wintypes.BOOL),
        "Rectangle": ([_EMF_P, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int], wintypes.BOOL),
        "ExtTextOutW": ([_EMF_P, ctypes.c_int, ctypes.c_int, wintypes.UINT, _EMF_P,
                         wintypes.LPCWSTR, wintypes.UINT, _EMF_P], wintypes.BOOL),
        "StretchDIBits": ([_EMF_P, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                           ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                           _EMF_P, _EMF_P, wintypes.UINT, wintypes.DWORD], ctypes.c_int),
    }
    for name, (argtypes, restype) in sig.items():
        fn = getattr(_EMF_GDI, name)
        fn.argtypes = argtypes
        fn.restype = restype
    _EMF_USER.FillRect.argtypes = [_EMF_P, _EMF_P, _EMF_P]
    _EMF_USER.GetDC.argtypes = [_EMF_P]
    _EMF_USER.GetDC.restype = _EMF_P
    _EMF_USER.ReleaseDC.argtypes = [_EMF_P, _EMF_P]


_emf_bind()

# PDF 폰트명 → 윈도우 글꼴. 부분집합 접두어(ABCDEF+)와 스타일 접미어는 미리 떼어낸다.
_EMF_FONT_MAP = {
    "hcrbatang": "함초롬바탕", "hcrdotum": "함초롬돋움", "hcrbatangext": "함초롬바탕",
    "hcrdotumext": "함초롬돋움", "hcrbatangextb": "함초롬바탕",
    "hymyeongjo": "HY견명조", "hygothic": "HY견고딕", "hyheadline": "HY헤드라인M",
    "hyshortsamul": "HY얕은샘물M", "hygsrb": "HY견명조", "hyrgotic": "HY중고딕",
    "batang": "바탕", "batangche": "바탕체", "gulim": "굴림", "gulimche": "굴림체",
    "dotum": "돋움", "dotumche": "돋움체", "gungsuh": "궁서", "gungsuhche": "궁서체",
    "malgungothic": "맑은 고딕", "malgun": "맑은 고딕",
    "nanumgothic": "나눔고딕", "nanummyeongjo": "나눔명조", "nanumbarungothic": "나눔바른고딕",
    "nanumsquare": "나눔스퀘어", "notosanskr": "본고딕", "notoserifkr": "본명조",
    "times": "Times New Roman", "timesnewroman": "Times New Roman", "timesroman": "Times New Roman",
    "arial": "Arial", "helvetica": "Arial", "courier": "Courier New",
    "couriernew": "Courier New", "symbol": "Symbol", "wingdings": "Wingdings",
    "calibri": "Calibri", "cambria": "Cambria", "verdana": "Verdana", "tahoma": "Tahoma",
    "romans": "Arial", "romand": "Arial", "txt": "Arial",     # AutoCAD 기본 글꼴
}


def _emf_printer_names() -> list:
    # 기본 프린터 → 선호 목록 → 나머지 순으로 후보를 만든다.
    names = []
    try:
        import win32print
        try:
            names.append(win32print.GetDefaultPrinter())
        except Exception as e:
            log_err("_emf_printer_names.default", e)
        names += [p[2] for p in win32print.EnumPrinters(2)]
    except Exception as e:
        log_err("_emf_printer_names.enum", e)
    ordered = [n for n in _EMF_PREF_PRINTERS if n in names]
    return ordered + [n for n in names if n not in ordered]


def _emf_ref_device() -> tuple:
    """(hdc, 픽셀폭, 픽셀높이) — 픽셀 범위가 가장 넓은 기준 장치. 프린터가 없으면 화면."""
    if "dc" in _EMF_REF_CACHE:
        return _EMF_REF_CACHE["dc"]
    best = None
    for name in _emf_printer_names():
        try:
            ic = _EMF_GDI.CreateICW("WINSPOOL", name, None, None)
        except Exception:
            ic = None
        if not ic:
            continue
        w = _EMF_GDI.GetDeviceCaps(ic, _EMF_HORZRES)
        h = _EMF_GDI.GetDeviceCaps(ic, _EMF_VERTRES)
        if w > 0 and h > 0 and (best is None or w * h > best[1] * best[2]):
            if best:
                _EMF_GDI.DeleteDC(best[0])
            best = (ic, w, h)
        else:
            _EMF_GDI.DeleteDC(ic)
    if best is None:
        scr = _EMF_USER.GetDC(None)
        best = (scr, _EMF_GDI.GetDeviceCaps(scr, _EMF_HORZRES) or 1920,
                _EMF_GDI.GetDeviceCaps(scr, _EMF_VERTRES) or 1080)
    _EMF_REF_CACHE["dc"] = best
    return best


def _emf_color(color, opacity: float = 1.0) -> int:
    # PDF 색(0~1 실수 3개) → GDI COLORREF(0x00BBGGRR).
    # GDI에는 알파가 없어 투명도는 흰 배경과 섞어 근사한다.
    if not color:
        return 0
    vals = [max(0.0, min(1.0, float(c))) for c in color[:3]]
    if len(vals) < 3:
        vals = (vals + vals + vals)[:3]
    r, g, b = vals
    if opacity is not None and opacity < 1.0:
        a = max(0.0, min(1.0, opacity))
        r, g, b = (v * a + (1 - a) for v in (r, g, b))
    return (int(b * 255) << 16) | (int(g * 255) << 8) | int(r * 255)


def _emf_face(name: str) -> tuple:
    # PDF 폰트명에서 부분집합 접두어·스타일 표기를 떼고 (윈도우 글꼴명, 굵게, 기울임) 반환.
    raw = name or ""
    if len(raw) > 7 and raw[6] == "+":
        raw = raw[7:]
    low = raw.lower()
    bold = any(k in low for k in ("bold", "heavy", "black", "semib", "extrab"))
    italic = any(k in low for k in ("italic", "oblique"))
    key = low
    for junk in ("-boldmt", "-bolditalicmt", "-italicmt", "-psmt", "mt", "ps",
                 "bolditalic", "boldoblique", "italic", "oblique", "bold",
                 "regular", "medium", "light", "semibold", "extrabold", "heavy", "black"):
        key = key.replace(junk, "")
    key = "".join(ch for ch in key if ch.isalnum())
    if key in _EMF_FONT_MAP:
        return _EMF_FONT_MAP[key], bold, italic
    face = raw.split("+")[-1]
    for sep in ("-", ","):
        if sep in face:
            face = face.split(sep)[0]
    return (face or "맑은 고딕")[:31], bold, italic


class EmfCanvas:
    """EMF 파일 하나를 열고 GDI 레코드를 쌓는 래퍼. with 블록으로 쓴다."""

    def __init__(self, path: str, width_pt: float, height_pt: float, mat=None) -> None:
        self.path = path
        # PyMuPDF의 도형·글자 좌표는 '회전 전' 좌표계다. /Rotate 가 걸린 페이지는
        # 여기서 회전 행렬을 곱해 화면에 보이는 대로 그린다.
        self.mat = mat
        href, ref_w, ref_h = _emf_ref_device()
        # 기준 장치 픽셀 범위를 벗어난 좌표는 기록 단계에서 잘린다 → 들어갈 만큼만 축소.
        self.scale = max(0.5, min(_EMF_TARGET_UPP,
                                  ref_w / max(1.0, width_pt), ref_h / max(1.0, height_pt)))
        self.width_pt, self.height_pt = width_pt, height_pt
        # rclFrame(0.01mm)이 삽입 시 물리 크기 = PDF 페이지 크기를 결정한다.
        frame = _EMF_RECT(0, 0,
                          max(1, int(round(width_pt * 2540.0 / 72.0))),
                          max(1, int(round(height_pt * 2540.0 / 72.0))))
        desc = ctypes.create_unicode_buffer("Ranet Toolkit\x00PDF\x00")
        self.hdc = _EMF_GDI.CreateEnhMetaFileW(href, path, ctypes.byref(frame),
                                               ctypes.cast(desc, wintypes.LPCWSTR))
        if not self.hdc:
            raise OSError("EMF 파일을 만들지 못했습니다.")
        _EMF_GDI.SetBkMode(self.hdc, _EMF_TRANSPARENT)
        _EMF_GDI.SetTextAlign(self.hdc, _EMF_TA_LEFT_BASELINE)
        _EMF_GDI.SetStretchBltMode(self.hdc, _EMF_HALFTONE)
        # GDI 객체는 프로세스당 개수 제한(기본 1만)이 있다. 도형 수만 개짜리 도면에서
        # 매번 새로 만들면 곧바로 고갈되므로 속성별로 캐시해 재사용한다.
        self._cache: dict = {}
        self._objs: list = []

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()
        return False

    def close(self) -> None:
        if self.hdc:
            hemf = _EMF_GDI.CloseEnhMetaFile(self.hdc)
            self.hdc = None
            if hemf:
                _EMF_GDI.DeleteEnhMetaFile(hemf)
            self._patch_header()
        for obj in self._objs:
            _EMF_GDI.DeleteObject(obj)
        self._objs.clear()
        self._cache.clear()

    def _patch_header(self) -> None:
        # 헤더의 szlDevice/szlMillimeters를 실제 사용한 배율에 맞춰 고쳐 쓴다.
        # (기준 장치 해상도가 아니라 우리 배율이 물리치수 환산 기준이 되도록)
        try:
            with open(self.path, "r+b") as f:
                head = bytearray(f.read(88))
                if len(head) < 88 or struct.unpack_from("<I", head, 0)[0] != 1:
                    return
                mm = 1000                                        # 가상 기준 장치 1000mm
                dev = int(round(mm * self.scale * 72.0 / 25.4))
                struct.pack_into("<ii", head, 72, dev, dev)       # szlDevice
                struct.pack_into("<ii", head, 80, mm, mm)         # szlMillimeters
                f.seek(0)
                f.write(head)
        except Exception as e:
            log_err("EmfCanvas._patch_header", e, self.path)

    # ── 내부 헬퍼 ──
    def _x(self, v: float) -> int:
        return int(round(v * self.scale))

    def _p(self, x: float, y: float) -> tuple:
        # 회전 행렬 적용 후 논리 단위로. fitz.Matrix(a,b,c,d,e,f) → x'=ax+cy+e, y'=bx+dy+f
        m = self.mat
        if m is not None:
            x, y = m.a * x + m.c * y + m.e, m.b * x + m.d * y + m.f
        return int(round(x * self.scale)), int(round(y * self.scale))

    def _rect(self, r):
        # 90도 배수 회전이라 사각형은 사각형으로 유지된다.
        if self.mat is None:
            return r
        return (r * self.mat).normalize()

    def _p2(self, x: float, y: float) -> tuple:
        # 이미 회전이 적용된 좌표(_rect 결과)를 논리 단위로만 바꾼다.
        return int(round(x * self.scale)), int(round(y * self.scale))

    def _dir(self, dx: float, dy: float) -> tuple:
        m = self.mat
        if m is None:
            return dx, dy
        return m.a * dx + m.c * dy, m.b * dx + m.d * dy

    def _cached(self, key, factory):
        obj = self._cache.get(key)
        if obj is None:
            obj = factory()
            if not obj:
                return None
            if len(self._cache) < 4000:      # 만일에 대비한 상한
                self._cache[key] = obj
            self._objs.append(obj)
        return obj

    def _use_brush(self, color: int) -> None:
        br = self._cached(("b", color), lambda: _EMF_GDI.CreateSolidBrush(color))
        if br:
            _EMF_GDI.SelectObject(self.hdc, br)

    def _use_null_brush(self) -> None:
        _EMF_GDI.SelectObject(self.hdc, _EMF_GDI.GetStockObject(_EMF_NULL_BRUSH))

    def _use_null_pen(self) -> None:
        _EMF_GDI.SelectObject(self.hdc, _EMF_GDI.GetStockObject(_EMF_NULL_PEN))

    def _use_pen(self, color: int, width_pt: float, cap, join) -> None:
        cap = cap[0] if isinstance(cap, (list, tuple)) and cap else cap
        join = join[0] if isinstance(join, (list, tuple)) and join else join
        style = (_EMF_PS_GEOMETRIC | _EMF_PS_SOLID
                 | _EMF_CAPS.get(cap, _EMF_CAPS[0]) | _EMF_JOINS.get(join, _EMF_JOINS[0]))
        w = max(1, self._x(width_pt or 0.5))

        def make():
            lb = _EMF_LOGBRUSH(_EMF_BS_SOLID, color, None)
            return (_EMF_GDI.ExtCreatePen(style, w, ctypes.byref(lb), 0, None)
                    or _EMF_GDI.CreatePen(_EMF_PS_SOLID, w, color))

        pen = self._cached(("p", color, w, style), make)
        if pen:
            _EMF_GDI.SelectObject(self.hdc, pen)

    # ── 그리기 ──
    def fill_page(self, color: int = 0xFFFFFF) -> None:
        # 페이지 전체를 칠해 EMF 경계를 페이지 크기에 고정한다(여백이 잘리지 않게).
        rc = _EMF_RECT(0, 0, self._x(self.width_pt), self._x(self.height_pt))
        br = self._cached(("b", color), lambda: _EMF_GDI.CreateSolidBrush(color))
        _EMF_USER.FillRect(self.hdc, ctypes.byref(rc), br)

    def draw_path(self, path: dict) -> None:
        ptype = path.get("type", "f")
        do_fill, do_stroke = "f" in ptype, "s" in ptype
        if not (do_fill or do_stroke):
            return
        if do_fill:
            self._use_brush(_emf_color(path.get("fill"), path.get("fill_opacity", 1.0)))
        else:
            self._use_null_brush()
        if do_stroke:
            self._use_pen(_emf_color(path.get("color"), path.get("stroke_opacity", 1.0)),
                          path.get("width") or 0.5, path.get("lineCap"), path.get("lineJoin"))
        else:
            self._use_null_pen()
        _EMF_GDI.SetPolyFillMode(self.hdc, _EMF_ALTERNATE if path.get("even_odd") else _EMF_WINDING)

        _EMF_GDI.BeginPath(self.hdc)
        cur = None
        for item in path.get("items", []):
            op = item[0]
            if op == "l":
                p1, p2 = item[1], item[2]
                if cur is None or abs(cur.x - p1.x) > 1e-4 or abs(cur.y - p1.y) > 1e-4:
                    _EMF_GDI.MoveToEx(self.hdc, *self._p(p1.x, p1.y), None)
                _EMF_GDI.LineTo(self.hdc, *self._p(p2.x, p2.y))
                cur = p2
            elif op == "c":
                p1, p2, p3, p4 = item[1], item[2], item[3], item[4]
                if cur is None or abs(cur.x - p1.x) > 1e-4 or abs(cur.y - p1.y) > 1e-4:
                    _EMF_GDI.MoveToEx(self.hdc, *self._p(p1.x, p1.y), None)
                pts = (wintypes.POINT * 3)(
                    wintypes.POINT(*self._p(p2.x, p2.y)),
                    wintypes.POINT(*self._p(p3.x, p3.y)),
                    wintypes.POINT(*self._p(p4.x, p4.y)))
                _EMF_GDI.PolyBezierTo(self.hdc, ctypes.byref(pts), 3)
                cur = p4
            elif op == "re":
                r = self._rect(item[1])
                _EMF_GDI.Rectangle(self.hdc, *self._p2(r.x0, r.y0), *self._p2(r.x1, r.y1))
                cur = None
            elif op == "qu":
                q = item[1]
                quad = (wintypes.POINT * 4)(
                    wintypes.POINT(*self._p(q.ul.x, q.ul.y)),
                    wintypes.POINT(*self._p(q.ur.x, q.ur.y)),
                    wintypes.POINT(*self._p(q.lr.x, q.lr.y)),
                    wintypes.POINT(*self._p(q.ll.x, q.ll.y)))
                _EMF_GDI.Polygon(self.hdc, ctypes.byref(quad), 4)
                cur = None
        if path.get("closePath"):
            _EMF_GDI.CloseFigure(self.hdc)
        _EMF_GDI.EndPath(self.hdc)
        if do_fill and do_stroke:
            _EMF_GDI.StrokeAndFillPath(self.hdc)
        elif do_fill:
            _EMF_GDI.FillPath(self.hdc)
        else:
            _EMF_GDI.StrokePath(self.hdc)

    def draw_span(self, span: dict, angle_deg: float = 0.0) -> None:
        # rawdict의 글자별 원점을 dx 배열로 넘겨, 대체 글꼴이 쓰여도 글자 위치가 어긋나지 않게 한다.
        chars = span.get("chars") or []
        text = "".join(c["c"] for c in chars) if chars else span.get("text", "")
        if not text.strip():
            return
        face, bold, italic = _emf_face(span.get("font", ""))
        size = span.get("size", 10.0) or 10.0
        flags = span.get("flags", 0)
        lf = _EMF_LOGFONTW()
        lf.lfHeight = -max(1, self._x(size))
        lf.lfWeight = _EMF_FW_BOLD if (bold or (flags & 16)) else _EMF_FW_NORMAL
        lf.lfItalic = 1 if (italic or (flags & 2)) else 0
        lf.lfCharSet = _EMF_DEFAULT_CHARSET
        lf.lfOutPrecision = _EMF_OUT_TT_PRECIS
        lf.lfQuality = _EMF_ANTIALIASED
        lf.lfEscapement = lf.lfOrientation = int(round(angle_deg * 10)) % 3600
        lf.lfFaceName = face[:31]
        key = ("f", lf.lfHeight, lf.lfWeight, lf.lfItalic, lf.lfEscapement, face)
        fnt = self._cached(key, lambda: _EMF_GDI.CreateFontIndirectW(ctypes.byref(lf)))
        if fnt:
            _EMF_GDI.SelectObject(self.hdc, fnt)
        v = span.get("color", 0)
        _EMF_GDI.SetTextColor(self.hdc, _emf_color(
            (((v >> 16) & 255) / 255.0, ((v >> 8) & 255) / 255.0, (v & 255) / 255.0)))

        if chars:
            ox, oy = chars[0]["origin"]
            dx = (ctypes.c_int * len(text))()
            i = 0
            for n, ch in enumerate(chars):
                w = (chars[n + 1]["origin"][0] - ch["origin"][0]) if n + 1 < len(chars) \
                    else (ch["bbox"][2] - ch["origin"][0])
                # 서로게이트 페어는 UTF-16 코드 단위 2개 → dx도 2칸을 차지한다.
                units = max(1, len(ch["c"].encode("utf-16-le")) // 2)
                if i + units > len(text):
                    break
                dx[i] = self._x(max(0.0, w))
                for k in range(1, units):
                    dx[i + k] = 0
                i += units
            _EMF_GDI.ExtTextOutW(self.hdc, *self._p(ox, oy), 0, None,
                                 text, len(text), ctypes.byref(dx))
        else:
            bbox = span.get("bbox", (0, 0, 0, 0))
            ox, oy = span.get("origin", (bbox[0], bbox[3]))
            _EMF_GDI.ExtTextOutW(self.hdc, *self._p(ox, oy), 0, None,
                                 text, len(text), None)

    def draw_image(self, pil_img, rect) -> None:
        # GDI 패스에 알파가 없어 투명 이미지는 흰 배경에 합성해 넣는다.
        if pil_img.mode in ("RGBA", "LA", "P"):
            rgba = pil_img.convert("RGBA")
            bg = Image.new("RGB", rgba.size, (255, 255, 255))
            bg.paste(rgba, mask=rgba.split()[3])
            pil_img = bg
        elif pil_img.mode != "RGB":
            pil_img = pil_img.convert("RGB")
        w, h = pil_img.size
        if w < 1 or h < 1:
            return
        # DIB 픽셀은 BGR 순서 + 행마다 4바이트 정렬.
        row = (w * 3 + 3) & ~3
        buf = bytearray(pil_img.tobytes("raw", "BGR", row, 0))
        bmi = _EMF_BITMAPINFOHEADER()
        bmi.biSize = ctypes.sizeof(_EMF_BITMAPINFOHEADER)
        bmi.biWidth = w
        bmi.biHeight = -h              # 음수 = top-down
        bmi.biPlanes = 1
        bmi.biBitCount = 24
        bmi.biCompression = 0
        cbuf = (ctypes.c_char * len(buf)).from_buffer(buf)
        rect = self._rect(rect)
        _EMF_GDI.StretchDIBits(self.hdc,
                               *self._p2(rect.x0, rect.y0),
                               max(1, self._x(rect.width)), max(1, self._x(rect.height)),
                               0, 0, w, h,
                               ctypes.cast(cbuf, _EMF_P), ctypes.byref(bmi),
                               _EMF_DIB_RGB_COLORS, _EMF_SRCCOPY)


def pdf_page_to_emf(page, out_path: str) -> dict:
    """PDF 페이지 하나를 벡터 EMF로 저장하고 처리 결과 수치를 돌려준다."""
    rect = page.rect
    stat = {"paths": 0, "spans": 0, "images": 0, "skipped": 0}
    mat = page.rotation_matrix if page.rotation else None
    with EmfCanvas(out_path, rect.width, rect.height, mat) as cv:
        cv.fill_page()

        # ① 이미지(맨 아래) — 배치 사각형에 맞춰 늘려 넣는다.
        try:
            infos = page.get_image_info(xrefs=True)
        except Exception:
            infos = []
        for info in infos:
            bbox = fitz.Rect(info.get("bbox") or (0, 0, 0, 0))
            xref = info.get("xref", 0)
            if bbox.is_empty or bbox.is_infinite or not xref:
                stat["skipped"] += 1
                continue
            try:
                pix = fitz.Pixmap(page.parent, xref)
                if pix.colorspace is None or pix.colorspace.n == 4:     # CMYK 등 → RGB
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                img = Image.frombytes("RGBA" if pix.alpha else "RGB",
                                      (pix.width, pix.height), pix.samples)
                cv.draw_image(img, bbox - rect.tl)
                stat["images"] += 1
            except Exception:
                stat["skipped"] += 1

        # ② 도형 — 콘텐츠 스트림 순서(seqno)대로
        for path in sorted(page.get_drawings(), key=lambda d: d.get("seqno", 0)):
            try:
                cv.draw_path(path)
                stat["paths"] += 1
            except Exception:
                stat["skipped"] += 1

        # ③ 글자(맨 위)
        for block in page.get_text("rawdict")["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                d = cv._dir(*line.get("dir", (1.0, 0.0)))
                angle = -math.degrees(math.atan2(d[1], d[0]))
                for span in line.get("spans", []):
                    try:
                        cv.draw_span(span, angle)
                        stat["spans"] += 1
                    except Exception:
                        stat["skipped"] += 1
    return stat
