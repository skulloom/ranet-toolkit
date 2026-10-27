# Ranet Toolkit - Copyright (C) 2026 Region & Agriculture NETwork Cooperative, Jeonnam
# SPDX-License-Identifier: AGPL-3.0-only  (see LICENSE)
"""
HWP/HWPX -> PDF 일괄 변환기

- 기본: 한글 내장 PDF 저장
- 선택: Microsoft Print to PDF를 pyautogui/pyperclip으로 제어

기존 pyautogui 좌표 방식은 해상도/배율/타이밍에 취약했기 때문에,
인쇄 창과 저장 창은 UI Automation 컨트롤 이름/값 기반으로 조작한다.
"""

from __future__ import annotations

import ctypes
import os
import shutil
import sys
import tempfile
import threading
import time
# tkinter는 이 모듈을 단독 GUI 앱으로 실행할 때만 필요하다. Ranet Toolkit에서
# import해 정적 메서드만 쓸 때는 tkinter가 없어도(미설치 환경) 로드돼야 하므로
# import 실패를 무시한다. GUI 클래스 본문은 from __future__ annotations 덕에
# tk 타입힌트가 문자열로 남아 import 시점엔 tk를 평가하지 않는다.
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk
except Exception:
    tk = None
    filedialog = messagebox = scrolledtext = ttk = None

import pythoncom
import win32api
import win32com.client
import win32gui
import win32process


BM_CLICK = 0x00F5
TARGET_PRINTER = "Microsoft Print to PDF"
class HwpToPdfApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("HWP_PDF 변환기 v5.5")
        self.root.geometry("820x820")
        self.root.minsize(820, 760)
        self.root.resizable(True, True)

        self._file_vars: dict[str, tk.BooleanVar] = {}
        self._converting = False
        self._stop_event = threading.Event()
        self.engine_var = tk.StringVar(value="hwp")

        self._build_ui()

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 4}

        src_frame = tk.LabelFrame(self.root, text="변환할 HWP 파일이 있는 폴더", **pad)
        src_frame.pack(fill="x", **pad)

        self.src_var = tk.StringVar()
        tk.Entry(src_frame, textvariable=self.src_var, width=58).pack(side="left", padx=5, pady=5)
        tk.Button(src_frame, text="폴더 선택", command=self._select_src, width=10).pack(side="left", padx=5)

        opt_frame = tk.Frame(self.root)
        opt_frame.pack(fill="x", padx=10, pady=2)

        self.subfolder_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            opt_frame,
            text="하위 폴더까지 검색",
            variable=self.subfolder_var,
            command=self._refresh_file_list,
        ).pack(side="left")

        self.hwpx_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            opt_frame,
            text=".hwpx 파일도 포함",
            variable=self.hwpx_var,
            command=self._refresh_file_list,
        ).pack(side="left", padx=15)

        list_frame = tk.LabelFrame(self.root, text="변환할 파일 목록", **pad)
        list_frame.pack(fill="x", **pad)

        top_row = tk.Frame(list_frame)
        top_row.pack(fill="x", padx=5, pady=2)

        self.select_all_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            top_row,
            text="전체 선택/해제",
            variable=self.select_all_var,
            command=self._toggle_select_all,
        ).pack(side="left")

        self.file_count_label = tk.Label(
            top_row,
            text="(폴더를 선택하면 파일 목록이 표시됩니다)",
            fg="gray",
        )
        self.file_count_label.pack(side="left", padx=10)

        list_container = tk.Frame(list_frame)
        list_container.pack(fill="x", padx=5, pady=(0, 5))

        self.file_canvas = tk.Canvas(
            list_container,
            height=130,
            bg="white",
            highlightthickness=1,
            highlightbackground="#cccccc",
        )
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.file_canvas.yview)
        self.file_canvas.configure(yscrollcommand=scrollbar.set)
        self.file_canvas.pack(side="left", fill="x", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.file_list_frame = tk.Frame(self.file_canvas, bg="white")
        self._canvas_win = self.file_canvas.create_window((0, 0), window=self.file_list_frame, anchor="nw")

        self.file_list_frame.bind(
            "<Configure>",
            lambda _e: self.file_canvas.configure(scrollregion=self.file_canvas.bbox("all")),
        )
        self.file_canvas.bind(
            "<Configure>",
            lambda e: self.file_canvas.itemconfig(self._canvas_win, width=e.width),
        )

        dst_frame = tk.LabelFrame(self.root, text="PDF를 저장할 폴더", **pad)
        dst_frame.pack(fill="x", **pad)

        self.dst_var = tk.StringVar()
        self.dst_entry = tk.Entry(dst_frame, textvariable=self.dst_var, width=58)
        self.dst_entry.pack(side="left", padx=5, pady=5)
        self.dst_btn = tk.Button(dst_frame, text="폴더 선택", command=self._select_dst, width=10)
        self.dst_btn.pack(side="left", padx=5)

        self.same_dir_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            dst_frame,
            text="변환할 폴더와 동일",
            variable=self.same_dir_var,
            command=self._toggle_same_dir,
        ).pack(side="left", padx=8)

        engine_frame = tk.LabelFrame(self.root, text="PDF 변환 방식", padx=10, pady=4)
        engine_frame.pack(fill="x", padx=10, pady=4)

        tk.Radiobutton(
            engine_frame,
            text="한글 내장 PDF 엔진 (빠름)",
            variable=self.engine_var,
            value="hwp",
            command=self._update_engine_label,
        ).pack(anchor="w")
        tk.Radiobutton(
            engine_frame,
            text="Microsoft Print to PDF (느리지만 깨지지 않음)",
            variable=self.engine_var,
            value="mspdf",
            command=self._update_engine_label,
        ).pack(anchor="w")

        self.engine_info_label = tk.Label(
            engine_frame,
            text="현재: 한글 내장 PDF 엔진",
            fg="#2962FF",
            font=("맑은 고딕", 9, "bold"),
        )
        self.engine_info_label.pack(anchor="w", pady=(2, 0))

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=6)

        self.convert_btn = tk.Button(
            btn_frame,
            text="변환 시작",
            command=self._start_conversion,
            bg="#2962FF",
            fg="white",
            font=("맑은 고딕", 11, "bold"),
            height=2,
            width=12,
        )
        self.convert_btn.pack(side="left", padx=8)

        self.stop_btn = tk.Button(
            btn_frame,
            text="변환 중지",
            command=self._stop_conversion,
            bg="#D32F2F",
            fg="white",
            font=("맑은 고딕", 11, "bold"),
            height=2,
            width=12,
            state="disabled",
        )
        self.stop_btn.pack(side="left", padx=8)

        prog_frame = tk.Frame(self.root)
        prog_frame.pack(fill="x", padx=10)

        self.status_label = tk.Label(prog_frame, text="대기 중", anchor="w")
        self.status_label.pack(fill="x")

        self.progress = ttk.Progressbar(prog_frame, mode="determinate")
        self.progress.pack(fill="x", pady=3)

        log_frame = tk.LabelFrame(self.root, text="변환 로그", **pad)
        log_frame.pack(fill="both", expand=True, **pad)

        self.log_box = scrolledtext.ScrolledText(log_frame, height=8, state="disabled", font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=True, padx=5, pady=5)

    def _update_engine_label(self) -> None:
        if self.engine_var.get() == "hwp":
            self.engine_info_label.config(text="현재: 한글 내장 PDF 엔진")
        else:
            self.engine_info_label.config(text=f"현재: {TARGET_PRINTER} (느리지만 깨지지 않음)")

    @staticmethod
    def _to_win_path(path: str) -> str:
        p = os.path.abspath(path).replace("/", "\\")
        parent = os.path.dirname(p)
        if os.path.exists(p):
            return os.path.realpath(p).replace("/", "\\")
        if os.path.exists(parent):
            return os.path.join(os.path.realpath(parent), os.path.basename(p)).replace("/", "\\")
        return p

    def _select_src(self) -> None:
        folder = filedialog.askdirectory(title="HWP 파일이 있는 폴더를 선택하세요")
        if folder:
            self.src_var.set(self._to_win_path(folder))
            if self.same_dir_var.get():
                self.dst_var.set(self._to_win_path(folder))
            self._refresh_file_list()

    def _select_dst(self) -> None:
        folder = filedialog.askdirectory(title="PDF를 저장할 폴더를 선택하세요")
        if folder:
            self.dst_var.set(self._to_win_path(folder))

    def _toggle_same_dir(self) -> None:
        if self.same_dir_var.get():
            src = self.src_var.get().strip()
            if src:
                self.dst_var.set(src)
            self.dst_entry.config(state="disabled")
            self.dst_btn.config(state="disabled")
        else:
            self.dst_entry.config(state="normal")
            self.dst_btn.config(state="normal")

    def _find_hwp_files(self, src_folder: str) -> list[str]:
        extensions = [".hwp"]
        if self.hwpx_var.get():
            extensions.append(".hwpx")

        found: list[str] = []
        if self.subfolder_var.get():
            for dirpath, _dirnames, filenames in os.walk(src_folder):
                for name in filenames:
                    if os.path.splitext(name)[1].lower() in extensions:
                        found.append(os.path.join(dirpath, name))
        else:
            for name in os.listdir(src_folder):
                if os.path.splitext(name)[1].lower() in extensions:
                    found.append(os.path.join(src_folder, name))
        return sorted(found)

    def _refresh_file_list(self) -> None:
        src = self.src_var.get().strip()
        if not src or not os.path.isdir(src):
            return

        files = self._find_hwp_files(src)

        for widget in self.file_list_frame.winfo_children():
            widget.destroy()
        self._file_vars.clear()

        for filepath in files:
            var = tk.BooleanVar(value=True)
            self._file_vars[filepath] = var
            label = os.path.relpath(filepath, src)
            tk.Checkbutton(
                self.file_list_frame,
                text=label,
                variable=var,
                anchor="w",
                bg="white",
                justify="left",
            ).pack(fill="x", padx=8, pady=1)

        count = len(files)
        self.file_count_label.config(
            text=f"총 {count}개 파일 발견" if count else "HWP 파일이 없습니다",
            fg="black" if count else "gray",
        )
        self.select_all_var.set(True)

    def _toggle_select_all(self) -> None:
        value = self.select_all_var.get()
        for var in self._file_vars.values():
            var.set(value)

    def _log(self, msg: str) -> None:
        def append() -> None:
            self.log_box.config(state="normal")
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.log_box.config(state="disabled")

        self.root.after(0, append)

    def _set_status(self, text: str) -> None:
        self.root.after(0, lambda: self.status_label.config(text=text))

    def _set_progress(self, value: float) -> None:
        self.root.after(0, lambda: self.progress.config(value=value))

    def _start_conversion(self) -> None:
        src = self.src_var.get().strip()
        dst = src if self.same_dir_var.get() else self.dst_var.get().strip()

        if not src:
            messagebox.showwarning("경고", "변환할 폴더를 선택해주세요.")
            return
        if not dst:
            messagebox.showwarning("경고", "저장할 폴더를 선택해주세요.")
            return
        if not os.path.isdir(src):
            messagebox.showerror("오류", f"폴더를 찾을 수 없습니다:\n{src}")
            return

        selected = [fp for fp, var in self._file_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("경고", "변환할 파일을 하나 이상 선택해주세요.")
            return

        os.makedirs(dst, exist_ok=True)

        if self.engine_var.get() == "mspdf":
            proceed = messagebox.askokcancel(
                "작업 중 조작 주의",
                "Microsoft Print to PDF 변환 중에는 자동으로 인쇄창과 저장창을 조작합니다.\n\n"
                "변환이 끝날 때까지 마우스와 키보드를 움직이거나 다른 창을 클릭하지 마세요.\n\n"
                "계속 진행할까요?",
            )
            if not proceed:
                return

        if self._converting:
            return

        self._stop_event.clear()
        self._converting = True
        self.convert_btn.config(state="disabled")
        self.stop_btn.config(state="normal")

        thread = threading.Thread(target=self._convert_thread, args=(selected, dst), daemon=True)
        thread.start()

    def _stop_conversion(self) -> None:
        self._stop_event.set()
        self._log("중지 요청됨. 현재 파일 처리 후 멈춥니다.")
        self._set_status("중지 중... 현재 파일 완료 대기")
        self.stop_btn.config(state="disabled")

    @staticmethod
    def _get_hwp_pid(_hwp) -> int:
        pids: list[int] = []

        def find(hwnd, _param) -> bool:
            try:
                if "Hwp" in win32gui.GetClassName(hwnd):
                    _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
                    pids.append(pid)
            except Exception:
                pass
            return True

        win32gui.EnumWindows(find, None)
        return pids[0] if pids else 0

    @staticmethod
    def _popup_watcher(stop_event: threading.Event, hwp_pid: int = 0) -> None:
        def click_child(hwnd, pid: int) -> None:
            def cb(child, _param) -> bool:
                try:
                    if win32gui.IsWindowVisible(child) and win32gui.IsWindowEnabled(child):
                        text = win32gui.GetWindowText(child)
                        if text == "모두 허용":
                            win32api.SendMessage(child, BM_CLICK, 0, 0)
                        elif text in ("확인", "OK", "닫기") and pid and pid == hwp_pid:
                            win32api.SendMessage(child, BM_CLICK, 0, 0)
                except Exception:
                    pass
                return True

            try:
                win32gui.EnumChildWindows(hwnd, cb, None)
            except Exception:
                pass

        def scan_top(hwnd, _param) -> bool:
            try:
                if win32gui.IsWindowVisible(hwnd):
                    _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
                    click_child(hwnd, pid)
            except Exception:
                pass
            return True

        while not stop_event.is_set():
            try:
                win32gui.EnumWindows(scan_top, None)
            except Exception:
                pass
            time.sleep(0.3)

    @staticmethod
    def _pdf_is_complete(output_path: str) -> bool:
        """PDF 파일이 끝까지 정상적으로 기록됐는지 확인한다.
        정상 PDF는 파일 끝에 '%%EOF' 마커가 있다. 스풀러가 아직 다 쓰지
        못한 잘린 파일에는 이 마커가 없으므로, 크기 안정화만으로는 잡지
        못하는 미완성(손상) 파일을 걸러낸다."""
        try:
            size = os.path.getsize(output_path)
            if size <= 0:
                return False
            with open(output_path, "rb") as f:
                # 끝부분 2KB만 읽어 EOF 마커 탐색(%%EOF 뒤 공백/개행 허용)
                f.seek(max(0, size - 2048))
                tail = f.read()
            return b"%%EOF" in tail
        except OSError:
            return False

    @staticmethod
    def wait_for_pdf_saved(output_path: str, stop_event: threading.Event, timeout: int = 600) -> None:
        """파일 생성, 잠금 해제, 크기 안정화, PDF 무결성(EOF)을 모두 확인한 뒤 반환한다."""
        deadline = time.time() + timeout
        prev_size = -1
        stable_count = 0

        while time.time() < deadline:
            if stop_event.is_set():
                raise RuntimeError("변환 중지됨")

            if os.path.exists(output_path):
                try:
                    size = os.path.getsize(output_path)
                    if size > 0:
                        with open(output_path, "rb"):
                            pass

                        if size == prev_size:
                            stable_count += 1
                            # 크기가 안정됐어도 EOF 마커가 있어야 완료로 인정
                            # (대용량 인쇄 중 일시 정체를 완료로 오판하지 않도록)
                            if stable_count >= 3 and HwpToPdfApp._pdf_is_complete(output_path):
                                return
                        else:
                            prev_size = size
                            stable_count = 0
                except OSError:
                    stable_count = 0

            time.sleep(0.5)

        if not os.path.exists(output_path):
            raise RuntimeError("PDF 파일이 생성되지 않았습니다.")
        if os.path.getsize(output_path) == 0:
            raise RuntimeError("PDF 파일이 비어 있습니다.")
        if not HwpToPdfApp._pdf_is_complete(output_path):
            raise RuntimeError("PDF가 끝까지 저장되지 않았습니다(손상). 제한 시간을 초과했거나 인쇄가 중단되었습니다.")
        raise RuntimeError("PDF 저장 완료를 제한 시간 안에 확인하지 못했습니다.")

    @staticmethod
    def _save_via_ms_print_to_pdf(
        hwp,
        hwp_pid: int,
        output_path: str,
        stop_event: threading.Event,
        timeout: int = 600,
    ) -> None:
        import pyautogui
        import pyperclip

        pyautogui.FAILSAFE = False

        print_titles = ("인쇄", "Print")
        save_titles = (
            "다음 이름으로 프린터 출력 저장",
            "Save Print Output As",
            "인쇄 출력을 다른 이름으로 저장",
        )
        error_box = [None]

        def wait_window(titles, seconds):
            # 예전 정상 버전 방식: HWP 인쇄 호출과 별도 watcher를 동시에 돌리며 창이 뜰 때까지 기다린다.
            deadline = time.time() + seconds
            while time.time() < deadline:
                if stop_event.is_set():
                    return None
                for title in titles:
                    hwnd = win32gui.FindWindow(None, title)
                    if hwnd and win32gui.IsWindowVisible(hwnd):
                        return hwnd
                time.sleep(0.2)
            return None

        def front(hwnd):
            # pyautogui 입력은 실제 포커스가 필요하다. 사용자가 다른 창을 쓰는 중이거나
            # 저장창이 보조 모니터에 뜨면 SetForegroundWindow가 '포커스 도용 방지'에 막혀
            # 입력이 엉뚱한 창으로 가 멈춘다. AttachThreadInput으로 대상/현재 포그라운드
            # 스레드 입력 큐를 붙여 강제로 포그라운드를 통과시킨다.
            try:
                user32 = ctypes.windll.user32
                cur_tid = win32api.GetCurrentThreadId()
                fg = user32.GetForegroundWindow()
                fg_tid = win32process.GetWindowThreadProcessId(fg)[0] if fg else 0
                tgt_tid = win32process.GetWindowThreadProcessId(hwnd)[0]
                tids = {t for t in (fg_tid, tgt_tid) if t and t != cur_tid}
                for t in tids:
                    try:
                        user32.AttachThreadInput(cur_tid, t, True)
                    except Exception:
                        pass
                try:
                    win32gui.ShowWindow(hwnd, 9)
                    user32.BringWindowToTop(hwnd)
                    win32gui.SetForegroundWindow(hwnd)
                    user32.SetActiveWindow(hwnd)
                finally:
                    for t in tids:
                        try:
                            user32.AttachThreadInput(cur_tid, t, False)
                        except Exception:
                            pass
                time.sleep(0.4)
            except Exception:
                pass

        def printer_list_y(hwnd_dlg) -> int:
            # 프린터 목록 컨트롤을 직접 못 찾는 경우를 대비해 예전 버전의 좌표 탐색 방식을 유지한다.
            left, top, right, bottom = win32gui.GetWindowRect(hwnd_dlg)
            width = right - left
            height = bottom - top
            center_x = left + width // 2
            skip = {"Button", "Static", "Edit", "ScrollBar", "ComboBox", "#32770"}
            for pct in (0.20, 0.25, 0.15, 0.30, 0.35):
                y = top + int(height * pct)
                child = win32gui.WindowFromPoint((center_x, y))
                if child and win32gui.IsChild(hwnd_dlg, child):
                    if win32gui.GetClassName(child) not in skip:
                        return y
            return top + int(height * 0.25)

        def select_ms_print(hwnd_dlg) -> None:
            # 먼저 텍스트가 보이는 자식 컨트롤을 직접 클릭하고, 실패하면 예전처럼 목록 첫 항목 쪽을 선택한다.
            found = [None]

            def by_text(child, _):
                try:
                    if TARGET_PRINTER in win32gui.GetWindowText(child):
                        found[0] = child
                        return False
                except Exception:
                    pass
                return True

            try:
                win32gui.EnumChildWindows(hwnd_dlg, by_text, None)
            except Exception:
                pass

            if found[0]:
                rect = win32gui.GetWindowRect(found[0])
                pyautogui.click((rect[0] + rect[2]) // 2, (rect[1] + rect[3]) // 2)
                time.sleep(0.4)
                return

            left, top, right, bottom = win32gui.GetWindowRect(hwnd_dlg)
            pyautogui.click(left + (right - left) // 2, printer_list_y(hwnd_dlg))
            time.sleep(0.4)
            pyautogui.press("home")
            time.sleep(0.3)

        WM_SETTEXT = 0x000C
        WM_GETTEXT = 0x000D

        def set_save_name_and_submit(hwnd_dlg) -> bool:
            # 표준 저장 대화상자(#32770)의 파일명 Edit에 WM_SETTEXT로 전체 경로를 직접
            # 써넣고 '저장' 버튼을 BM_CLICK으로 누른다. 키보드 포커스가 전혀 필요 없어
            # 포커스 도용 방지·사용자 입력 간섭에 면역이다. 성공 여부를 돌려준다.
            # 주의: 탐색기형 저장창엔 검색창 등 다른 Edit도 있다 — 파일명 칸은
            # ComboBox 안에 들어 있으므로 부모 클래스가 ComboBox인 Edit을 우선한다.
            edit_box = [0]
            fallback_edit = [0]

            def find_edit(child, _):
                try:
                    if win32gui.GetClassName(child) == "Edit":
                        parent = win32gui.GetParent(child)
                        if parent and "ComboBox" in win32gui.GetClassName(parent):
                            edit_box[0] = child
                            return False
                        if not fallback_edit[0]:
                            fallback_edit[0] = child
                except Exception:
                    pass
                return True

            try:
                win32gui.EnumChildWindows(hwnd_dlg, find_edit, None)
            except Exception:
                pass
            if not edit_box[0]:
                edit_box[0] = fallback_edit[0]
            if not edit_box[0]:
                return False

            win32gui.SendMessage(edit_box[0], WM_SETTEXT, 0, output_path)
            time.sleep(0.2)
            # 실제로 써졌는지 검증(다른 프로세스 Edit이라 WM_GETTEXT로 읽는다)
            buf = ctypes.create_unicode_buffer(1024)
            ctypes.windll.user32.SendMessageW(edit_box[0], WM_GETTEXT, 1024, buf)
            if buf.value.strip() != output_path:
                return False

            # '저장' 버튼 탐색 → BM_CLICK. 못 찾으면 실패 처리(폴백이 이어받음).
            btn_box = [0]

            def find_btn(child, _):
                try:
                    if win32gui.GetClassName(child) == "Button":
                        text = win32gui.GetWindowText(child).replace("&", "")
                        if text.startswith("저장") or text.lower().startswith("save"):
                            btn_box[0] = child
                            return False
                except Exception:
                    pass
                return True

            try:
                win32gui.EnumChildWindows(hwnd_dlg, find_btn, None)
            except Exception:
                pass
            if not btn_box[0]:
                return False
            win32api.SendMessage(btn_box[0], BM_CLICK, 0, 0)
            return True

        def watcher():
            try:
                # HAction.Run("Print")는 창이 닫힐 때까지 막히므로, watcher가 먼저 떠서 창을 조작해야 한다.
                print_hwnd = wait_window(print_titles, 30)
                if print_hwnd is None:
                    error_box[0] = RuntimeError("HWP 인쇄창이 나타나지 않았습니다.")
                    return
                if stop_event.is_set():
                    error_box[0] = RuntimeError("변환 중지됨")
                    return

                front(print_hwnd)
                select_ms_print(print_hwnd)
                pyautogui.hotkey("alt", "d")
                time.sleep(0.5)

                save_hwnd = wait_window(save_titles, 30)
                if save_hwnd is None:
                    error_box[0] = RuntimeError("저장 다이얼로그가 나타나지 않았습니다.")
                    return
                if stop_event.is_set():
                    error_box[0] = RuntimeError("변환 중지됨")
                    return

                # 1차: 포커스가 전혀 필요 없는 Win32 메시지 직접 주입.
                #   (pyautogui 키 입력은 포커스 도용 방지/사용자 클릭/보조모니터 상황에서
                #    엉뚱한 창으로 들어가 경로가 안 써지는 간헐 실패가 반복됐음)
                if not set_save_name_and_submit(save_hwnd):
                    # 2차 폴백: 예전 클립보드 붙여넣기 방식(포커스 필요).
                    front(save_hwnd)
                    pyperclip.copy(output_path)
                    pyautogui.hotkey("ctrl", "a")
                    time.sleep(0.15)
                    pyautogui.hotkey("ctrl", "v")
                    time.sleep(0.4)
                    pyautogui.hotkey("alt", "s")
            except Exception as exc:
                error_box[0] = exc

        # ── 1차: 대화창 자동화 없이 인쇄 액션에 출력 파일명을 직접 지정 ──
        # 병합 워커의 Hancom PDF 인쇄와 동일한 검증된 기법(CreateAction "Print" +
        # PrintToFile/FileName). 저장창이 떠도 FileName이 자동 채워져 곧 닫힌다.
        # 드라이버가 FileName을 무시하고 저장창이 멈춰 있으면 메시지 주입으로 마무리.
        def try_direct_print() -> bool:
            try:
                import win32print
                # 'Microsoft Print to PDF' 프린터가 실제로 있어야만 시도.
                # (이름이 틀리면 기본 프린터로 종이 인쇄될 위험이 있으므로 필수 확인)
                names = [p[2] for p in win32print.EnumPrinters(
                    win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
                if TARGET_PRINTER not in names:
                    return False
                act = hwp.CreateAction("Print")
                pset = act.CreateSet()
                act.GetDefault(pset)
                pset.SetItem("PrinterName", TARGET_PRINTER)
                pset.SetItem("PrintMethod", 0)   # 기본 인쇄(모아 찍기 해제)
                pset.SetItem("PageLayout", 0)
                pset.SetItem("PrintToFile", 1)
                pset.SetItem("FileName", os.path.abspath(output_path))
                pset.SetItem("Pause", 0)
                act.Execute(pset)
                return True
            except Exception:
                return False

        if try_direct_print():
            # 파일이 생길 때까지 살피되, 저장창이 보이면(드라이버가 FileName 무시)
            # 포커스 불필요한 메시지 주입으로 경로를 채우고 저장을 누른다.
            deadline = time.time() + timeout
            while time.time() < deadline:
                if stop_event.is_set():
                    raise RuntimeError("변환 중지됨")
                if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                    break
                hwnd = None
                for title in save_titles:
                    h = win32gui.FindWindow(None, title)
                    if h and win32gui.IsWindowVisible(h):
                        hwnd = h
                        break
                if hwnd:
                    set_save_name_and_submit(hwnd)
                time.sleep(0.5)
            HwpToPdfApp.wait_for_pdf_saved(output_path, stop_event, timeout=timeout)
            return

        # ── 2차 폴백: 예전 대화창 자동화(인쇄창 → 저장창) ──
        watcher_thread = threading.Thread(target=watcher, daemon=True)
        watcher_thread.start()

        # 예전 정상 버전과 동일하게 COM 인쇄 호출은 메인 COM 스레드에서 블로킹 실행한다.
        hwp.HAction.Run("Print")
        watcher_thread.join(timeout=15)

        if error_box[0] is not None:
            raise error_box[0]

        HwpToPdfApp.wait_for_pdf_saved(output_path, stop_event, timeout=timeout)

    def _convert_thread(self, files: list[str], dst: str) -> None:
        pythoncom.CoInitialize()
        hwp = None
        stop_popup = threading.Event()
        threading.Thread(target=self._popup_watcher, args=(stop_popup, 0), daemon=True).start()

        try:
            total = len(files)
            self._log(f"총 {total}개 파일 변환을 시작합니다.\n")
            self._set_status(f"0 / {total} 변환 중...")
            self._set_progress(0)

            hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")

            hwp_pid = self._get_hwp_pid(hwp)
            stop_popup.set()
            stop_popup = threading.Event()
            threading.Thread(target=self._popup_watcher, args=(stop_popup, hwp_pid), daemon=True).start()

            success = 0
            fail = 0

            for idx, hwp_path in enumerate(files, 1):
                if self._stop_event.is_set():
                    self._log(f"\n중지됨: 완료 {success}개 / 실패 {fail}개")
                    self._set_status(f"중지됨 - 완료 {success}개 / 실패 {fail}개")
                    break

                filename = os.path.basename(hwp_path)
                pdf_name = os.path.splitext(filename)[0] + ".pdf"
                pdf_path = self._unique_path(dst, pdf_name)
                self._set_status(f"{idx} / {total} - {filename}")

                tmp_dir = tempfile.mkdtemp()
                try:
                    ext = os.path.splitext(hwp_path)[1].lower()
                    fmt = "HWPX" if ext == ".hwpx" else "HWP"
                    tmp_hwp = os.path.join(tmp_dir, f"src{ext}")
                    tmp_pdf = os.path.join(tmp_dir, "out.pdf")
                    shutil.copy2(hwp_path, tmp_hwp)
                    hwp.Open(tmp_hwp, fmt, "forceopen:true")

                    if self.engine_var.get() == "mspdf":
                        self._log(f"  [진행] {filename} - {TARGET_PRINTER} 선택 확인 후 인쇄")
                        self._save_via_ms_print_to_pdf(hwp, hwp_pid, tmp_pdf, self._stop_event)
                    else:
                        hwp.SaveAs(tmp_pdf, "PDF", "")
                        self.wait_for_pdf_saved(tmp_pdf, self._stop_event, timeout=60)

                    for _attempt in range(10):
                        try:
                            shutil.move(tmp_pdf, pdf_path)
                            break
                        except OSError:
                            time.sleep(1)
                    else:
                        raise RuntimeError("PDF 파일 이동 실패: 파일이 아직 잠겨 있습니다.")

                    self._log(f"  [완료] {filename}")
                    success += 1
                except Exception as exc:
                    self._log(f"  [실패] {filename}\n         -> {exc}")
                    fail += 1
                finally:
                    try:
                        hwp.Run("FileClose")
                    except Exception:
                        pass
                    shutil.rmtree(tmp_dir, ignore_errors=True)

                self._set_progress(idx / total * 100)

            else:
                self._log(f"\n완료: 성공 {success}개 / 실패 {fail}개")
                self._log(f"저장 위치: {dst}")
                self._set_status(f"완료 - 성공 {success}개 / 실패 {fail}개")
                self.root.after(
                    0,
                    lambda count=success: messagebox.showinfo(
                        "변환 완료",
                        f"{count}개 파일 변환 완료!",
                    ),
                )

        except Exception as exc:
            self._log(f"\n[예상치 못한 오류] {exc}")
            self._set_status("오류 발생")
        finally:
            stop_popup.set()
            if hwp:
                try:
                    hwp.Quit()
                except Exception:
                    pass
            self._converting = False
            self.root.after(0, lambda: self.convert_btn.config(state="normal"))
            self.root.after(0, lambda: self.stop_btn.config(state="disabled"))
            pythoncom.CoUninitialize()

    @staticmethod
    def _unique_path(folder: str, filename: str) -> str:
        base, ext = os.path.splitext(filename)
        path = os.path.join(folder, filename)
        counter = 1
        while os.path.exists(path):
            path = os.path.join(folder, f"{base}_{counter}{ext}")
            counter += 1
        return path


def _resource(name: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


if __name__ == "__main__":
    try:
        import win32com.client  # noqa: F401
    except ImportError as exc:
        root_for_error = tk.Tk()
        root_for_error.withdraw()
        messagebox.showerror(
            "모듈 없음",
            f"필수 모듈이 설치되지 않았습니다.\n\n{exc}\n\npip install pywin32",
        )
        raise SystemExit from exc

    root = tk.Tk()

    try:
        from PIL import Image, ImageTk

        img = Image.open(_resource("hwptopdf_icon.png")).convert("RGBA")
        photos = [ImageTk.PhotoImage(img.resize((size, size), Image.LANCZOS)) for size in (256, 128, 64, 48, 32, 16)]
        root._icon_photos = photos
        root.iconphoto(True, *photos)
    except Exception:
        try:
            root.iconbitmap(_resource("hwptopdf.ico"))
        except Exception:
            pass

    app = HwpToPdfApp(root)
    root.mainloop()
