# Ranet Toolkit - Copyright (C) 2026 Region & Agriculture NETwork Cooperative, Jeonnam
# SPDX-License-Identifier: AGPL-3.0-only  (see LICENSE)
"""라넷툴킷 공통 유틸 - 실패 기록.

`Ranet ToolKit.py`와 형제 모듈(rtk_emf, rtk_hwpx_report)이 함께 쓴다.
여기 두는 이유: 두 모듈 다 log_err가 필요한데, 본체에서 가져오면 순환 참조가 된다.
"""

import datetime
import os
import tempfile


# ─── 실패 기록 (조용한 실패 방지) ───
# 실패해도 흐름은 그대로 이어가되, 흔적은 남긴다. GUI 앱이라 콘솔이 없으므로
# 최근작업(recent.json)과 같은 폴더에 파일로 쓴다.
#   %LOCALAPPDATA%\RanetToolkit\error.log
# 같은 지점은 5회까지만 쓰고 횟수는 계속 센다(대량 파일 루프에서 로그 폭주 방지).
ERROR_LOG_PATH = os.path.join(
    os.environ.get("LOCALAPPDATA", tempfile.gettempdir()), "RanetToolkit", "error.log")
_ERR_COUNT: dict = {}
_ERR_LOG_CAP = 5


def log_err(where: str, exc: BaseException, extra: str = "") -> None:
    # 예외를 삼키는 자리에서 호출한다. 기록만 하고 아무것도 되돌리지 않는다.
    n = _ERR_COUNT.get(where, 0) + 1
    _ERR_COUNT[where] = n
    if n > _ERR_LOG_CAP:
        return
    try:
        os.makedirs(os.path.dirname(ERROR_LOG_PATH), exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        tail = f" | {extra}" if extra else ""
        with open(ERROR_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{stamp}] {where}: {type(exc).__name__}: {exc}{tail}\n")
            if n == _ERR_LOG_CAP:
                f.write(f"[{stamp}] {where}: 이후 동일 지점 기록 생략(횟수는 계속 셈)\n")
    except Exception:
        pass          # 기록 자체가 실패하면 그냥 포기한다(여기서 또 터지면 안 됨)


def error_log_summary() -> str:
    # 지점별 누적 실패 횟수. 산발적 실패인지 전량 실패인지 구분하는 용도.
    if not _ERR_COUNT:
        return "실패 없음"
    items = sorted(_ERR_COUNT.items(), key=lambda kv: -kv[1])
    return " / ".join(f"{k} {v}회" for k, v in items)
