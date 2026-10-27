# Ranet Toolkit - Copyright (C) 2026 Region & Agriculture NETwork Cooperative, Jeonnam
# SPDX-License-Identifier: AGPL-3.0-only  (see LICENSE)
"""폴더 일괄병합 결과보고서(.hwpx) 생성.

한글 COM 없이 템플릿 hwpx(zip+XML)를 lxml로 열어 표를 복제·채운 뒤 다시 묶는다.
`Ranet ToolKit.py`에서 분리했다(동작 변경 없음).
"""

import copy
import datetime
import os
import shutil
import tempfile
import zipfile

from lxml import etree

from rtk_common import log_err


# ─────────────────────────────────────────────────────────────────────────
# HWPX 보고서 생성 헬퍼 (한글 COM 불필요 — 템플릿 표를 lxml로 복제·채워 .hwpx 생성)
# ─────────────────────────────────────────────────────────────────────────
_HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"


def _hx_ln(e) -> str:
    return etree.QName(e.tag).localname


def _hx_cells(tr) -> list:
    return [c for c in tr if _hx_ln(c) == "tc"]


def _hx_cell_addr(tc):
    for s in tc:
        if _hx_ln(s) == "cellAddr":
            return s
    return None


def _hx_cell_span(tc):
    for s in tc:
        if _hx_ln(s) == "cellSpan":
            return s
    return None


def _hx_col_of(tc):
    a = _hx_cell_addr(tc)
    return a.get("colAddr") if a is not None else None


def _hx_first_col0(tr):
    for tc in _hx_cells(tr):
        if _hx_col_of(tc) == "0":
            return tc
    return None


def _hx_set_tc_text(tc, text: str) -> None:
    # tc > subList > p > run > t (첫 t에 넣고 나머지 t·run 제거, linesegarray 삭제)
    sub = next((s for s in tc if _hx_ln(s) == "subList"), None)
    if sub is None:
        return
    ps = [p for p in sub if _hx_ln(p) == "p"]
    if not ps:
        return
    p = ps[0]
    for ch in list(p):
        if _hx_ln(ch) == "linesegarray":
            p.remove(ch)
    runs = [r for r in p if _hx_ln(r) == "run"]
    if not runs:
        runs = [etree.SubElement(p, f"{{{_HP_NS}}}run")]
    run = runs[0]
    ts = [t for t in run if _hx_ln(t) == "t"]
    if ts:
        ts[0].text = text
        for t in ts[1:]:
            run.remove(t)
    else:
        etree.SubElement(run, f"{{{_HP_NS}}}t").text = text
    for r in runs[1:]:
        p.remove(r)


def _hx_set_cell_by_col(tr, colstr: str, text: str) -> bool:
    for tc in _hx_cells(tr):
        if _hx_col_of(tc) == colstr:
            _hx_set_tc_text(tc, text)
            return True
    return False


def _hx_set_row_addr(tr, row_idx: int) -> None:
    for tc in _hx_cells(tr):
        a = _hx_cell_addr(tc)
        if a is not None:
            a.set("rowAddr", str(row_idx))


def _hx_set_rowspan(tc, rowspan: int) -> None:
    s = _hx_cell_span(tc)
    if s is not None:
        s.set("rowSpan", str(rowspan))


# 한 행 기준 셀 높이(HWPUNIT). 템플릿 개별 행 높이(3682)를 기준값으로 사용.
_HX_BASE_ROW_H = 2400


def _hx_cell_sz(tc):
    for s in tc:
        if _hx_ln(s) == "cellSz":
            return s
    return None


def _hx_set_cell_height(tc, height: int) -> None:
    # cellSz height를 행 수에 맞게 재설정해 표가 과도하게 커지는 것 방지
    s = _hx_cell_sz(tc)
    if s is not None:
        s.set("height", str(int(height)))


def _hx_normalize_heights(rows) -> None:
    # 각 셀 높이를 (기준행높이 × rowSpan)으로 재계산.
    # 세로병합 첫 셀이 템플릿 고정 높이(18410)를 그대로 복제해 표가 비대해지는 문제 해결.
    for r in rows:
        for tc in _hx_cells(r):
            span = 1
            sp = _hx_cell_span(tc)
            if sp is not None:
                try:
                    span = max(1, int(sp.get("rowSpan", "1")))
                except (TypeError, ValueError):
                    span = 1
            _hx_set_cell_height(tc, _HX_BASE_ROW_H * span)


def _hx_set_para_text(p, text: str) -> None:
    # 일반 단락(폴더 제목)의 첫 run/t 텍스트 교체
    for ch in list(p):
        if _hx_ln(ch) == "linesegarray":
            p.remove(ch)
    for run in p:
        if _hx_ln(run) == "run":
            ts = [t for t in run if _hx_ln(t) == "t"]
            if ts:
                ts[0].text = text
                for t in ts[1:]:
                    run.remove(t)
            return


def _hx_tbl_child(tbl, name: str):
    for c in tbl:
        if _hx_ln(c) == name:
            return c
    return None


# 표 아래쪽 바깥 여백(다음 폴더 제목까지 간격). 페이지 경계에서는 한글이 자동으로 죽임.
_HX_TABLE_GAP_BOTTOM = 4500


def _hx_set_tbl_bottom_gap(tbl, gap: int = _HX_TABLE_GAP_BOTTOM) -> None:
    om = _hx_tbl_child(tbl, "outMargin")
    if om is not None:
        om.set("bottom", str(gap))


def _hx_table_layout_fix(tbl, row_count: int) -> None:
    # 표 페이지 넘김/크기 보정.
    #  - pos treatAsChar="0": '글자처럼 취급' 해제 → 표가 페이지 경계에서 셀 단위로 나뉨.
    #  - sz height: 셀 높이 합과 맞춰 표 전체 높이 재계산(고정 큰 값 복제 방지).
    #  - outMargin bottom: 다음 폴더 제목까지 간격(페이지 넘어가면 한글이 자동으로 죽여 맨 위 정렬).
    pos = _hx_tbl_child(tbl, "pos")
    if pos is not None:
        pos.set("treatAsChar", "0")
    sz = _hx_tbl_child(tbl, "sz")
    if sz is not None:
        sz.set("height", str(_HX_BASE_ROW_H * max(1, row_count)))
    _hx_set_tbl_bottom_gap(tbl)


# 템플릿 header.xml에 추가한 오른쪽 정렬 paraPr id
_HX_RIGHT_PARA_PR = "24"


def _hx_make_date_para(ref_blank, when=None):
    # 빈 단락을 복제해 오른쪽 정렬 '변환일자' 단락 생성
    if when is None:
        when = datetime.datetime.now()
    p = copy.deepcopy(ref_blank)
    p.set("paraPrIDRef", _HX_RIGHT_PARA_PR)
    # 캐시된 줄배치 정보 제거(한글이 재계산)
    for ch in list(p):
        if _hx_ln(ch) == "linesegarray":
            p.remove(ch)
    runs = [r for r in p if _hx_ln(r) == "run"]
    run = runs[0] if runs else etree.SubElement(p, f"{{{_HP_NS}}}run")
    for t in [t for t in run if _hx_ln(t) == "t"]:
        run.remove(t)
    txt = "변환일자 : " + when.strftime("%y.%m.%d. %H:%M:%S")
    etree.SubElement(run, f"{{{_HP_NS}}}t").text = txt
    return p


# 폴더 타이틀용 paraPr(위 간격 ~2줄) / 상태 태그 색상 charPr
_HX_TITLE_PARA_PR = "25"
_HX_TAG_BLUE = "11"   # (성공)
_HX_TAG_RED = "12"    # (실패)/(실패의심)


def _hx_folder_status(fd):
    # 폴더 결과 → (라벨, 태그 charPrIDRef)
    if not fd.get("ok"):
        return "(실패)", _HX_TAG_RED
    if fd.get("warns"):
        return "(실패의심)", _HX_TAG_RED
    return "(성공)", _HX_TAG_BLUE


def _hx_set_folder_title(hp, base_text: str, tag_text: str, tag_charpr: str) -> None:
    # 폴더 제목 단락: 본문 run(검정) + 상태 태그 run(색상) 2개로 구성
    for ch in list(hp):
        if _hx_ln(ch) == "linesegarray":
            hp.remove(ch)
    runs = [r for r in hp if _hx_ln(r) == "run"]
    if not runs:
        runs = [etree.SubElement(hp, f"{{{_HP_NS}}}run")]
    base_run = runs[0]
    # 본문 텍스트(앞에 공백 1칸 포함해 태그와 띄움)
    for t in [t for t in base_run if _hx_ln(t) == "t"]:
        base_run.remove(t)
    etree.SubElement(base_run, f"{{{_HP_NS}}}t").text = base_text + " "
    for r in runs[1:]:
        hp.remove(r)
    # 상태 태그 run(색상)
    tag_run = etree.SubElement(hp, f"{{{_HP_NS}}}run")
    tag_run.set("charPrIDRef", tag_charpr)
    etree.SubElement(tag_run, f"{{{_HP_NS}}}t").text = tag_text


def _hx_find_tbls(body) -> list:
    out = []
    for p in [c for c in body if _hx_ln(c) == "p"]:
        for run in p:
            if _hx_ln(run) != "run":
                continue
            for rc in run:
                if _hx_ln(rc) == "tbl":
                    out.append((p, rc))
    return out


def build_hwpx_report(template_path: str, out_path: str, meta: dict, folders: list) -> str:
    # 템플릿 .hwpx를 풀어 section0.xml의 요약표/폴더표를 복제·채운 뒤 다시 묶는다.
    work = tempfile.mkdtemp(prefix="ranet_hwpx_")
    ext = os.path.join(work, "ext")
    try:
        with zipfile.ZipFile(template_path) as z:
            z.extractall(ext)
        secname = "Contents/section0.xml"
        secpath = os.path.join(ext, secname)
        tree = etree.parse(secpath)
        root = tree.getroot()
        body = [e for e in root.iter() if _hx_ln(e) == "sec"][0]

        tbls = _hx_find_tbls(body)
        sum_p, sum_tbl = tbls[0]
        fold_p, fold_tbl = tbls[1]
        children = [c for c in body if _hx_ln(c) == "p"]
        head_p = children[children.index(fold_p) - 1]   # 폴더표 앞 제목 단락

        ref_head = copy.deepcopy(head_p)
        ref_fold_p = copy.deepcopy(fold_p)

        # 폴더 제목과 표 사이에 넣을 빈 단락 1개 확보(둘째 시군부터 간격이 좁은 문제 해결)
        ref_blank = None
        for p in children:
            if p is head_p or p is fold_p:
                continue
            if any(_hx_ln(x) == "tbl" for x in p.iter()):
                continue
            if "".join(p.itertext()).strip():
                continue
            ref_blank = copy.deepcopy(p)
            break
        if ref_blank is None:
            ref_blank = etree.Element(f"{{{_HP_NS}}}p")

        # 요약표 값 채우기 (row1~4 메타, row5 처리결과 — col '1')
        srows = [r for r in sum_tbl if _hx_ln(r) == "tr"]
        for i, v in enumerate([meta["root_dir"], meta["out_dir"], meta["layout"], meta["page_numbers"]]):
            _hx_set_cell_by_col(srows[1 + i], "1", v)
        _hx_set_cell_by_col(srows[5], "1", meta["result"])
        _hx_set_tbl_bottom_gap(sum_tbl)   # 요약표 아래 첫 폴더 제목까지 간격

        # 폴더표 행 참조 추출 (원본 9행: 0헤더 1생성 2첫파일 3다음파일 7미병합 8깨짐)
        fr = [r for r in fold_tbl if _hx_ln(r) == "tr"]
        ref_header = copy.deepcopy(fr[0])
        ref_gen = copy.deepcopy(fr[1])
        ref_file_first = copy.deepcopy(fr[2])
        ref_file_rest = copy.deepcopy(fr[3])
        ref_unmerged = copy.deepcopy(fr[7])
        ref_warn = copy.deepcopy(fr[8])

        # 본문에서 샘플 제목·폴더표 제거(참조만 사용)
        for p in (head_p, fold_p):
            body.remove(p)

        # 제목 아래: 변환일자(오른쪽 정렬) + 빈 단락 1개를 요약표 앞에 삽입
        sum_idx = list(body).index(sum_p)
        body.insert(sum_idx, _hx_make_date_para(ref_blank))
        body.insert(sum_idx + 1, copy.deepcopy(ref_blank))

        # 요약표 뒤에 남은 빈 단락 제거(폴더 간 간격은 아래서 직접 제어)
        si = list(body).index(sum_p)
        for p in list(body)[si + 1:]:
            if _hx_ln(p) != "p":
                continue
            if any(_hx_ln(x) == "tbl" for x in p.iter()):
                continue
            if "".join(p.itertext()).strip():
                continue
            body.remove(p)

        def make_block(fd):
            hp = copy.deepcopy(ref_head)
            hp.set("paraPrIDRef", _HX_TITLE_PARA_PR)   # keepWithNext: 제목이 표와 떨어져 페이지 갈리지 않게
            tag, tag_cp = _hx_folder_status(fd)
            _hx_set_folder_title(hp, f"{fd['idx']}. 폴더명 : {fd['name']} 변환결과", tag, tag_cp)
            tp = copy.deepcopy(ref_fold_p)
            tbl = next(rc for run in tp if _hx_ln(run) == "run"
                       for rc in run if _hx_ln(rc) == "tbl")
            for r in [r for r in tbl if _hx_ln(r) == "tr"]:
                tbl.remove(r)
            rows = [copy.deepcopy(ref_header)]
            _hx_set_cell_by_col(rows[0], "3", "쪽 구분")   # 헤더: 용량(mb) → 쪽 구분
            # 생성파일
            g = copy.deepcopy(ref_gen)
            if fd["ok"]:
                _hx_set_cell_by_col(g, "2", f"{fd['out']} (총 {fd['pages']}쪽, {fd['out_mb']}mb)")
                _hx_set_cell_by_col(g, "3", "-")
                _hx_set_cell_by_col(g, "4", str(fd["pages"]))
                _hx_set_cell_by_col(g, "5", "PDF")
            else:
                _hx_set_cell_by_col(g, "2", "PDF 생성 안 됨 (병합 가능한 파일 없음)")
                for c in ("3", "4", "5"):
                    _hx_set_cell_by_col(g, c, "-")
            rows.append(g)
            # 병합 파일
            files = fd["files"]
            start_page = 1   # 통합 PDF에서 각 파일이 시작하는 쪽
            for i, rec in enumerate(files):
                if i == 0:
                    row = copy.deepcopy(ref_file_first)
                    c0 = _hx_first_col0(row)
                    if c0 is not None:
                        _hx_set_rowspan(c0, max(1, len(files)))
                else:
                    row = copy.deepcopy(ref_file_rest)
                _hx_set_cell_by_col(row, "1", str(rec["no"]))
                nm = rec["name"] + (f"  [{rec['status']}]" if rec["status"] else "")
                _hx_set_cell_by_col(row, "2", nm)
                # 빈 용량 칸에 통합 PDF 시작 쪽수 표기(예: 19쪽~)
                _hx_set_cell_by_col(row, "3", f"{start_page}쪽~")
                _hx_set_cell_by_col(row, "4", str(rec["pages"]))
                _hx_set_cell_by_col(row, "5", rec["type"])
                rows.append(row)
                # 통합 PDF에 실제 추가된 쪽수만큼 다음 파일 시작 쪽 누적.
                # added_pages(정확) 우선, 없으면 pages 문자열 파싱(모아찍기 'A→B' 형식이면 실패 → 0).
                if "added_pages" in rec:
                    try:
                        start_page += int(rec["added_pages"])
                    except (TypeError, ValueError) as e:
                        log_err("build_hwpx_report.added_pages", e, str(rec.get("added_pages")))
                else:
                    try:
                        start_page += int(rec["pages"])
                    except (TypeError, ValueError) as e:
                        log_err("build_hwpx_report.pages", e, str(rec.get("pages")))
            # 미병합
            um = fd["unmerged"]
            for i, nm in enumerate(um):
                row = copy.deepcopy(ref_unmerged)
                c0 = _hx_first_col0(row)
                if i == 0:
                    if c0 is not None:
                        _hx_set_rowspan(c0, max(1, len(um)))
                elif c0 is not None:
                    row.remove(c0)   # 둘째 행부터 구분 셀 제거(병합)
                _hx_set_cell_by_col(row, "1", str(i + 1))
                _hx_set_cell_by_col(row, "2", nm)
                _hx_set_cell_by_col(row, "3", "-")
                _hx_set_cell_by_col(row, "4", "-")
                _hx_set_cell_by_col(row, "5", os.path.splitext(nm)[1][1:].upper() or "-")
                rows.append(row)
            # 깨짐/실패 의심
            w = copy.deepcopy(ref_warn)
            _hx_set_cell_by_col(w, "0", "깨짐/실패 의심")
            _hx_set_cell_by_col(w, "2", " / ".join(fd["warns"]) if fd["warns"] else "없음")
            rows.append(w)
            # 구성(있으면)
            if fd.get("note"):
                cc = copy.deepcopy(ref_warn)
                _hx_set_cell_by_col(cc, "0", "구성")
                _hx_set_cell_by_col(cc, "2", fd["note"])
                rows.append(cc)
            _hx_normalize_heights(rows)
            for ri, r in enumerate(rows):
                _hx_set_row_addr(r, ri)
                tbl.append(r)
            tbl.set("rowCnt", str(len(rows)))
            _hx_table_layout_fix(tbl, len(rows))
            return hp, tp

        for fd in folders:
            hp, tp = make_block(fd)
            # 폴더 간 간격은 타이틀 paraPr '위 간격'으로 처리(페이지 첫줄이면 한글이 자동 무시).
            # 폴더명 단락은 표에 바로 붙임.
            body.append(hp)
            body.append(tp)

        enc = tree.docinfo.encoding or "UTF-8"
        sa = tree.docinfo.standalone
        tree.write(secpath, xml_declaration=True, encoding=enc, standalone=sa)
        if os.path.exists(out_path):
            os.remove(out_path)
        with zipfile.ZipFile(out_path, "w") as zf:
            mt = os.path.join(ext, "mimetype")
            if os.path.exists(mt):
                zf.write(mt, "mimetype", compress_type=zipfile.ZIP_STORED)
            for dp, _, fns in os.walk(ext):
                for fn in fns:
                    fp = os.path.join(dp, fn)
                    arc = os.path.relpath(fp, ext)
                    if arc == "mimetype":
                        continue
                    zf.write(fp, arc, compress_type=zipfile.ZIP_DEFLATED)
        with zipfile.ZipFile(out_path) as zf:
            assert zf.testzip() is None
            etree.parse(zf.open(secname))
        return out_path
    finally:
        shutil.rmtree(work, ignore_errors=True)
