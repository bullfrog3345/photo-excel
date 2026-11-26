# -*- coding: utf-8 -*-
"""
엑셀 사진 원스톱 (Streamlit 웹·아이패드 버전)

원본(Windows + Excel + tkinter)의 '사진 연결' 화면을 웹으로 옮긴 버전입니다.
- 엑셀의 {{이미지}} / {{이미지:파일명}} 표식 칸과 기존 사진 자리를 찾아 사진을 연결·삽입
- 칸 순서대로 진행: 지금 칸을 크게 보여 주고, 사진을 고르면 다음 칸으로 자동 이동 (사진이 칸보다 많을 때 편함)
- A1 회사 로고 보호, 인쇄 영역 밖 칸 제외, 기존 사진 교체/삭제 예약, EXIF 회전 보정
필요 라이브러리: streamlit, openpyxl, pillow, pillow-heif(아이패드 HEIC 사진)
"""
from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
import streamlit as st
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import (
    AnchorMarker,
    OneCellAnchor,
    TwoCellAnchor,
)
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.utils.units import pixels_to_EMU
from PIL import Image, ImageOps

try:  # 아이패드 사진(HEIC) 읽기 지원 (없어도 JPG/PNG는 정상 동작)
    from pillow_heif import register_heif_opener
    register_heif_opener()
except Exception:
    pass

# ─────────────────────────────── 상수 ───────────────────────────────
VERSION = "3.1-web"
SLOT_MARKUP = re.compile(r"\{\{\s*이미지\s*(?::\s*([^}]*?))?\s*\}\}")
IMG_EXT = ("jpg", "jpeg", "png", "heic", "heif", "webp", "bmp", "gif", "tif", "tiff")
IMAGE_MARGIN_PX = 4      # 셀 안쪽 여백(px)
THUMB = 320              # 현재 칸 미리보기 최대 크기(px)
GRID_THUMB = 220         # 사진 격자 미리보기 크기(px) - 작게 해서 화면을 가볍게
PAGE = 24                # 사진 격자에 한 번에 보여 줄 장수
STORE_SCALE = 2          # 엑셀에 넣는 사진 해상도 배수(인쇄 품질)
NONE = "(선택 안 함)"

st.set_page_config(page_title="엑셀 사진 원스톱", page_icon="🖼", layout="wide")


# ─────────────────────────────── 유틸리티 ───────────────────────────────
def natural_key(text: str) -> list:
    """파일명 자연 정렬 (img2 < img10)"""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", text)]


def cell_address(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def sanitize_filename(raw: str) -> str:
    """파일명 금지 문자(\\ / : * ? " < > |)를 '_'로 바꿉니다."""
    return re.sub(r'[\\/*?:"<>|]', "_", raw).strip(" .")


def make_thumb(data: bytes, size: int = THUMB) -> Optional[bytes]:
    """화면 표시용 작은 JPEG 미리보기를 만듭니다 (EXIF 회전 반영)."""
    try:
        im = Image.open(io.BytesIO(data))
        im.draft("RGB", (size * 2, size * 2))
        im = ImageOps.exif_transpose(im).convert("RGB")
        im.thumbnail((size, size))
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=80)
        return buf.getvalue()
    except Exception:
        return None


# ─────────────────────── 셀 크기 계산 (엑셀 → 픽셀) ───────────────────────
def column_width_map(ws) -> Dict[int, float]:
    """열 번호 → 너비(문자 수). 숨김 열은 0. (column_dimensions를 건드려 값이 바뀌지 않도록 읽기만 함)"""
    out: Dict[int, float] = {}
    for key, dim in ws.column_dimensions.items():
        try:
            mn = dim.min or column_index_from_string(key)
        except Exception:
            continue
        mx = dim.max or mn
        if dim.hidden:
            w = 0.0
        elif dim.width is not None:
            w = float(dim.width)
        else:
            continue
        for c in range(mn, min(mx, mn + 300) + 1):
            out[c] = w
    return out


def col_px(widths: Dict[int, float], default_w: float, c: int) -> float:
    w = widths.get(c, default_w)
    return 0.0 if w <= 0 else w * 7.0 + 5.0


def row_px(ws, r: int) -> float:
    dim = ws.row_dimensions.get(r)
    if dim is not None and dim.hidden:
        return 0.0
    h = dim.height if (dim is not None and dim.height) else (ws.sheet_format.defaultRowHeight or 15.0)
    return float(h) * 96.0 / 72.0


def merged_bounds(ws, row: int, col: int) -> Tuple[int, int, int, int]:
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            return m.min_row, m.min_col, m.max_row, m.max_col
    return row, col, row, col


def area_px(ws, widths, default_w, row, col) -> Tuple[int, int, int, int]:
    """(시작 행, 시작 열, 가로 px, 세로 px) - 병합 셀이면 병합 영역 전체 크기."""
    min_r, min_c, max_r, max_c = merged_bounds(ws, row, col)
    w = sum(col_px(widths, default_w, c) for c in range(min_c, max_c + 1))
    h = sum(row_px(ws, r) for r in range(min_r, max_r + 1))
    return min_r, min_c, int(max(w, 20)), int(max(h, 15))


def print_boxes(ws) -> Optional[List[Tuple[int, int, int, int]]]:
    """인쇄 영역을 (r1, c1, r2, c2) 목록으로. 지정 안 했으면 None(=전체 사용)."""
    pa = ws.print_area
    if not pa:
        return None
    boxes = []
    for part in str(pa).split(","):
        rng = part.split("!")[-1].replace("$", "").strip()
        try:
            c1, r1, c2, r2 = range_boundaries(rng)
            boxes.append((r1, c1, r2, c2))
        except Exception:
            return None
    return boxes or None


def in_boxes(boxes, row: int, col: int) -> bool:
    if boxes is None:
        return True
    return any(r1 <= row <= r2 and c1 <= col <= c2 for r1, c1, r2, c2 in boxes)


def image_anchor_pos(img) -> Tuple[int, int]:
    """openpyxl 이미지의 시작 위치 (행, 열) - 1부터 시작."""
    anchor = img.anchor
    if isinstance(anchor, (OneCellAnchor, TwoCellAnchor)):
        return anchor._from.row + 1, anchor._from.col + 1
    if isinstance(anchor, str):
        m = re.match(r"([A-Z]+)(\d+)", anchor.upper())
        if m:
            return int(m.group(2)), column_index_from_string(m.group(1))
    return 1, 1


# ─────────────────────────── 엑셀: 사진 칸 찾기 ───────────────────────────
def scan_excel_slots(excel_bytes: bytes) -> List[dict]:
    """{{이미지}} 표식 칸과 기존 사진 자리를 찾습니다 (A1 로고·인쇄 영역 밖 제외)."""
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    slots: List[dict] = []
    for ws in wb.worksheets:
        widths = column_width_map(ws)
        default_w = ws.sheet_format.defaultColWidth or 8.43
        boxes = print_boxes(ws)
        found: Dict[Tuple[int, int], dict] = {}

        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None or (cell.row == 1 and cell.column == 1):
                    continue
                text = str(cell.value).strip()
                m = SLOT_MARKUP.fullmatch(text)
                if not m:
                    continue
                r0, c0, w, h = area_px(ws, widths, default_w, cell.row, cell.column)
                if not in_boxes(boxes, r0, c0):
                    continue
                found[(r0, c0)] = {
                    "sheet": ws.title, "row": r0, "col": c0, "addr": cell_address(r0, c0),
                    "kind": "marker", "label": (m.group(1) or "").strip(), "marker": text,
                    "w_px": w, "h_px": h, "existing": False, "old_thumb": None,
                }

        for img in getattr(ws, "_images", []):
            r, c = image_anchor_pos(img)
            if r == 1 and c == 1:
                continue  # A1 회사 로고
            r0, c0, w, h = area_px(ws, widths, default_w, r, c)
            if not in_boxes(boxes, r0, c0):
                continue
            thumb = None
            try:
                thumb = make_thumb(img._data())
            except Exception:
                pass
            if (r0, c0) in found:
                found[(r0, c0)]["existing"] = True
                found[(r0, c0)]["old_thumb"] = thumb
            else:
                iw = int(getattr(img, "width", w) or w)
                ih = int(getattr(img, "height", h) or h)
                found[(r0, c0)] = {
                    "sheet": ws.title, "row": r0, "col": c0, "addr": cell_address(r0, c0),
                    "kind": "pic", "label": "", "marker": "",
                    "w_px": max(iw, w), "h_px": max(ih, h), "existing": True, "old_thumb": thumb,
                }

        for key in sorted(found):
            slots.append(found[key])
    wb.close()
    return slots


# ─────────────────────────── 엑셀: 사진 삽입 ───────────────────────────
def insert_photos(
    excel_bytes: bytes,
    slots: List[dict],
    assignments: Dict[int, str],
    deletes: set,
    photo_map: Dict[str, bytes],
    fit: bool = True,
) -> Tuple[bytes, int, List[str]]:
    """연결된 사진을 삽입하고(기존 사진은 교체), 삭제 예약된 칸의 기존 사진은 지웁니다."""
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    logs: List[str] = []
    ok = 0

    todo: Dict[str, List[Tuple[dict, str]]] = {}
    for i, name in sorted(assignments.items()):
        if name and name != NONE and i < len(slots):
            todo.setdefault(slots[i]["sheet"], []).append((slots[i], name))
    del_by_sheet: Dict[str, set] = {}
    for i in deletes:
        if i < len(slots) and i not in assignments:
            s = slots[i]
            del_by_sheet.setdefault(s["sheet"], set()).add((s["row"], s["col"]))

    for sheet in set(todo) | set(del_by_sheet):
        ws = wb[sheet]
        replace = {(s["row"], s["col"]) for s, _ in todo.get(sheet, [])} | del_by_sheet.get(sheet, set())

        kept = []
        for img in getattr(ws, "_images", []):
            r, c = image_anchor_pos(img)
            if r == 1 and c == 1:
                kept.append(img)
                continue
            r0, c0, _, _ = area_px(ws, column_width_map(ws), ws.sheet_format.defaultColWidth or 8.43, r, c)
            if (r0, c0) in replace:
                if (r0, c0) in del_by_sheet.get(sheet, set()):
                    logs.append(f"🗑 [{sheet}!{cell_address(r0, c0)}] 기존 사진 삭제")
                continue
            kept.append(img)
        ws._images = kept

        for slot, name in todo.get(sheet, []):
            addr = f"{sheet}!{slot['addr']}"
            raw = photo_map.get(name)
            if not raw:
                logs.append(f"❌ [{addr}] 사진을 찾을 수 없음: {name}")
                continue
            try:
                im = ImageOps.exif_transpose(Image.open(io.BytesIO(raw))).convert("RGB")
                nat_w, nat_h = im.size
                avail_w = max(20, slot["w_px"] - 2 * IMAGE_MARGIN_PX)
                avail_h = max(20, slot["h_px"] - 2 * IMAGE_MARGIN_PX)
                if fit:
                    scale = min(avail_w / nat_w, avail_h / nat_h)
                    disp_w, disp_h = max(1, int(nat_w * scale)), max(1, int(nat_h * scale))
                else:
                    disp_w, disp_h = avail_w, avail_h
                # 인쇄 품질을 위해 표시 크기의 2배로 저장(원본보다 크게는 키우지 않음)
                store_w = min(nat_w, disp_w * STORE_SCALE)
                store_h = min(nat_h, disp_h * STORE_SCALE) if fit else min(nat_h, disp_h * STORE_SCALE)
                small = im.resize((max(1, store_w), max(1, store_h)), Image.Resampling.LANCZOS)
                buf = io.BytesIO()
                small.save(buf, format="JPEG", quality=90)
                buf.seek(0)

                xl = XLImage(buf)
                xl.width, xl.height = disp_w, disp_h
                marker = AnchorMarker(
                    col=slot["col"] - 1, colOff=pixels_to_EMU(IMAGE_MARGIN_PX),
                    row=slot["row"] - 1, rowOff=pixels_to_EMU(IMAGE_MARGIN_PX),
                )
                xl.anchor = OneCellAnchor(
                    _from=marker, ext=XDRPositiveSize2D(cx=pixels_to_EMU(disp_w), cy=pixels_to_EMU(disp_h))
                )
                ws.add_image(xl)
                if slot["kind"] == "marker":
                    ws.cell(row=slot["row"], column=slot["col"]).value = None
                ok += 1
                logs.append(f"✓ [{addr}] {name} 삽입 ({disp_w}×{disp_h}px)")
            except Exception as e:
                logs.append(f"❌ [{addr}] {name} 실패: {e}")

    out = io.BytesIO()
    wb.save(out)
    wb.close()
    return out.getvalue(), ok, logs


# ─────────────────────────── 자동 연결·이름 제안 ───────────────────────────
def auto_fill(slots: List[dict], photo_names: List[str], current: Dict[int, str], sequential: bool = True) -> Dict[int, str]:
    """이미 연결된 칸은 유지. 1) {{이미지:파일명}} 이름 일치 → 2) 이름 없는 빈 칸은 남은 사진을 순서대로."""
    assign = {k: v for k, v in current.items() if v in photo_names}
    by_name: Dict[str, str] = {}
    for n in photo_names:
        low = n.casefold()
        by_name[low] = n
        by_name.setdefault(Path(low).stem, n)
    used = set(assign.values())

    for i, s in enumerate(slots):
        if i in assign:
            continue
        label = (s.get("label") or "").strip().replace("\\", "/").split("/")[-1]
        if not label:
            continue
        key = label.casefold()
        hit = by_name.get(key) or by_name.get(Path(key).stem)
        if hit and hit not in used:
            assign[i] = hit
            used.add(hit)

    if not sequential:      # 파일명이 지정된 칸만 맞추고, 나머지는 직접 고르게 둠
        return assign
    free = iter([n for n in photo_names if n not in used])
    for i, s in enumerate(slots):
        if i in assign or (s.get("label") or "").strip():
            continue
        nxt = next(free, None)
        if nxt is None:
            break
        assign[i] = nxt
    return assign


# ─────────────────────────────── 세션 상태 ───────────────────────────────
def init_state():
    defaults = {
        "excel_bytes": None, "excel_name": "", "slots": [],
        "assignments": {}, "deletes": set(), "ver": 0,
        "pick_target": 0, "show_n": PAGE,
        "upload_sig": None, "result_bytes": None, "thumbs": {},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def bump():
    """프로그램이 연결을 바꾼 뒤 화면 위젯을 새 값으로 다시 그리게 합니다."""
    st.session_state.ver += 1
    st.session_state.result_bytes = None


def assign(slot: int, name: str):
    a = st.session_state.assignments
    for k in [k for k, v in a.items() if v == name and k != slot]:
        del a[k]  # 같은 사진은 한 칸에만
    a[slot] = name
    st.session_state.deletes.discard(slot)
    bump()


def goto_slot(i: int):
    st.session_state.pick_target = i


def unassign(i: int):
    st.session_state.assignments.pop(i, None)
    bump()


def more_photos():
    st.session_state.show_n += PAGE


def pick_photo(name: str):
    """사진을 고르면 현재 칸에 넣고, (옵션) 뒤쪽의 다음 빈 칸으로 이동합니다."""
    ss = st.session_state
    n = len(ss.slots)
    if not n:
        return
    cur = min(ss.pick_target, n - 1)
    assign(cur, name)
    if ss.get("auto_next", True):
        nxt = next((k for k in range(cur + 1, n) if k not in ss.assignments and k not in ss.deletes), None)
        if nxt is not None:
            ss.pick_target = nxt


def get_thumb(name: str, data: bytes, size: int = THUMB) -> Optional[bytes]:
    key = (name, len(data), size)
    cache = st.session_state.thumbs
    if key not in cache:
        cache[key] = make_thumb(data, size)
    return cache[key]


def slot_title(i: int, s: dict) -> str:
    tag = s["label"] or ("기존 사진 자리" if s["kind"] == "pic" else "빈 칸")
    return f"#{i + 1}  [{s['sheet']}!{s['addr']}]  {tag}"


# ─────────────────────────────── 메인 화면 ───────────────────────────────
def main():
    init_state()
    ss = st.session_state

    st.title("🖼 엑셀 사진 원스톱")
    st.caption(f"v{VERSION} · 엑셀의 사진 위치를 찾아 사진을 연결하고 한 번에 삽입·저장합니다. (아이패드 Safari 지원)")

    up1, up2 = st.columns(2)
    with up1:
        up_excel = st.file_uploader(
            "① 대상 엑셀 (.xlsx)", type=["xlsx", "xlsm"],
            help="사진 넣을 칸에 {{이미지}} 또는 {{이미지:파일명.jpg}}가 적힌 엑셀",
        )
    with up2:
        up_photos = st.file_uploader(
            "② 점검 사진 (여러 장 선택 가능)", type=list(IMG_EXT), accept_multiple_files=True,
            help="아이패드 사진(HEIC)도 올릴 수 있습니다. 올린 사진은 아래 '② 사진 목록'에서 미리 보고 선택합니다.",
        )

    upload_map: Dict[str, bytes] = {}
    for f in sorted(up_photos or [], key=lambda f: natural_key(f.name)):
        upload_map[f.name] = f.getvalue()

    if up_excel is not None:
        raw = up_excel.getvalue()
        if raw != ss.excel_bytes:
            ss.excel_bytes, ss.excel_name = raw, up_excel.name
            ss.assignments, ss.deletes = {}, set()
            ss.result_bytes, ss.pick_target, ss.upload_sig, ss.show_n = None, 0, None, PAGE
            bump()
            try:
                with st.spinner("엑셀에서 사진 위치를 찾는 중..."):
                    ss.slots = scan_excel_slots(raw)
            except Exception as e:
                ss.slots = []
                st.error(f"엑셀을 읽지 못했습니다: {e}")

    slots: List[dict] = ss.slots

    # 업로드한 사진이 바뀌면, 연결 안 된 칸만 자동 연결 (기존 연결은 유지)
    sig = tuple(upload_map.keys())
    if sig != ss.upload_sig:
        ss.upload_sig = sig
        if slots and upload_map:
            ss.assignments = auto_fill(slots, list(upload_map), ss.assignments, sequential=False)
        bump()

    photo_map: Dict[str, bytes] = dict(upload_map)
    photo_names = sorted(photo_map, key=natural_key)
    ss.assignments = {k: v for k, v in ss.assignments.items() if v in photo_map and k < len(slots)}
    ss.deletes = {k for k in ss.deletes if k < len(slots)}

    if up_excel is None and ss.excel_bytes is None:
        st.info("👆 대상 엑셀과 사진을 올려 주세요.")
        with st.expander("💡 사용 방법"):
            st.markdown(
                "1. 사진이 들어갈 셀에 `{{이미지}}` 또는 `{{이미지:101.jpg}}`를 적은 엑셀을 올립니다. (A1 로고는 자동 보호)\n"
                "2. 점검 사진을 여러 장 올립니다. 올린 사진은 미리보기로 확인하며 고를 수 있습니다.\n"
                "3. 화면 위에 지금 칸이 크게 나옵니다. 아래에서 사진을 고르면 그 칸에 들어가고 다음 칸으로 넘어갑니다.\n"
                "4. **사진 삽입** 후 **결과 파일 다운로드**를 누릅니다."
            )
        return

    if not slots:
        st.warning("이 엑셀에서 {{이미지}} 표식이나 기존 사진을 찾지 못했습니다. (인쇄 영역 밖의 칸은 제외됩니다.)")
        return

    st.divider()
    c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.5, 3])
    with c1:
        if st.button("🪄 빈 칸 자동 채우기", disabled=not photo_names, help="비어 있는 칸에 남은 사진을 이름 순서대로 채웁니다"):
            ss.assignments = auto_fill(slots, photo_names, ss.assignments)
            bump()
            st.rerun()
    with c2:
        if st.button("🗑 연결 초기화"):
            ss.assignments, ss.deletes = {}, set()
            bump()
            st.rerun()
    with c3:
        fit = st.checkbox("사진 비율 유지", value=True)
    with c4:
        linked = len(ss.assignments)
        blank = sum(1 for i, s in enumerate(slots) if i not in ss.assignments and not s["existing"])
        st.markdown(f"**요약:** 칸 {len(slots)}개 · 연결 {linked}개 · 비어 있는 칸 {blank}개 · 사진 {len(photo_names)}장")

    n_slots = len(slots)
    ss.pick_target = max(0, min(ss.pick_target, n_slots - 1))
    cur = ss.pick_target
    cs = slots[cur]
    linked_name = ss.assignments.get(cur)

    # ── 칸 이동줄: 번호를 누르면 그 칸으로 이동 (✓=사진 연결됨, 🗑=기존 사진 삭제 예정)
    st.markdown("**칸 이동**")
    per_row = 8
    for r0 in range(0, n_slots, per_row):
        cols = st.columns(per_row)
        for k in range(per_row):
            i = r0 + k
            if i >= n_slots:
                break
            mark = "✓" if i in ss.assignments else ("🗑" if i in ss.deletes else "")
            cols[k].button(
                f"{i + 1}{mark}", key=f"nav_{i}", on_click=goto_slot, args=(i,),
                type="primary" if i == cur else "secondary",
            )

    # ── 지금 칸
    with st.container(border=True):
        st.markdown(f"### 지금 칸 {cur + 1} / {n_slots}")
        st.caption(f"[{cs['sheet']}!{cs['addr']}]  {cs['label'] or ('기존 사진 자리' if cs['kind'] == 'pic' else '빈 칸')}")
        pa, pb = st.columns([1, 1.4])
        with pa:
            if linked_name and linked_name in photo_map:
                t = get_thumb(linked_name, photo_map[linked_name])
                if t:
                    st.image(t, caption=f"연결됨: {linked_name}")
            elif cs["old_thumb"]:
                st.image(cs["old_thumb"], caption="삭제 예정" if cur in ss.deletes else "기존 사진")
            else:
                st.caption("아직 사진이 없습니다.")
        with pb:
            st.button("◀ 이전 칸", key="prev_slot", on_click=goto_slot, args=(cur - 1,), disabled=cur == 0)
            st.button("건너뛰기 ▶", key="next_slot", on_click=goto_slot, args=(cur + 1,), disabled=cur >= n_slots - 1)
            if linked_name:
                st.button("연결 해제", key=f"unlink_{cur}_{ss.ver}", on_click=unassign, args=(cur,))
            elif cs["existing"]:
                want_del = st.checkbox("기존 사진 삭제", value=cur in ss.deletes, key=f"del_{cur}_{ss.ver}")
                (ss.deletes.add if want_del else ss.deletes.discard)(cur)

    if n_slots and all(i in ss.assignments or i in ss.deletes or slots[i]["existing"] for i in range(n_slots)) and ss.assignments:
        st.success("모든 칸이 채워졌습니다. 아래 '사진 삽입 실행'을 누르세요.")

    # ── 사진 고르기
    used = set(ss.assignments.values())
    hide_used = st.checkbox("사용한 사진 숨기기", value=True, key="hide_used")
    st.checkbox("넣은 뒤 다음 빈 칸으로 이동", value=True, key="auto_next")
    pool = [n for n in photo_names if not (hide_used and n in used)]
    st.subheader(f"사진 고르기 — 남은 사진 {len(pool)}장 / 전체 {len(photo_names)}장")
    if not photo_names:
        st.info("위에서 점검 사진을 올리면 여기에 미리보기가 나타납니다.")
    elif not pool:
        st.info("남은 사진이 없습니다. '사용한 사진 숨기기'를 끄면 모든 사진이 보입니다.")
    else:
        shown = pool[: ss.show_n]
        used_by = {v: k + 1 for k, v in ss.assignments.items()}
        grid = st.columns(3)
        for idx, n in enumerate(shown):
            with grid[idx % 3]:
                with st.container(border=True):
                    t = get_thumb(n, photo_map[n], GRID_THUMB)
                    if t:
                        st.image(t)
                    else:
                        st.warning("미리보기를 만들 수 없는 사진입니다")
                    st.caption(n)
                    if n in used_by:
                        st.success(f"✓ {used_by[n]}번 칸")
                    st.button("선택", key=f"pickbtn_{idx}_{ss.ver}", on_click=pick_photo, args=(n,))
        if len(pool) > len(shown):
            st.button(f"사진 더 보기 ({len(pool) - len(shown)}장 남음)", on_click=more_photos)

    with st.expander("연결 현황 보기"):
        for i, sl in enumerate(slots):
            nm = ss.assignments.get(i)
            state = f"✓ {nm}" if nm else ("🗑 기존 사진 삭제" if i in ss.deletes else ("기존 사진 유지" if sl["existing"] else "비어 있음"))
            st.text(f"{i + 1:>2}. [{sl['sheet']}!{sl['addr']}] {state}")

    st.divider()
    b1, b2 = st.columns(2)
    with b1:
        if st.button("🚀 사진 삽입 실행", type="primary"):
            if not ss.assignments and not ss.deletes:
                st.error("연결된 사진이나 삭제 예약이 최소 1개 필요합니다.")
            else:
                with st.spinner("엑셀에 사진을 삽입하는 중입니다..."):
                    res, ok, logs = insert_photos(ss.excel_bytes, slots, ss.assignments, ss.deletes, photo_map, fit)
                ss.result_bytes = res
                st.success(f"완료! 사진 {ok}개를 삽입했습니다.")
                with st.expander("상세 처리 로그", expanded=True):
                    for line in logs:
                        st.text(line)
    with b2:
        if ss.result_bytes:
            out_name = f"{Path(ss.excel_name).stem}_사진삽입.xlsx"
            st.download_button(
                f"💾 결과 파일 다운로드 ({out_name})", ss.result_bytes, file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary",
            )


if __name__ == "__main__":
    main()
