# -*- coding: utf-8 -*-
"""
엑셀 사진 원스톱 (Streamlit 웹·아이패드 버전)

원본(Windows + Excel + tkinter)의 '사진 연결' 화면을 웹으로 옮긴 버전입니다.
- 엑셀의 {{이미지}} / {{이미지:파일명}} 표식 칸과 기존 사진 자리를 찾아 사진을 연결·삽입
- ①에서 칸을 누르면 굵은 테두리로 표시, ②에서 사진을 누르면 그 칸에 연결 (버튼·드롭다운 없음)
- A1 회사 로고 보호, 인쇄 영역 밖 칸 제외, 기존 사진 교체/삭제 예약, EXIF 회전 보정
필요 라이브러리: streamlit, openpyxl, pillow, pillow-heif(아이패드 HEIC 사진)
"""
from __future__ import annotations

import base64
import io
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
import streamlit as st
import streamlit.components.v1 as components
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

# 폴더 선택 부품 (app.py와 같은 위치의 folder_picker 폴더가 필요)
try:
    _folder_picker = components.declare_component(
        "folder_picker", path=str(Path(__file__).parent / "folder_picker")
    )
    HAS_FOLDER_PICKER = True
except Exception:
    _folder_picker = None
    HAS_FOLDER_PICKER = False

# ─────────────────────────────── 상수 ───────────────────────────────
VERSION = "3.0-web"
SLOT_MARKUP = re.compile(r"\{\{\s*이미지\s*(?::\s*([^}]*?))?\s*\}\}")
IMG_EXT = ("jpg", "jpeg", "png", "heic", "heif", "webp", "bmp", "gif", "tif", "tiff")
IMAGE_MARGIN_PX = 4      # 셀 안쪽 여백(px)
THUMB = 320              # 화면용 미리보기 최대 크기(px)
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
def auto_fill(slots: List[dict], photo_names: List[str], current: Dict[int, str]) -> Dict[int, str]:
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
        "pick_target": 0, "folder_photos": {}, "folder_last_id": None,
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


def select_slot(i: int):
    st.session_state.pick_target = i


def unassign(i: int):
    st.session_state.assignments.pop(i, None)
    bump()


def pick_photo(name: str):
    """사진 목록에서 사진을 누르면, 선택된 칸에 연결합니다(필요하면 다음 빈 칸으로 이동)."""
    ss = st.session_state
    n = len(ss.slots)
    if not n:
        return
    tgt = min(ss.pick_target, n - 1)
    assign(tgt, name)
    if ss.get("auto_next", True):
        order = list(range(tgt + 1, n)) + list(range(0, tgt))
        nxt = next((k for k in order if k not in ss.assignments and k not in ss.deletes), None)
        if nxt is not None:
            ss.pick_target = nxt


def slot_box(i: int):
    """칸 카드를 만듭니다. key를 달아 두면 선택된 칸만 굵은 테두리로 표시할 수 있습니다."""
    try:
        return st.container(border=True, key=f"slot_{i}")
    except TypeError:  # 오래된 Streamlit 버전
        return st.container(border=True)


def photo_box(idx: int):
    try:
        return st.container(border=True, key=f"photo_{idx}")
    except TypeError:
        return st.container(border=True)


def ui_css(active: int):
    """① 칸 카드와 ② 사진 카드 전체를 '눌러지는 영역'으로 만듭니다.
    (카드 안에 투명한 버튼을 카드 크기로 펼쳐 놓는 방식 - 화면에는 버튼이 보이지 않습니다.)"""
    st.markdown(
        f"""<style>
        /* 카드 기준 위치 */
        [class*="st-key-slot_"], [class*="st-key-photo_"] {{ position: relative; }}
        /* 투명 버튼을 카드 전체로 확장 */
        [class*="st-key-selslot_"], [class*="st-key-pickbtn_"] {{
            position: absolute !important; inset: 0; width: 100% !important; height: 100% !important;
            z-index: 10; margin: 0 !important;
        }}
        [class*="st-key-selslot_"] button, [class*="st-key-pickbtn_"] button {{
            width: 100%; height: 100%; opacity: 0; cursor: pointer;
        }}
        /* 카드 안의 '연결 해제', '기존 사진 삭제'는 위에 두어 따로 누를 수 있게 */
        [class*="st-key-unlink_"], [class*="st-key-del_"] {{ position: relative; z-index: 20; }}
        /* 손가락/마우스가 올라간 카드 표시 */
        [class*="st-key-slot_"]:hover, [class*="st-key-photo_"]:hover {{
            box-shadow: 0 0 0 2px rgba(255, 75, 75, 0.45);
        }}
        /* 선택된 칸 */
        .st-key-slot_{active} {{
            border: 4px solid #ff4b4b !important;
            box-shadow: 0 0 0 3px rgba(255, 75, 75, 0.35);
            background: rgba(255, 75, 75, 0.06);
        }}
        </style>""",
        unsafe_allow_html=True,
    )


def unique_name(name: str, taken: set) -> str:
    low = {t.casefold() for t in taken}
    if name.casefold() not in low:
        return name
    stem, dot, ext = name.rpartition(".")
    if not dot:
        stem, ext = name, ""
    n = 1
    while True:
        cand = f"{stem}_{n}" + (f".{ext}" if ext else "")
        if cand.casefold() not in low:
            return cand
        n += 1


def ingest_folder_value(val, upload_names) -> Tuple[int, List[str]]:
    """폴더 선택 부품이 보내 온 사진을 세션에 추가합니다. (같은 선택은 한 번만 처리)"""
    ss = st.session_state
    if not isinstance(val, dict) or not val.get("id") or val["id"] == ss.folder_last_id:
        return 0, []
    ss.folder_last_id = val["id"]
    added = 0
    failed = list(val.get("failed") or [])
    for f in val.get("files") or []:
        try:
            raw = base64.b64decode(str(f["data"]).split(",", 1)[1])
        except Exception:
            failed.append(str(f.get("name", "?")))
            continue
        name = unique_name(str(f["name"]), set(ss.folder_photos) | set(upload_names))
        ss.folder_photos[name] = raw
        added += 1
    return added, failed


def get_thumb(name: str, data: bytes) -> Optional[bytes]:
    key = (name, len(data))
    cache = st.session_state.thumbs
    if key not in cache:
        cache[key] = make_thumb(data)
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

    up1, up2 = st.columns([1, 1.5])
    with up1:
        up_excel = st.file_uploader(
            "① 대상 엑셀 (.xlsx)", type=["xlsx", "xlsm"],
            help="사진 넣을 칸에 {{이미지}} 또는 {{이미지:파일명.jpg}}가 적힌 엑셀",
        )

    def photo_uploader():
        return st.file_uploader(
            "② 점검 사진 (여러 장 선택 가능)", type=list(IMG_EXT), accept_multiple_files=True,
            help="아이패드 사진(HEIC)도 올릴 수 있습니다. 올린 사진은 아래 '② 사진 목록'에서 미리 보고 선택합니다.",
        )

    folder_val = None
    with up2:
        if HAS_FOLDER_PICKER:
            c_up, c_fold = st.columns([3, 1])      # Upload 버튼 오른쪽에 '폴더 선택'
            with c_up:
                up_photos = photo_uploader()
            with c_fold:
                folder_val = _folder_picker(key="folder_picker", default=None)
            st.caption("📁 폴더 선택: 폴더 안의 사진을 한 번에 불러옵니다(긴 변 1600px로 축소). 기기에 따라 폴더 대신 사진 여러 장 선택 창이 열릴 수 있습니다.")
        else:
            up_photos = photo_uploader()
            st.caption("⚠ folder_picker 폴더가 없어 '폴더 선택'을 쓸 수 없습니다. GitHub에 folder_picker 폴더도 올려 주세요.")

    upload_map: Dict[str, bytes] = {}
    for f in sorted(up_photos or [], key=lambda f: natural_key(f.name)):
        upload_map[f.name] = f.getvalue()

    added, failed = ingest_folder_value(folder_val, list(upload_map))
    if added:
        st.success(f"폴더에서 사진 {added}장을 불러왔습니다.")
    if failed:
        st.warning("불러오지 못한 사진: " + ", ".join(failed))
    if ss.folder_photos:
        fc1, fc2 = st.columns([3, 1])
        fc1.caption(f"📁 폴더에서 불러온 사진 {len(ss.folder_photos)}장")
        if fc2.button("폴더 사진 비우기"):
            ss.folder_photos = {}  # folder_last_id는 유지: 부품이 기억하는 이전 선택이 다시 들어오지 않게
            bump()
            st.rerun()

    if up_excel is not None:
        raw = up_excel.getvalue()
        if raw != ss.excel_bytes:
            ss.excel_bytes, ss.excel_name = raw, up_excel.name
            ss.assignments, ss.deletes = {}, set()
            ss.result_bytes, ss.pick_target, ss.upload_sig = None, 0, None
            bump()
            try:
                with st.spinner("엑셀에서 사진 위치를 찾는 중..."):
                    ss.slots = scan_excel_slots(raw)
            except Exception as e:
                ss.slots = []
                st.error(f"엑셀을 읽지 못했습니다: {e}")

    slots: List[dict] = ss.slots

    # 업로드한 사진이 바뀌면, 연결 안 된 칸만 자동 연결 (기존 연결은 유지)
    photo_map: Dict[str, bytes] = {**ss.folder_photos, **upload_map}
    sig = tuple(photo_map.keys())
    if sig != ss.upload_sig:
        ss.upload_sig = sig
        if slots and photo_map:
            ss.assignments = auto_fill(slots, sorted(photo_map, key=natural_key), ss.assignments)
        bump()

    photo_names = sorted(photo_map, key=natural_key)
    ss.assignments = {k: v for k, v in ss.assignments.items() if v in photo_map and k < len(slots)}
    ss.deletes = {k for k in ss.deletes if k < len(slots)}

    if up_excel is None and ss.excel_bytes is None:
        st.info("👆 대상 엑셀과 사진을 올려 주세요.")
        with st.expander("💡 사용 방법"):
            st.markdown(
                "1. 사진이 들어갈 셀에 `{{이미지}}` 또는 `{{이미지:101.jpg}}`를 적은 엑셀을 올립니다. (A1 로고는 자동 보호)\n"
                "2. 점검 사진을 여러 장 올립니다. 올린 사진은 미리보기로 확인하며 고를 수 있습니다.\n"
                "3. 파일명이 같은 사진은 자동 연결됩니다. 칸마다 선택 상자로 바꿀 수 있습니다.\n"
                "4. **사진 삽입** 후 **결과 파일 다운로드**를 누릅니다."
            )
        return

    if not slots:
        st.warning("이 엑셀에서 {{이미지}} 표식이나 기존 사진을 찾지 못했습니다. (인쇄 영역 밖의 칸은 제외됩니다.)")
        return

    st.divider()
    c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.5, 3])
    with c1:
        if st.button("🪄 자동 연결", disabled=not photo_names):
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

    ss.pick_target = max(0, min(ss.pick_target, len(slots) - 1))
    ui_css(ss.pick_target)

    left, right = st.columns([1.3, 1.0], gap="large")

    with left:
        st.subheader("① 엑셀 사진 위치")
        st.caption("사진을 넣을 **칸을 누르면** 빨간 굵은 테두리로 표시됩니다.")
        for i, s in enumerate(slots):
            linked_name = ss.assignments.get(i)
            with slot_box(i):
                st.button("선택", key=f"selslot_{i}", on_click=select_slot, args=(i,))  # 화면에는 안 보이는 투명 버튼
                a, b = st.columns([1.4, 1.0])
                with a:
                    st.markdown(f"**{slot_title(i, s)}**")
                    if linked_name:
                        st.caption(f"연결: {linked_name}")
                    elif s["existing"] and i in ss.deletes:
                        st.caption("기존 사진 삭제 예정")
                    elif s["existing"]:
                        st.caption("기존 사진 있음")
                    else:
                        st.caption("비어 있음")
                    if linked_name:
                        st.button("연결 해제", key=f"unlink_{i}_{ss.ver}", on_click=unassign, args=(i,))
                    elif s["existing"]:
                        want_del = st.checkbox("기존 사진 삭제", value=i in ss.deletes, key=f"del_{i}_{ss.ver}")
                        (ss.deletes.add if want_del else ss.deletes.discard)(i)
                with b:
                    if linked_name and linked_name in photo_map:
                        t = get_thumb(linked_name, photo_map[linked_name])
                        if t:
                            st.image(t, caption="연결됨")
                    elif s["old_thumb"]:
                        st.image(s["old_thumb"], caption="삭제 예정" if i in ss.deletes else "기존 사진")
                    else:
                        st.caption("미리보기 없음")

    with right:
        st.subheader(f"② 사진 목록 ({len(photo_names)}장)")
        if not photo_names:
            st.info("사진을 올리면 여기에 미리보기가 표시됩니다.")
        else:
            st.checkbox("넣은 뒤 다음 빈 칸으로 이동", value=True, key="auto_next")
            used_by = {v: k + 1 for k, v in ss.assignments.items()}
            cols = st.columns(2)
            for idx, n in enumerate(photo_names):
                with cols[idx % 2]:
                    with photo_box(idx):
                        st.button("선택", key=f"pickbtn_{idx}", on_click=pick_photo, args=(n,))  # 투명 버튼
                        t = get_thumb(n, photo_map[n])
                        if t:
                            st.image(t)
                        else:
                            st.warning("미리보기를 만들 수 없는 사진입니다 (형식 확인)")
                        st.caption(n)
                        if n in used_by:
                            st.success(f"✓ {used_by[n]}번 칸에 연결됨")

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
