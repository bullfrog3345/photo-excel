# -*- coding: utf-8 -*-
"""
엑셀 사진 원스톱 (Streamlit 웹·아이패드 버전)

원본(Windows + Excel + tkinter)의 '사진 연결' 화면을 웹으로 옮긴 버전입니다.
- 엑셀의 {{이미지}} / {{이미지:파일명}} 표식 칸과 기존 사진 자리를 찾아 사진을 연결·삽입
- 사진 촬영(카메라) → 선택한 칸에 바로 연결, 다음 빈 칸으로 자동 이동
- A1 회사 로고 보호, 인쇄 영역 밖 칸 제외, 기존 사진 교체/삭제 예약, EXIF 회전 보정
필요 라이브러리: streamlit, openpyxl, pillow, streamlit-back-camera-input(후면 카메라)
"""
from __future__ import annotations

import datetime
import io
import re
import zipfile
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

try:  # 후면 카메라용 부품 (없으면 기본 카메라로 자동 대체)
    from streamlit_back_camera_input import back_camera_input
    HAS_BACK_CAM = True
except Exception:
    HAS_BACK_CAM = False

# ─────────────────────────────── 상수 ───────────────────────────────
VERSION = "3.0-web"
SLOT_MARKUP = re.compile(r"\{\{\s*이미지\s*(?::\s*([^}]*?))?\s*\}\}")
IMG_EXT = ("jpg", "jpeg", "png", "bmp", "gif", "tif", "tiff")
IMAGE_MARGIN_PX = 4      # 셀 안쪽 여백(px)
THUMB = 240              # 화면용 미리보기 최대 크기(px)
CAM_MAX_SIDE = 2400      # 촬영본 긴 변 최대(px) - 메모리 절약
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


def normalize_photo(data: bytes, max_side: int = CAM_MAX_SIDE) -> bytes:
    """촬영본을 회전 보정 후 적당한 크기의 JPEG로 정리합니다."""
    im = Image.open(io.BytesIO(data))
    im = ImageOps.exif_transpose(im).convert("RGB")
    im.thumbnail((max_side, max_side))
    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=92)
    return buf.getvalue()


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


def suggest_photo_name(slots: List[dict], idx: int, taken: set) -> str:
    """선택한 칸의 표식({{이미지:...}})을 바탕으로 저장 파일명을 제안 (중복이면 _1, _2 …)."""
    if idx is None or idx >= len(slots):
        return datetime.datetime.now().strftime("IMG_%Y%m%d_%H%M%S.jpg")
    raw = (slots[idx].get("label") or "").strip() or f"칸_{idx + 1}"
    safe = sanitize_filename(raw) or f"칸_{idx + 1}"
    if safe.lower().endswith((".jpg", ".jpeg", ".png")):
        stem, ext = safe.rsplit(".", 1)[0], "." + safe.rsplit(".", 1)[1]
    else:
        stem, ext = safe, ".jpg"
    name = f"{stem}{ext}"
    low = {t.casefold() for t in taken}
    n = 1
    while name.casefold() in low:
        name = f"{stem}_{n}{ext}"
        n += 1
    return name


def finalize_name(raw: str, taken: set) -> str:
    """사용자가 입력한 파일명을 안전하게 정리하고 중복을 피합니다."""
    name = sanitize_filename(raw) or datetime.datetime.now().strftime("IMG_%Y%m%d_%H%M%S")
    if not name.lower().endswith((".jpg", ".jpeg", ".png")):
        name += ".jpg"
    stem, ext = name.rsplit(".", 1)[0], "." + name.rsplit(".", 1)[1]
    low = {t.casefold() for t in taken}
    out, n = name, 1
    while out.casefold() in low:
        out = f"{stem}_{n}{ext}"
        n += 1
    return out


# ─────────────────────────────── 세션 상태 ───────────────────────────────
def init_state():
    defaults = {
        "excel_bytes": None, "excel_name": "", "slots": [],
        "assignments": {}, "deletes": set(), "ver": 0,
        "captured": {}, "cam_n": 0, "cam_target": 0,
        "upload_sig": None, "result_bytes": None, "thumbs": {}, "flash": "",
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


def get_thumb(name: str, data: bytes) -> Optional[bytes]:
    key = (name, len(data))
    cache = st.session_state.thumbs
    if key not in cache:
        cache[key] = make_thumb(data)
    return cache[key]


def slot_title(i: int, s: dict) -> str:
    tag = s["label"] or ("기존 사진 자리" if s["kind"] == "pic" else "빈 칸")
    return f"#{i + 1}  [{s['sheet']}!{s['addr']}]  {tag}"


# ─────────────────────────────── 카메라 ───────────────────────────────
def camera_section(slots: List[dict], upload_names: List[str]):
    ss = st.session_state
    with st.container(border=True):
        st.markdown("#### 📷 사진 촬영")
        on = st.checkbox("카메라 켜기 (처음에는 Safari가 카메라 권한을 물어봅니다 → 허용)", key="cam_on")
        if not on:
            st.caption("촬영한 사진은 아래 사진 목록에 추가되고, 선택한 칸에 바로 연결할 수 있습니다.")
            return

        taken = set(upload_names) | set(ss.captured) | {v for v in ss.assignments.values()}
        tgt = None
        if slots:
            ss.cam_target = min(ss.cam_target, len(slots) - 1)
            tgt = st.selectbox(
                "촬영한 사진을 연결할 칸",
                options=list(range(len(slots))),
                index=ss.cam_target,
                format_func=lambda i: slot_title(i, slots[i]),
                key=f"camtgt_{ss.ver}",
            )
            ss.cam_target = tgt
        auto = st.checkbox("촬영 즉시 위 칸에 연결하고, 다음 빈 칸으로 이동", value=True, key="cam_auto")
        name_in = st.text_input(
            "저장 파일명",
            value=suggest_photo_name(slots, tgt, taken),
            key=f"camname_{ss.cam_n}_{tgt}_{ss.ver}",
        )
        OPT_NATIVE = "기기 카메라 앱 (후면·고화질, 아이패드 권장)"
        OPT_REAR = "앱 내 후면 카메라 (영상을 터치하면 촬영)"
        OPT_DEFAULT = "기본 카메라 (기기에 따라 전·후면 전환 버튼)"
        opts = [OPT_NATIVE] + ([OPT_REAR] if HAS_BACK_CAM else []) + [OPT_DEFAULT]
        try:  # 안드로이드는 기본 카메라에 전환 버튼이 생기므로 그쪽을 먼저 선택
            ua = str(st.context.headers.get("User-Agent", ""))
        except Exception:
            ua = ""
        mode = st.radio(
            "카메라 방식", opts, index=opts.index(OPT_DEFAULT) if "Android" in ua else 0,
            horizontal=False, key="cam_side",
        )
        if mode == OPT_NATIVE:
            st.caption("아래 칸을 누르고 **‘사진 찍기’**를 고르면 기기의 카메라 앱이 열립니다(후면). 찍은 뒤 **‘사진 사용’**을 누르면 저장됩니다.")
            shot = st.file_uploader(
                "카메라 앱으로 촬영", type=list(IMG_EXT), key=f"native_{ss.cam_n}", label_visibility="collapsed",
            )
        elif mode == OPT_REAR:
            st.caption("카메라 화면(영상)을 **한 번 터치**하면 촬영됩니다.")
            shot = back_camera_input(height=450, width=480, key=f"bcam_{ss.cam_n}")
        else:
            shot = st.camera_input("셔터를 누르면 저장됩니다", key=f"cam_{ss.cam_n}", label_visibility="collapsed")
        if shot is not None:
            try:
                data = normalize_photo(shot.getvalue())
            except Exception as e:
                st.error(f"사진 처리 실패: {e}")
                return
            name = finalize_name(name_in, taken)
            ss.captured[name] = data
            msg = f"✓ 촬영 저장: {name}"
            if auto and tgt is not None:
                assign(tgt, name)
                order = list(range(tgt + 1, len(slots))) + list(range(0, tgt))
                nxt = next((k for k in order if k not in ss.assignments and k not in ss.deletes), None)
                if nxt is not None:
                    ss.cam_target = nxt
                msg += f" → {tgt + 1}번 칸에 연결"
            ss.flash = msg
            ss.cam_n += 1
            st.rerun()

        if ss.captured:
            st.caption(f"이번 접속에서 촬영한 사진 {len(ss.captured)}장 (새로고침하면 사라지니, 아래 버튼으로 보관하세요)")
            zbuf = io.BytesIO()
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_STORED) as zf:
                for n, d in ss.captured.items():
                    zf.writestr(n, d)
            st.download_button(
                "💾 촬영 사진 모두 내려받기 (.zip)", zbuf.getvalue(),
                file_name=datetime.datetime.now().strftime("촬영사진_%Y%m%d_%H%M.zip"),
                mime="application/zip",
            )


# ─────────────────────────────── 메인 화면 ───────────────────────────────
def main():
    init_state()
    ss = st.session_state

    st.title("🖼 엑셀 사진 원스톱")
    st.caption(f"v{VERSION} · 엑셀의 사진 위치를 찾아 사진을 연결하고 한 번에 삽입·저장합니다. (아이패드 Safari 지원)")
    if ss.flash:
        st.success(ss.flash)
        ss.flash = ""

    up1, up2 = st.columns(2)
    with up1:
        up_excel = st.file_uploader(
            "① 대상 엑셀 (.xlsx)", type=["xlsx", "xlsm"],
            help="사진 넣을 칸에 {{이미지}} 또는 {{이미지:파일명.jpg}}가 적힌 엑셀",
        )
    with up2:
        up_photos = st.file_uploader(
            "② 점검 사진 (여러 장 선택 가능)", type=list(IMG_EXT), accept_multiple_files=True,
            help="아이패드: 선택 창에서 '사진 보관함' 또는 '사진 찍기'를 고를 수 있습니다.",
        )

    upload_map: Dict[str, bytes] = {}
    for f in sorted(up_photos or [], key=lambda f: natural_key(f.name)):
        upload_map[f.name] = f.getvalue()

    if up_excel is not None:
        raw = up_excel.getvalue()
        if raw != ss.excel_bytes:
            ss.excel_bytes, ss.excel_name = raw, up_excel.name
            ss.assignments, ss.deletes = {}, set()
            ss.result_bytes, ss.cam_target, ss.upload_sig = None, 0, None
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
            ss.assignments = auto_fill(slots, list(upload_map), ss.assignments)
        bump()

    # 카메라 (엑셀을 올린 뒤에 칸 선택 가능)
    camera_section(slots, list(upload_map))

    photo_map: Dict[str, bytes] = {**ss.captured, **upload_map}
    photo_names = sorted(photo_map, key=natural_key)
    ss.assignments = {k: v for k, v in ss.assignments.items() if v in photo_map and k < len(slots)}
    ss.deletes = {k for k in ss.deletes if k < len(slots)}

    if up_excel is None and ss.excel_bytes is None:
        st.info("👆 대상 엑셀과 사진을 올려 주세요.")
        with st.expander("💡 사용 방법"):
            st.markdown(
                "1. 사진이 들어갈 셀에 `{{이미지}}` 또는 `{{이미지:101.jpg}}`를 적은 엑셀을 올립니다. (A1 로고는 자동 보호)\n"
                "2. 사진을 여러 장 올리거나, **📷 사진 촬영**으로 현장에서 바로 찍습니다.\n"
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

    left, right = st.columns([1.3, 1.0], gap="large")

    with left:
        st.subheader("① 엑셀 사진 위치")
        options = [NONE] + photo_names
        for i, s in enumerate(slots):
            cur = ss.assignments.get(i, NONE)
            with st.container(border=True):
                a, b = st.columns([2.2, 1.0])
                with a:
                    st.markdown(f"**{slot_title(i, s)}**")
                    sel = st.selectbox(
                        f"칸 {i + 1}", options, index=options.index(cur) if cur in options else 0,
                        key=f"sel_{i}_{ss.ver}", label_visibility="collapsed",
                    )
                    if sel == NONE:
                        ss.assignments.pop(i, None)
                    else:
                        ss.assignments[i] = sel
                        ss.deletes.discard(i)
                    if s["existing"] and sel == NONE:
                        want_del = st.checkbox(
                            "기존 사진 삭제", value=i in ss.deletes, key=f"del_{i}_{ss.ver}",
                        )
                        (ss.deletes.add if want_del else ss.deletes.discard)(i)
                with b:
                    shown = ss.assignments.get(i)
                    if shown and shown in photo_map:
                        t = get_thumb(shown, photo_map[shown])
                        if t:
                            st.image(t, caption="연결됨")
                    elif s["old_thumb"]:
                        st.image(s["old_thumb"], caption="삭제 예약" if i in ss.deletes else "기존 사진")
                    else:
                        st.caption("미리보기 없음")
        dup = [n for n in set(ss.assignments.values()) if list(ss.assignments.values()).count(n) > 1]
        if dup:
            st.warning("같은 사진이 여러 칸에 연결되어 있습니다: " + ", ".join(dup))

    with right:
        st.subheader(f"② 사진 목록 ({len(photo_names)}장)")
        if not photo_names:
            st.info("사진을 올리거나 촬영하면 여기에 표시됩니다.")
        else:
            used_by = {v: k + 1 for k, v in ss.assignments.items()}
            cols = st.columns(3)
            for idx, n in enumerate(photo_names):
                with cols[idx % 3]:
                    with st.container(border=True):
                        t = get_thumb(n, photo_map[n])
                        if t:
                            st.image(t)
                        if n in used_by:
                            st.success(f"✓ {used_by[n]}번 칸")
                        st.caption(("📷 " if n in ss.captured else "") + n)

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
