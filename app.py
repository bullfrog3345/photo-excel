import io
import os
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
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.units import pixels_to_EMU
from PIL import Image, ImageOps

# ─────────────────────────────── 상수 및 설정 ───────────────────────────────
SLOT_MARKUP = re.compile(r"\{\{\s*이미지\s*(?::\s*([^}]*?))?\s*\}\}")
IMG_EXT = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff")
IMAGE_MARGIN_PX = 4  # 셀 안쪽 여백(px)

st.set_page_config(
    page_title="엑셀 사진 원스톱 (웹 버전)",
    page_icon="🖼",
    layout="wide",
)


# ─────────────────────────────── 유틸리티 함수 ───────────────────────────────
def natural_key(text: str) -> list:
    """파일명 자연 정렬 (img2 < img10)"""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", text)]


def cell_address(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def col_width_to_px(width: Optional[float]) -> float:
    """openpyxl 열 너비(문자 수) -> 대략적인 픽셀 변환"""
    w = width if width is not None else 8.43
    return max(20.0, w * 7.5)


def row_height_to_px(height: Optional[float]) -> float:
    """openpyxl 행 높이(pt) -> 픽셀 변환 (1pt ≈ 1.333px)"""
    h = height if height is not None else 15.0
    return max(15.0, h * 1.3333)


def get_merged_cell_bounds_px(ws, row: int, col: int) -> Tuple[int, int, int, int]:
    """해당 셀(병합 포함)의 시작 행/열과 전체 가로·세로 픽셀 크기를 계산"""
    min_r, min_c, max_r, max_c = row, col, row, col
    for m_range in ws.merged_cells.ranges:
        if (
            m_range.min_row <= row <= m_range.max_row
            and m_range.min_col <= col <= m_range.max_col
        ):
            min_r, min_c, max_r, max_c = (
                m_range.min_row,
                m_range.min_col,
                m_range.max_row,
                m_range.max_col,
            )
            break

    total_w = 0.0
    for c in range(min_c, max_c + 1):
        col_letter = get_column_letter(c)
        total_w += col_width_to_px(ws.column_dimensions[col_letter].width)

    total_h = 0.0
    for r in range(min_r, max_r + 1):
        total_h += row_height_to_px(ws.row_dimensions[r].height)

    return min_r, min_c, int(total_w), int(total_h)


def get_image_anchor_pos(img) -> Tuple[int, int]:
    """openpyxl 이미지 객체에서 1-based (row, col) 위치를 추출"""
    anchor = img.anchor
    if isinstance(anchor, (OneCellAnchor, TwoCellAnchor)):
        return anchor._from.row + 1, anchor._from.col + 1
    if isinstance(anchor, str):
        m = re.match(r"([A-Z]+)(\d+)", anchor.upper())
        if m:
            return int(m.group(2)), column_index_from_string(m.group(1))
    return 1, 1


def scan_excel_slots(excel_bytes: bytes) -> List[dict]:
    """엑셀 파일에서 {{이미지}} 표식과 기존 삽입된 이미지를 찾아 슬롯 목록 반환"""
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=False)
    slots = []

    for ws in wb.worksheets:
        sheet_slots: Dict[Tuple[int, int], dict] = {}

        # 1. 셀 텍스트에서 {{이미지}} 또는 {{이미지:파일명}} 탐색 (A1 로고 제외)
        for row in ws.iter_rows():
            for cell in row:
                if cell.row == 1 and cell.column == 1:
                    continue
                if cell.value is None:
                    continue
                text = str(cell.value).strip()
                match = SLOT_MARKUP.fullmatch(text)
                if match:
                    min_r, min_c, w_px, h_px = get_merged_cell_bounds_px(
                        ws, cell.row, cell.column
                    )
                    key = (min_r, min_c)
                    sheet_slots[key] = {
                        "sheet": ws.title,
                        "row": min_r,
                        "col": min_c,
                        "addr": cell_address(min_r, min_c),
                        "kind": "marker",
                        "label": (match.group(1) or "").strip(),
                        "marker": text,
                        "w_px": w_px,
                        "h_px": h_px,
                        "existing_pil": None,
                    }

        # 2. 시트에 이미 포함된 그림(_images) 탐색 (A1 로고 제외)
        for img in getattr(ws, "_images", []):
            r, c = get_image_anchor_pos(img)
            if r == 1 and c == 1:
                continue

            # 기존 이미지 바이너리 추출 (미리보기용)
            old_pil = None
            try:
                raw_data = img._data()
                old_pil = Image.open(io.BytesIO(raw_data)).convert("RGB")
            except Exception:
                pass

            # 병합 셀 범위 기준으로 키 정규화
            min_r, min_c, w_px, h_px = get_merged_cell_bounds_px(ws, r, c)
            key = (min_r, min_c)

            if key in sheet_slots:
                sheet_slots[key]["existing_pil"] = old_pil
            else:
                img_w = int(getattr(img, "width", w_px) or w_px)
                img_h = int(getattr(img, "height", h_px) or h_px)
                sheet_slots[key] = {
                    "sheet": ws.title,
                    "row": min_r,
                    "col": min_c,
                    "addr": cell_address(min_r, min_c),
                    "kind": "pic",
                    "label": "",
                    "marker": "",
                    "w_px": max(img_w, w_px),
                    "h_px": max(img_h, h_px),
                    "existing_pil": old_pil,
                }

        # 행, 열 순으로 정렬하여 추가
        for key in sorted(sheet_slots.keys()):
            slots.append(sheet_slots[key])

    wb.close()
    return slots


def process_and_insert_images(
    excel_bytes: bytes,
    slots: List[dict],
    assignments: Dict[int, str],
    photo_map: Dict[str, bytes],
    keep_aspect_ratio: bool = True,
) -> Tuple[bytes, int, List[str]]:
    """연결된 사진들을 엑셀 슬롯에 삽입(기존 사진 교체 포함)하고 결과 바이너리 반환"""
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    logs: List[str] = []
    ok_count = 0

    # 시트별로 처리할 슬롯 정리
    by_sheet: Dict[str, List[Tuple[dict, str]]] = {}
    for idx, photo_name in sorted(assignments.items()):
        if not photo_name or photo_name == "(선택 안 함)":
            continue
        slot = slots[idx]
        by_sheet.setdefault(slot["sheet"], []).append((slot, photo_name))

    for sheet_name, items in by_sheet.items():
        ws = wb[sheet_name]
        target_coords = {(s["row"], s["col"]) for s, _ in items}

        # 교체 대상 위치에 있는 기존 이미지 제거 (A1 로고는 유지)
        kept_images = []
        for img in getattr(ws, "_images", []):
            r, c = get_image_anchor_pos(img)
            if r == 1 and c == 1:
                kept_images.append(img)
                continue
            min_r, min_c, _, _ = get_merged_cell_bounds_px(ws, r, c)
            if (min_r, min_c) in target_coords:
                continue  # 새 사진으로 교체할 기존 이미지 삭제
            kept_images.append(img)
        ws._images = kept_images

        # 새 사진 삽입
        for slot, photo_name in items:
            raw_bytes = photo_map.get(photo_name)
            if not raw_bytes:
                logs.append(
                    f"❌ [{sheet_name}!{slot['addr']}] 사진 데이터를 찾을 수 없음: {photo_name}"
                )
                continue

            try:
                # EXIF 회전 정보 반영 및 RGB 변환
                pil_img = Image.open(io.BytesIO(raw_bytes))
                pil_img = ImageOps.exif_transpose(pil_img).convert("RGB")
                nat_w, nat_h = pil_img.size

                avail_w = max(20, slot["w_px"] - 2 * IMAGE_MARGIN_PX)
                avail_h = max(20, slot["h_px"] - 2 * IMAGE_MARGIN_PX)

                if keep_aspect_ratio:
                    scale = min(avail_w / nat_w, avail_h / nat_h)
                    new_w = max(1, int(nat_w * scale))
                    new_h = max(1, int(nat_h * scale))
                else:
                    new_w, new_h = avail_w, avail_h

                # 엑셀 용량 최적화를 위해 삽입 크기에 맞춰 리사이즈 후 메모리에 저장
                resized = pil_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                img_buf = io.BytesIO()
                resized.save(img_buf, format="JPEG", quality=90)
                img_buf.seek(0)

                xl_img = XLImage(img_buf)
                xl_img.width = new_w
                xl_img.height = new_h

                # 셀 여백(offset) 적용을 위해 OneCellAnchor 사용
                marker = AnchorMarker(
                    col=slot["col"] - 1,
                    colOff=pixels_to_EMU(IMAGE_MARGIN_PX),
                    row=slot["row"] - 1,
                    rowOff=pixels_to_EMU(IMAGE_MARGIN_PX),
                )
                size = XDRPositiveSize2D(
                    cx=pixels_to_EMU(new_w), cy=pixels_to_EMU(new_h)
                )
                xl_img.anchor = OneCellAnchor(_from=marker, ext=size)

                ws.add_image(xl_img)

                # {{이미지}} 표식 텍스트 지우기
                if slot["kind"] == "marker":
                    ws.cell(row=slot["row"], column=slot["col"]).value = None

                ok_count += 1
                logs.append(
                    f"✓ [{sheet_name}!{slot['addr']}] {photo_name} 삽입 완료 ({new_w}×{new_h}px)"
                )
            except Exception as e:
                logs.append(
                    f"❌ [{sheet_name}!{slot['addr']}] {photo_name} 실패: {e}"
                )

    out_buf = io.BytesIO()
    wb.save(out_buf)
    wb.close()
    out_buf.seek(0)
    return out_buf.getvalue(), ok_count, logs


def auto_assign_photos(slots: List[dict], photo_names: List[str]) -> Dict[int, str]:
    """1순위: {{이미지:파일명}} 일치 매칭 / 2순위: 남은 빈 칸에 순서대로 자동 연결"""
    assign: Dict[int, str] = {}
    by_name = {}
    for name in photo_names:
        low = name.casefold()
        stem = os.path.splitext(low)[0]
        by_name[low] = name
        by_name.setdefault(stem, name)

    used = set()

    # 1. 파일명 지정 슬롯 우선 매칭
    for i, slot in enumerate(slots):
        label = (slot.get("label") or "").strip().replace("\\", "/").split("/")[-1]
        if not label:
            continue
        key = label.casefold()
        matched = by_name.get(key) or by_name.get(os.path.splitext(key)[0])
        if matched and matched not in used:
            assign[i] = matched
            used.add(matched)

    # 2. 라벨 없는 빈 슬롯 순서대로 매칭
    free_iter = iter([p for p in photo_names if p not in used])
    for i, slot in enumerate(slots):
        if i in assign or (slot.get("label") or "").strip():
            continue
        nxt = next(free_iter, None)
        if nxt:
            assign[i] = nxt

    return assign


# ─────────────────────────────── Streamlit UI ───────────────────────────────
def main():
    st.title("🖼 엑셀 사진 원스톱 (아이패드·웹 버전)")
    st.caption(
        "엑셀 파일의 {{이미지}} 표식과 기존 사진 위치를 찾아 사진을 연결하고 한 번에 저장합니다."
    )

    # 세션 상태 초기화
    if "slots" not in st.session_state:
        st.session_state.slots = []
    if "excel_bytes" not in st.session_state:
        st.session_state.excel_bytes = None
    if "excel_name" not in st.session_state:
        st.session_state.excel_name = ""
    if "assignments" not in st.session_state:
        st.session_state.assignments = {}
    if "result_bytes" not in st.session_state:
        st.session_state.result_bytes = None

    # 1. 파일 업로드 영역 (상단 2단)
    up_col1, up_col2 = st.columns(2)
    with up_col1:
        uploaded_excel = st.file_uploader(
            "① 대상 엑셀 파일 선택 (.xlsx)",
            type=["xlsx", "xlsm"],
            help="사진을 넣을 칸에 {{이미지}} 또는 {{이미지:파일명.jpg}}가 적힌 엑셀 파일",
        )
    with up_col2:
        uploaded_photos = st.file_uploader(
            "② 점검 사진 선택 (여러 장 동시 선택 가능)",
            type=["jpg", "jpeg", "png", "bmp", "gif", "tif", "tiff"],
            accept_multiple_files=True,
        )

    # 사진 데이터 정리 (자연 정렬)
    photo_map: Dict[str, bytes] = {}
    if uploaded_photos:
        sorted_files = sorted(uploaded_photos, key=lambda f: natural_key(f.name))
        for f in sorted_files:
            photo_map[f.name] = f.getvalue()
    photo_names = list(photo_map.keys())

    # 엑셀 파일 변경 감지 및 슬롯 스캔
    if uploaded_excel is not None:
        raw_excel = uploaded_excel.getvalue()
        if raw_excel != st.session_state.excel_bytes:
            st.session_state.excel_bytes = raw_excel
            st.session_state.excel_name = uploaded_excel.name
            st.session_state.result_bytes = None
            with st.spinner("엑셀에서 사진 위치를 찾는 중..."):
                st.session_state.slots = scan_excel_slots(raw_excel)
            st.session_state.assignments = (
                auto_assign_photos(st.session_state.slots, photo_names)
                if photo_names
                else {}
            )

    slots: List[dict] = st.session_state.slots

    if not uploaded_excel:
        st.info("👆 위에서 대상 엑셀 파일과 사진들을 먼저 선택해 주세요.")
        with st.expander("💡 사용 방법 보기"):
            st.markdown(
                """
1. **대상 엑셀 선택**: 사진이 들어갈 셀에 `{{이미지}}` 또는 `{{이미지:101.jpg}}`가 적힌 엑셀을 올립니다. (A1 로고는 자동 보호됩니다.)
2. **사진 여러 장 선택**: 아이패드 사진 보관함이나 파일 앱에서 사진들을 한꺼번에 선택합니다.
3. **자동 연결 / 개별 수정**: 파일명이 일치하는 사진은 자동 연결되며, 각 칸의 선택 상자를 눌러 사진을 바꿀 수 있습니다.
4. **사진 삽입 및 다운로드**: 버튼을 누르면 사진이 삽입된 새 엑셀 파일(`파일명_사진삽입.xlsx`)을 바로 내려받을 수 있습니다.
"""
            )
        return

    if not slots:
        st.warning(
            "이 엑셀에서 {{이미지}} 표식이나 기존 사진을 찾지 못했습니다. "
            "셀에 {{이미지}}를 입력한 뒤 다시 올려주세요."
        )
        return

    st.divider()

    # 컨트롤 버튼 바
    c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.5, 2.5])
    with c1:
        if st.button("🪄 자동 연결 실행", use_container_width=True, disabled=not photo_names):
            st.session_state.assignments = auto_assign_photos(slots, photo_names)
            for i in range(len(slots)):
                st.session_state[f"sel_{i}"] = st.session_state.assignments.get(
                    i, "(선택 안 함)"
                )
            st.rerun()
    with c2:
        if st.button("🗑 연결 초기화", use_container_width=True):
            st.session_state.assignments = {}
            for i in range(len(slots)):
                st.session_state[f"sel_{i}"] = "(선택 안 함)"
            st.rerun()
    with c3:
        keep_ratio = st.checkbox("사진 비율 유지", value=True)
    with c4:
        linked_cnt = sum(
            1
            for v in st.session_state.assignments.values()
            if v and v != "(선택 안 함)"
        )
        st.markdown(
            f"**요약:** 전체 칸 {len(slots)}개 · 연결됨 {linked_cnt}개 · 업로드된 사진 {len(photo_names)}장"
        )

    # 메인 2단 작업 영역 (좌: 엑셀 슬롯 매칭 / 우: 업로드된 사진 갤러리)
    left_col, right_col = st.columns([1.3, 1.0], gap="large")

    with left_col:
        st.subheader("① 엑셀 사진 위치 및 연결 설정")
        options = ["(선택 안 함)"] + photo_names

        for i, slot in enumerate(slots):
            current_val = st.session_state.assignments.get(i, "(선택 안 함)")
            if current_val not in options:
                current_val = "(선택 안 함)"

            with st.container(border=True):
                row_c1, row_c2 = st.columns([2.2, 1.0])
                with row_c1:
                    title_tag = (
                        slot["label"]
                        if slot["label"]
                        else ("기존 사진 자리" if slot["kind"] == "pic" else f"위치 {i + 1}")
                    )
                    st.markdown(f"**#{i + 1}** `[{slot['sheet']}!{slot['addr']}]` — {title_tag}")
                    selected = st.selectbox(
                        f"칸 {i + 1} 사진 선택",
                        options=options,
                        index=options.index(current_val),
                        key=f"sel_{i}",
                        label_visibility="collapsed",
                    )
                    st.session_state.assignments[i] = selected

                with row_c2:
                    if selected != "(선택 안 함)" and selected in photo_map:
                        st.image(photo_map[selected], caption="연결됨", use_container_width=True)
                    elif slot.get("existing_pil") is not None:
                        st.image(slot["existing_pil"], caption="기존 사진", use_container_width=True)
                    else:
                        st.caption("미리보기 없음")

    with right_col:
        st.subheader(f"② 업로드된 사진 목록 ({len(photo_names)}장)")
        if not photo_names:
            st.info("위에서 사진 파일을 업로드하면 여기에 썸네일이 표시됩니다.")
        else:
            used_by = {
                v: k + 1
                for k, v in st.session_state.assignments.items()
                if v and v != "(선택 안 함)"
            }
            grid_cols = st.columns(3)
            for idx, name in enumerate(photo_names):
                with grid_cols[idx % 3]:
                    with st.container(border=True):
                        st.image(photo_map[name], use_container_width=True)
                        if name in used_by:
                            st.success(f"✓ {used_by[name]}번 칸")
                        st.caption(name)

    st.divider()

    # 하단 실행 및 다운로드 영역
    btn_col, dl_col = st.columns([1, 1])
    with btn_col:
        if st.button("🚀 사진 삽입 실행", type="primary", use_container_width=True):
            valid_assign = {
                k: v
                for k, v in st.session_state.assignments.items()
                if v and v != "(선택 안 함)"
            }
            if not valid_assign:
                st.error("연결된 사진이 최소 1개 이상 있어야 합니다.")
            else:
                with st.spinner("엑셀에 사진을 삽입하는 중입니다..."):
                    res_bytes, ok_cnt, logs = process_and_insert_images(
                        st.session_state.excel_bytes,
                        slots,
                        valid_assign,
                        photo_map,
                        keep_aspect_ratio=keep_ratio,
                    )
                st.session_state.result_bytes = res_bytes
                st.success(f"완료! 총 {ok_cnt}개의 사진이 삽입되었습니다.")
                with st.expander("상세 처리 로그 보기", expanded=True):
                    for line in logs:
                        st.text(line)

    with dl_col:
        if st.session_state.result_bytes:
            orig_stem = Path(st.session_state.excel_name).stem
            out_filename = f"{orig_stem}_사진삽입.xlsx"
            st.download_button(
                label=f"💾 결과 파일 다운로드 ({out_filename})",
                data=st.session_state.result_bytes,
                file_name=out_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )


if __name__ == "__main__":
    main()
