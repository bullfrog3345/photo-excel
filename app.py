# -*- coding: utf-8 -*-
"""
엑셀 사진 원스톱 v2.0
엑셀의 사진 위치를 찾고 사진 폴더의 이미지를 자동 연결한 뒤 한 번에 삽입·저장하는 단일 기능 앱.
필요 환경: Windows + Microsoft Excel, pip install xlwings tkinterdnd2 pillow
"""
from __future__ import annotations

import contextlib
import datetime
import json
import os
import platform
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter import font as tkfont
from urllib.parse import unquote


def _require(pip_name: str):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        f"{pip_name} 미설치",
        f"{pip_name} 라이브러리가 필요합니다.\n\n터미널에서 'pip install {pip_name}' 실행",
    )
    root.destroy()
    sys.exit(1)


try:
    import xlwings as xw
except ImportError:
    _require("xlwings")

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    _require("tkinterdnd2")


# ─────────────────────────────── 상수 ───────────────────────────────
VERSION = "2.2"
FONT = "Malgun Gothic"
# exe(PyInstaller)로 묶였을 때는 exe가 있는 폴더, 아니면 이 스크립트가 있는 폴더를 기준으로 한다.
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "recent_config.json"
TEMPLATE_DIR = BASE_DIR / "InspectionSheets"
SEOUL_TEMPLATE_DIR = TEMPLATE_DIR / "서울형 기계설비성능점검보고서_설비별"
EXCLUDED_FILES = {"0대상점검및수량.xlsx", "0엑셀 마크업.xlsx", "0정보 엑셀 파일.xlsx"}
IMAGE_MARKUP = re.compile(r"\{\{이미지:([^}]+)\}\}")
IMAGE_MARGIN = 3  # 셀 안쪽 여백(pt)

DEFAULT_EQUIPMENT = [
    "1냉동기", "1냉온수기", "2냉각탑", "3축열조", "4보일러", "5열교환기",
    "6팽창탱크", "7펌프", "8신재생(지열)", "8신재생(태양광)", "8신재생(연료전지)",
    "11패키지에어컨", "12항온항습기", "13공기조화기", "14팬코일유닛", "15환기",
    "16필터", "17위생", "18급수급탕", "19고저수조", "20오배수", "21오수정화",
    "22물재이용", "23배관", "24덕트", "25보온", "26자동제어", "27방음방진",
]

TABS = [
    ("quantity", "수량 입력", "#4CAF50", "#388E3C"),
    ("copy", "시트 복사", "#2196F3", "#1976D2"),
    ("image", "이미지 삽입", "#FF9800", "#F57C00"),
    ("merge", "엑셀 통합", "#9C27B0", "#7B1FA2"),
    ("pdf", "PDF 변환", "#F44336", "#D32F2F"),
]

HELP_TEXTS = {
    "quantity": (
        "수량 입력 사용방법",
        "1. 설비명과 수량을 입력합니다.\n"
        "2. '엑셀 불러오기' 버튼으로 기존 수량 데이터를 불러올 수 있습니다.\n"
        "3. '설비별 시트 복사 실행' 버튼으로 입력된 수량에 따라 시트를 복사합니다.\n"
        "4. '서울형' 체크박스 선택 시, 서울형 양식 폴더를 사용하고 저장 파일명에 '_서울형'이 붙습니다.",
    ),
    "copy": (
        "시트 복사 사용방법",
        "1. '대상점검및수량.xlsx' 파일을 드래그&드롭하거나 찾아보기로 선택합니다.\n"
        "2. '서울형' 체크박스 선택 시, 'InspectionSheets/서울형 기계설비성능점검보고서_설비별' 폴더의 양식을 사용합니다.\n"
        "3. '설비별 시트 복사 시작' 버튼을 클릭하여 시트 복사를 실행합니다.",
    ),
    "image": (
        "이미지 삽입 사용방법",
        "1. '대상 엑셀' 파일을 드래그&드롭하거나 찾아보기로 선택합니다.\n"
        "2. 각 시트별로 이미지가 저장된 폴더를 지정합니다.\n"
        "3. '실행' 버튼을 클릭하여 엑셀에 이미지를 삽입합니다.\n\n"
        "※ 엑셀 셀에 {{이미지:파일명.jpg}} 형태로 적어 둔 위치에 이미지가 삽입됩니다.",
    ),
    "merge": (
        "엑셀 통합 사용방법",
        "1. 고객사명을 입력합니다.\n"
        "2. 통합할 엑셀 파일들이 있는 폴더를 선택합니다.\n"
        "3. '엑셀 통합 시작' 버튼을 클릭하여 모든 엑셀 파일을 하나의 파일로 통합합니다.",
    ),
    "pdf": (
        "PDF 변환 사용방법",
        "1. 고객사명을 입력합니다.\n"
        "2. PDF로 변환할 엑셀 파일을 드래그&드롭하거나 찾아보기로 선택합니다.\n"
        "3. '고품질' 체크박스 선택 시 표준(고품질), 해제 시 최소(작은 용량) 품질로 저장합니다.\n"
        "4. 'PDF 변환 시작' 버튼을 클릭하여 엑셀 파일을 PDF로 변환합니다.",
    ),
}


# ─────────────────────────── 공용 유틸리티 ───────────────────────────
def load_config() -> dict:
    try:
        if CONFIG_PATH.exists():
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"⚠️ 최근 설정 불러오기 실패: {e}")
    return {}


def save_config(config: dict) -> None:
    try:
        CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        print(f"⚠️ 최근 설정 저장 실패: {e}")


def leading_number(name: str) -> float:
    """파일명 앞의 숫자(정렬 키). 숫자가 없으면 맨 뒤로."""
    m = re.match(r"\d+", name)
    return int(m.group()) if m else float("inf")


def list_xlsx(folder: str | Path, exclude: set[str] | None = None) -> list[str]:
    """폴더 안 .xlsx 파일을 앞자리 숫자 순으로 정렬해 반환 (임시파일·제외 목록 제외)."""
    skip = EXCLUDED_FILES | (exclude or set())
    files = [
        f for f in os.listdir(folder)
        if f.lower().endswith(".xlsx") and not f.startswith("~$") and f not in skip
    ]
    return sorted(files, key=leading_number)


def report_filename(company: str, ext: str) -> str:
    year = datetime.datetime.now().year
    return f"{year}년 {company.strip()} 기계설비 성능점검 결과 보고서.{ext}"


def open_with_default_app(path: str | Path) -> None:
    path = str(path)
    if platform.system() == "Windows":
        os.startfile(path)  # type: ignore[attr-defined]
    elif platform.system() == "Darwin":
        subprocess.run(["open", path], check=False)
    else:
        subprocess.run(["xdg-open", path], check=False)


@contextlib.contextmanager
def excel_session():
    """엑셀 인스턴스를 안전하게 열고, 어떤 경우에도 종료합니다."""
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        yield app
    finally:
        try:
            app.quit()
        except Exception:
            pass


def iter_cell_values(sheet):
    """시트의 사용 범위를 한 번에 읽어 (행, 열, 값)을 반환 (셀 단위 COM 호출보다 훨씬 빠름)."""
    used = sheet.used_range
    values = used.options(ndim=2).value
    row0, col0 = used.row, used.column
    for i, row in enumerate(values):
        for j, value in enumerate(row):
            if value is not None and value != "":
                yield row0 + i, col0 + j, value


# ───────────────────────────── 엑셀 작업 ─────────────────────────────
def get_image_sheet_names(excel_path: str) -> list[str]:
    """이름에 '-'가 들어간 설비 시트 목록 (예: 냉동기-엔진)."""
    pattern = re.compile(r"^.+-\s*.*$")
    with excel_session() as app:
        wb = app.books.open(excel_path, read_only=True)
        names = [s.name for s in wb.sheets if pattern.match(s.name)]
        wb.close()
    return names


def insert_images(excel_path: str, folder_map: dict[str, str], log, out_path: str | None = None) -> tuple[int, int]:
    """{{이미지:파일명}} 표식이 있는 셀(병합셀 포함)에 이미지를 삽입하고 파일을 저장."""
    ok = fail = 0
    with excel_session() as app:
        wb = app.books.open(excel_path)
        for sheet_name, folder in folder_map.items():
            if not folder or not os.path.isdir(folder):
                log(f"[건너뜀] '{sheet_name}' 시트: 폴더 미지정 또는 경로 오류")
                continue
            sheet = wb.sheets[sheet_name]
            log(f"\n🔄 [{sheet_name}] 시트 처리 중... (이미지 폴더: {folder})")
            for row, col, value in iter_cell_values(sheet):
                match = IMAGE_MARKUP.fullmatch(str(value).strip())
                if not match:
                    continue
                img_file = match.group(1)
                img_path = os.path.abspath(os.path.join(folder, img_file))
                cell = sheet.range((row, col))
                addr = cell.get_address(False, False)
                if not os.path.exists(img_path):
                    log(f"  ❌ '{img_file}' 파일을 찾을 수 없습니다. ({addr})")
                    fail += 1
                    continue
                try:
                    area = cell.merge_area
                    sheet.pictures.add(
                        img_path,
                        left=area.left + IMAGE_MARGIN,
                        top=area.top + IMAGE_MARGIN,
                        width=area.width - 2 * IMAGE_MARGIN,
                        height=area.height - 2 * IMAGE_MARGIN,
                    )
                    log(f"  ✓ {img_file} → {addr} 완료")
                    ok += 1
                except Exception as e:
                    log(f"  ❌ {img_file} → {addr} 실패: {e}")
                    fail += 1
        if out_path:
            wb.save(out_path)
        else:
            wb.save()
        wb.close()
    return ok, fail


def read_quantity_file(path: str) -> dict[str, int]:
    """A열=설비명, B열=수량 형식의 엑셀을 읽는다. 숫자가 아닌 행(헤더 등)은 건너뜀."""
    with excel_session() as app:
        wb = app.books.open(path, read_only=True)
        sheet = wb.sheets[0]
        last_row = sheet.used_range.last_cell.row
        values = sheet.range((1, 1), (last_row, 2)).options(ndim=2).value
        wb.close()
    result: dict[str, int] = {}
    for name, qty in values:
        if not name or not qty:
            continue
        try:
            result[str(name).strip()] = int(qty)
        except (ValueError, TypeError):
            continue
    return result


def save_quantity_file(path: str, data: list[tuple[str, int]]) -> None:
    with excel_session() as app:
        wb = app.books.add()
        ws = wb.sheets[0]
        ws.range("A1").value = [["설비명", "수량"], *[[n, q] for n, q in data]]
        wb.save(path)
        wb.close()


def lookup_quantity(stem: str, equip_qty: dict[str, int]) -> int:
    """파일명(확장자 제외)에 해당하는 수량. 정확히 일치 → 공백 제거 후 일치 순으로 탐색."""
    if stem in equip_qty:
        return equip_qty[stem]
    norm = stem.replace(" ", "")
    for key, qty in equip_qty.items():
        if key.replace(" ", "") == norm:
            return qty
    return 0


def copy_sheets_by_quantity(equip_qty: dict[str, int], folder: Path, save_name: str, log) -> Path | None:
    """설비 양식 파일의 시트를 수량만큼 복사해 하나의 통합본으로 저장."""
    files = list_xlsx(folder, exclude={"설비별_점검_통합본.xlsx", "설비별_점검_통합본_서울형.xlsx"})
    if not files:
        log("❌ 점검 설비 파일이 없습니다.")
        return None
    log(f"정렬된 파일 목록: {files}")

    save_path = folder / save_name
    with excel_session() as app:
        new_wb = app.books.add()
        log("✓ 새 통합 파일 생성")
        for fname in files:
            qty = lookup_quantity(Path(fname).stem, equip_qty)
            if qty <= 0:
                log(f"- {fname}: 수량 0, 건너뜀")
                continue
            wb = app.books.open(str(folder / fname), read_only=True)  # 파일은 한 번만 연다
            for i in range(1, qty + 1):
                for sheet in wb.sheets:
                    sheet.api.Copy(After=new_wb.sheets[-1].api)
                suffix = f" (복제 {i}/{qty})" if qty > 1 else ""
                log(f"  ✓ {fname} 시트 복사 완료{suffix}")
            wb.close()
        if len(new_wb.sheets) > 1:
            new_wb.sheets[0].delete()  # 새 통합 파일의 기본 빈 시트 제거
        new_wb.save(str(save_path))
        new_wb.close()
    return save_path


def merge_workbooks(folder: str, output_path: str, log) -> bool:
    """폴더 안 모든 엑셀의 시트를 하나의 파일로 통합."""
    files = list_xlsx(folder, exclude={os.path.basename(output_path)})
    if not files:
        log("  ❌ 통합할 엑셀 파일이 없습니다.")
        return False
    with excel_session() as app:
        new_wb = app.books.add()
        for fname in files:
            log(f"  - {fname} 처리 중...")
            wb = app.books.open(os.path.join(folder, fname), read_only=True)
            for sheet in wb.sheets:
                sheet.api.Copy(After=new_wb.sheets[-1].api)
            wb.close()
        if len(new_wb.sheets) > 1:
            new_wb.sheets[0].delete()
        new_wb.save(output_path)
        new_wb.close()
    return True


def export_pdf(excel_path: str, pdf_path: str, high_quality: bool) -> None:
    with excel_session() as app:
        wb = app.books.open(excel_path, read_only=True)
        wb.api.ExportAsFixedFormat(
            0,  # xlTypePDF
            pdf_path,
            Quality=0 if high_quality else 1,  # 0: 표준(고품질), 1: 최소
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
        wb.close()


# ───────────────────────────── 디자인 시스템 ─────────────────────────────
from tkinter import ttk  # noqa: E402

try:
    from PIL import Image, ImageOps, ImageTk
    PIL_OK = True
    PIL_ERR = ""
except ImportError as _e:  # Pillow 없으면 미리보기를 만들 수 없음
    PIL_OK = False
    PIL_ERR = str(_e)

# ───────── v14 신규: 사진 연결(미리보기 방식)용 상수·엑셀 작업 ─────────
SLOT_MARKUP = re.compile(r"\{\{\s*이미지\s*(?::\s*([^}]*?))?\s*\}\}")  # {{이미지}} 또는 {{이미지:설명}}
IMG_EXT = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff")  # 엑셀이 지원하는 형식
THUMB = 160  # 미리 만들어 두는 최대 크기 (화면에는 선택한 크기로 표시)
SIZES = {"작게": 72, "보통": 96, "크게": 150}
BIG = 720  # 마우스를 올렸거나 클릭했을 때 보이는 큰 미리보기 크기 (화면이 작으면 자동으로 줄어듦)
SCROLL_CANVASES: list = []


def natural_key(text: str) -> list:
    """파일명 자연 정렬 (img2 < img10)"""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", text)]


def cell_address(row: int, col: int) -> str:
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"


def com_error_text(e: Exception) -> str:
    """엑셀(COM) 오류에서 읽을 수 있는 설명만 꺼낸다."""
    try:
        detail = e.args[2]
        if isinstance(detail, tuple) and detail[2]:
            return str(detail[2])
    except Exception:
        pass
    return str(e)


PIC_PREFIX = "ximgpro|"  # 이 프로그램이 넣은 사진의 이름표: ximgpro|행|열  (또는 ximgpro|행|열|왼쪽,위,가로,세로)


class Box:
    """사진이 들어갈 영역(pt). 기존 사진의 위치·크기를 그대로 쓸 때 사용."""

    def __init__(self, left, top, width, height):
        self.left, self.top, self.width, self.height = left, top, width, height


def _new_slot(kind: str, row: int, col: int, label: str = "", marker: str = "{{이미지}}", geom=None) -> dict:
    return {"kind": kind, "row": row, "col": col, "label": label, "marker": marker, "geom": geom,
            "old": [], "existing": False, "old_pil": None}


def _area_bounds(sheet, row: int, col: int) -> tuple:
    a = sheet.range((row, col)).merge_area
    return a.left, a.top, a.left + a.width, a.top + a.height


def _shape_is_picture(shp) -> bool:
    try:
        return int(shp.Type) in (11, 13)  # 13: 그림, 11: 연결된 그림
    except Exception:
        return False


def _export_shape_thumb(sheet, shp, tmpdir: str, n: int, log=None, context: str = ""):
    """엑셀 기존 그림을 PNG 미리보기로 변환한다.

    우선 Excel CopyPicture 결과를 Windows 클립보드에서 Pillow로 직접 읽는다.
    이 경로는 Chart.Paste/Chart.Export COM 자동화에 의존하지 않는다.
    클립보드 직접 읽기가 안 되는 환경에서는 기존 차트 내보내기를 보조 경로로 사용한다.
    """
    if not PIL_OK:
        return None
    import time
    chart_obj = None
    errors = []

    try:
        try:
            sheet.activate()
        except Exception:
            try:
                sheet.api.Activate()
            except Exception as exc:
                errors.append(f"시트 활성화: {exc}")

        # CopyPicture(xlScreen, xlBitmap) → Windows 클립보드 → Pillow
        # Excel은 화면 렌더링 상태에 따라 클립보드 반영이 늦을 수 있어 반복 확인한다.
        try:
            from PIL import ImageGrab
            for attempt in range(4):
                try:
                    shp.CopyPicture(Appearance=1, Format=2)
                    time.sleep(0.25 + attempt * 0.15)
                    clip = ImageGrab.grabclipboard()
                    if isinstance(clip, Image.Image):
                        clip.load()
                        result = clip.convert("RGB")
                        result.thumbnail((600, 600))
                        return result
                    if isinstance(clip, list):
                        errors.append("클립보드에 이미지 대신 파일 목록이 들어왔습니다.")
                    else:
                        errors.append(f"클립보드 이미지 없음(시도 {attempt + 1}/4)")
                except Exception as exc:
                    errors.append(f"클립보드 복사 시도 {attempt + 1}: {exc}")
        except Exception as exc:
            errors.append(f"클립보드 직접 읽기 실패: {exc}")

        # 보조 경로: 일부 Excel/Pillow 조합에서는 클립보드 직접 읽기가 제한됨.
        try:
            chart_obj = sheet.api.ChartObjects().Add(
                0, 0, max(float(shp.Width), 2.0), max(float(shp.Height), 2.0)
            )
            chart = chart_obj.Chart
            last_error = None
            for _ in range(4):
                try:
                    shp.CopyPicture(Appearance=1, Format=2)
                    time.sleep(0.35)
                    chart.Paste()
                    last_error = None
                    break
                except Exception as exc:
                    last_error = exc
                    time.sleep(0.4)
            if last_error is not None:
                raise RuntimeError(f"Chart.Paste 실패: {last_error}")

            path = os.path.join(tmpdir, f"preview_{n}.png")
            try:
                exported = chart.Export(Filename=path, FilterName="PNG")
            except TypeError:
                exported = chart.Export(path)
            if exported is False or not os.path.isfile(path) or os.path.getsize(path) == 0:
                raise RuntimeError("엑셀이 미리보기 PNG를 생성하지 못했습니다.")
            with Image.open(path) as image:
                image.load()
                result = image.convert("RGB")
            result.thumbnail((600, 600))
            return result
        except Exception as exc:
            errors.append(f"차트 보조 변환 실패: {exc}")

        raise RuntimeError(" | ".join(errors[-6:]))
    except Exception as exc:
        if log:
            prefix = f" [{context}]" if context else ""
            log(f"  ⚠ 미리보기 변환 단계 실패{prefix}: {exc}")
        return None
    finally:
        if chart_obj is not None:
            try:
                chart_obj.Delete()
            except Exception:
                pass


def _sheet_print_bounds(sheet) -> list[tuple[float, float, float, float]]:
    """인쇄 영역을 포인트 좌표로 반환. 인쇄 영역 미지정 시 사용 범위를 사용한다."""
    try:
        print_area = str(sheet.api.PageSetup.PrintArea or "").strip()
    except Exception:
        print_area = ""
    result = []
    if print_area:
        for part in print_area.split(","):
            try:
                area = sheet.range(part.strip())
                left, top = float(area.left), float(area.top)
                result.append((left, top, left + float(area.width), top + float(area.height)))
            except Exception:
                continue
    if not result:
        try:
            area = sheet.used_range
            left, top = float(area.left), float(area.top)
            result.append((left, top, left + float(area.width), top + float(area.height)))
        except Exception:
            result.append((-1e12, -1e12, 1e12, 1e12))
    return result


def _rect_intersects(a: tuple, b: tuple) -> bool:
    """두 사각형이 면적으로 겹치는지 확인한다."""
    return a[0] < b[2] and a[2] > b[0] and a[1] < b[3] and a[3] > b[1]


def scan_image_slots(excel_path: str, log=None) -> tuple[list[dict], list]:
    """인쇄 영역 안의 {{이미지}} 표식과 기존 사진을 찾아 미리보기를 만들고,
    문서 전체의 시트 인쇄 미리보기(PDF 렌더링 이미지 목록)도 함께 생성한다."""
    import shutil
    import tempfile
    log = log or (lambda m: None)
    slots: list[dict] = []
    previews: list = []
    tmpdir = tempfile.mkdtemp(prefix="ximgpro_")
    thumbs = excluded_shapes = preview_failures = 0
    pdf_tmp = os.path.join(tmpdir, "doc_preview.pdf")
    try:
        with excel_session() as app:
            # 엑셀 창이 화면에 깜빡이며 열리지 않도록 백그라운드 숨김 모드로 실행
            app.screen_updating = False
            app.visible = False
            wb = app.books.open(excel_path, read_only=False, update_links=False)
            try:
                for sheet in wb.sheets:
                    markers: dict[tuple, dict] = {}
                    bounds: dict[tuple, tuple] = {}
                    standalone: list[dict] = []
                    print_bounds = _sheet_print_bounds(sheet)

                    # 인쇄 영역 밖의 {{이미지}} 표식 칸은 제외
                    for row, col, value in iter_cell_values(sheet):
                        text = str(value).strip()
                        match = SLOT_MARKUP.fullmatch(text)
                        if not match or (row == 1 and col == 1):
                            continue  # A1은 보고서 회사 로고 영역이므로 사진 대상에서 제외
                        cell_bounds = _area_bounds(sheet, row, col)
                        if not any(_rect_intersects(cell_bounds, pb) for pb in print_bounds):
                            continue
                        markers[(row, col)] = _new_slot(
                            "marker", row, col, (match.group(1) or "").strip(), text
                        )
                        bounds[(row, col)] = cell_bounds

                    shapes = sheet.api.Shapes
                    for k in range(1, shapes.Count + 1):
                        shp = shapes.Item(k)
                        if not _shape_is_picture(shp):
                            continue
                        try:
                            name = str(shp.Name)
                            left, top = float(shp.Left), float(shp.Top)
                            width, height = float(shp.Width), float(shp.Height)
                            top_left = shp.TopLeftCell
                            row0, col0 = int(top_left.Row), int(top_left.Column)
                            alt = str(shp.AlternativeText or "").strip()
                            if row0 == 1 and col0 == 1:
                                excluded_shapes += 1
                                continue  # A1 회사 로고는 교체 대상에서 제외
                        except Exception:
                            continue

                        shape_bounds = (left, top, left + width, top + height)
                        if not any(_rect_intersects(shape_bounds, pb) for pb in print_bounds):
                            excluded_shapes += 1
                            continue

                        ref = {"name": name, "left": left, "top": top}
                        if name.startswith(PIC_PREFIX):
                            parts = name.split("|")
                            try:
                                key = (int(parts[1]), int(parts[2]))
                            except (ValueError, IndexError):
                                key = (row0, col0)
                            box = None
                            if len(parts) >= 4:
                                try:
                                    values = tuple(float(v) for v in parts[3].split(","))
                                    box = values if len(values) == 4 else None
                                except ValueError:
                                    pass
                            match = SLOT_MARKUP.fullmatch(alt)
                            label = (match.group(1) or "").strip() if match else ""
                            if box is not None:
                                entry = _new_slot("pic", key[0], key[1], label, "", box)
                                standalone.append(entry)
                            else:
                                entry = markers.get(key)
                                if entry is None:
                                    try:
                                        cell_bounds = _area_bounds(sheet, key[0], key[1])
                                    except Exception:
                                        continue
                                    if not any(_rect_intersects(cell_bounds, pb) for pb in print_bounds):
                                        continue
                                    entry = markers[key] = _new_slot(
                                        "marker", key[0], key[1], label,
                                        alt if match else "{{이미지}}"
                                    )
                                    bounds[key] = cell_bounds
                            entry["existing"] = True
                            entry["old"].append(ref)
                            entry.setdefault("_shp", shp)
                            continue

                        if width < 25 or height < 25:
                            continue
                        center_x, center_y = left + width / 2, top + height / 2
                        target = next(
                            (key for key, (l, t, r, bottom) in bounds.items()
                             if l <= center_x <= r and t <= center_y <= bottom),
                            None
                        )
                        if target is not None:
                            entry = markers[target]
                        else:
                            entry = _new_slot(
                                "pic", row0, col0, "", "",
                                (left, top, width, height)
                            )
                            standalone.append(entry)
                        entry["existing"] = True
                        entry["old"].append(ref)
                        entry.setdefault("_shp", shp)

                    found = list(markers.values()) + standalone
                    found.sort(key=lambda item: (item["row"], item["col"]))
                    for entry in found:
                        shape = entry.pop("_shp", None)
                        if shape is not None and thumbs < 200:
                            if thumbs == 0:
                                log("기존 사진 미리보기를 만드는 중입니다...")
                            entry["old_pil"] = _export_shape_thumb(
                                sheet, shape, tmpdir, thumbs, log=log,
                                context=f"{sheet.name}!{cell_address(entry['row'], entry['col'])}"
                            )
                            if entry["old_pil"] is None:
                                preview_failures += 1
                            thumbs += 1
                        entry["sheet"] = sheet.name
                        entry["addr"] = cell_address(entry["row"], entry["col"])
                        slots.append(entry)

                # 시트 미리보기용 임시 PDF 생성
                try:
                    wb.api.ExportAsFixedFormat(0, pdf_tmp, Quality=1, IncludeDocProperties=False, IgnorePrintAreas=False)
                except Exception as pe:
                    if log:
                        log(f"  ⚠ 미리보기 PDF 내보내기 생략: {pe}")
            finally:
                wb.close()

        # PDF 페이지들을 PIL 이미지로 렌더링
        if os.path.exists(pdf_tmp):
            try:
                import fitz
                doc = fitz.open(pdf_tmp)
                for page_num in range(min(len(doc), 40)):
                    page = doc[page_num]
                    pix = page.get_pixmap(dpi=110)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    previews.append(img)
                doc.close()
                if log:
                    log(f"✓ 엑셀 시트 미리보기 {len(previews)}페이지 준비 완료")
            except Exception as fe:
                if log:
                    log(f"  ⚠ 미리보기 이미지 변환: {fe}")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    if excluded_shapes:
        log(f"인쇄 범위 밖 사진 {excluded_shapes}개는 제외했습니다.")
    if preview_failures:
        log(f"⚠ 기존 사진 {preview_failures}개의 미리보기 변환 실패 — 위 항목별 오류를 확인하세요.")
    if thumbs >= 200:
        log("⚠ 미리보기는 성능을 위해 최대 200개까지만 생성합니다.")
    return slots, previews


def _find_shape(sheet, name: str, left: float, top: float):
    """기존 그림을 찾는다. 위치가 미세하게 바뀐 경우에도 고유한 도형 이름으로 재확인한다."""
    shapes = sheet.api.Shapes
    named = []
    for k in range(1, shapes.Count + 1):
        shp = shapes.Item(k)
        if str(shp.Name) == name:
            named.append(shp)
            if abs(float(shp.Left) - left) < 1.0 and abs(float(shp.Top) - top) < 1.0:
                return shp
    # Excel 도형 이름은 시트 내에서 고유하므로 위치 오차가 큰 경우 이름으로 보완 검색
    return named[0] if len(named) == 1 else None


def _add_shape(sheet, path: str, area, fit: bool, margin: float, check_small: bool):
    """사진을 원본 크기로 넣은 뒤 영역에 맞게 위치·크기를 조절한다. (도형, 가로, 세로, 확대배치 여부) 반환"""
    shp = sheet.api.Shapes.AddPicture(path, 0, -1, area.left, area.top, -1, -1)  # 파일 연결 X, 문서에 포함 O, 원본 크기
    try:
        shp.LockAspectRatio = 0
        nat_w, nat_h = float(shp.Width), float(shp.Height)
        avail_w, avail_h = area.width - 2 * margin, area.height - 2 * margin
        small = check_small and (avail_w < 30 or avail_h < 30)  # 칸이 너무 작거나 숨겨진 행/열이면 사진이 안 보이므로 확대 배치
        if small:
            scale = min(1.0, 200.0 / nat_w)
            w, h, left, top = nat_w * scale, nat_h * scale, area.left, area.top
        elif fit:
            scale = min(avail_w / nat_w, avail_h / nat_h)
            w, h = nat_w * scale, nat_h * scale
            left, top = area.left + margin + (avail_w - w) / 2, area.top + margin + (avail_h - h) / 2
        else:
            w, h, left, top = avail_w, avail_h, area.left + margin, area.top + margin
        shp.Left, shp.Top, shp.Width, shp.Height = left, top, w, h
        shp.Placement = 1  # 셀과 함께 이동하고 크기 변경
        shp.Visible = -1
    except Exception:
        shp.Delete()
        raise
    return shp, w, h, small


def insert_images_by_slots(excel_path: str, items: list[tuple], out_path: str, log, fit: bool = True, delete_slots: list[dict] = None) -> tuple[int, int, str]:
    """(자리 정보, 사진경로) 목록대로 사진을 넣는다. 기존 사진이 있는 자리는 같은 위치·크기로 교체한다.
    delete_slots 목록의 슬롯은 기존 사진만 삭제한다.
    out_path 가 원본과 다르면 원본은 바꾸지 않고 새 파일로 저장한다."""
    ok = fail = 0
    delete_slots = delete_slots or []
    in_place = os.path.normcase(os.path.abspath(excel_path)) == os.path.normcase(os.path.abspath(out_path))
    with excel_session() as app:
        wb = app.books.open(excel_path, read_only=not in_place)

        # 1. 기존 사진 단독 삭제 대상 처리
        for slot in delete_slots:
            row, col = slot["row"], slot["col"]
            sheet = wb.sheets[slot["sheet"]]
            addr = f'{slot["sheet"]}!{cell_address(row, col)}'
            if row == 1 and col == 1:
                continue
            old_objs = [o for o in (_find_shape(sheet, r["name"], r["left"], r["top"]) for r in slot["old"]) if o is not None]
            for old_obj in old_objs:
                try:
                    old_obj.Delete()
                except Exception:
                    pass
            # 원래 셀의 표식(예: {{이미지}})이 있었다면 복원
            if slot.get("marker"):
                try:
                    sheet.range((row, col)).value = slot["marker"]
                except Exception:
                    pass
            log(f"🗑 {addr}: 기존 사진 삭제 완료")
            ok += 1

        # 2. 새 사진 삽입 및 교체 처리
        for slot, img_path in items:
            row, col = slot["row"], slot["col"]
            sheet = wb.sheets[slot["sheet"]]
            addr = f'{slot["sheet"]}!{cell_address(row, col)}'
            img_path = os.path.normpath(os.path.abspath(img_path))  # 엑셀은 '/'가 섞인 경로를 읽지 못하는 경우가 있음
            name = os.path.basename(img_path)
            if not os.path.exists(img_path):
                log(f"❌ 사진 파일을 찾을 수 없습니다: {img_path} ({addr})")
                fail += 1
                continue
            tmp = None
            try:
                if slot["kind"] == "pic":  # 기존 사진 자리: 현재 위치·크기를 그대로 사용
                    area, margin, check_small = Box(*slot["geom"]), 0, False
                else:  # {{이미지}} 칸: 칸(병합 포함) 크기를 사용
                    area, margin, check_small = sheet.range((row, col)).merge_area, IMAGE_MARGIN, True
                # 기존 사진은 새 사진을 넣기 전에 먼저 제거한다. 이름과 위치를 함께 확인해
                # 해당 슬롯의 그림만 지우며, 다른 그림과 A1 로고는 건드리지 않는다.
                old_objs = [o for o in (_find_shape(sheet, r["name"], r["left"], r["top"]) for r in slot["old"]) if o is not None]
                if row == 1 and col == 1:
                    log(f"⚠ {addr}: A1 회사 로고 영역은 보호되어 건너뜁니다.")
                    fail += 1
                    continue
                for old_obj in old_objs:
                    old_obj.Delete()
                try:
                    shp, w, h, small = _add_shape(sheet, img_path, area, fit, margin, check_small)
                except Exception:
                    # 경로 문제에 대비: 임시 폴더에 영문 이름으로 복사해서 한 번 더 시도
                    import shutil
                    import tempfile
                    tmp = os.path.join(tempfile.gettempdir(), f"ximgpro_{row}_{col}{os.path.splitext(img_path)[1].lower()}")
                    shutil.copyfile(img_path, tmp)
                    shp, w, h, small = _add_shape(sheet, tmp, area, fit, margin, check_small)
                if slot["kind"] == "pic":
                    box = ",".join(f"{v:.2f}" for v in slot["geom"])
                    shp.Name = f"{PIC_PREFIX}{row}|{col}|{box}"
                    shp.AlternativeText = ""
                else:
                    shp.Name = f"{PIC_PREFIX}{row}|{col}"
                    shp.AlternativeText = slot["marker"]
                    sheet.range((row, col)).value = None  # 표식 글자 지우기
                extra = (" · 기존 사진 교체" if old_objs else "") + (" · ⚠ 칸이 작아 확대 배치" if small else "")
                log(f"✓ {name} → {addr} (가로 {w:.0f} × 세로 {h:.0f}pt{extra})")
                ok += 1
            except Exception as e:
                log(f"❌ {name} → {addr} 실패: {com_error_text(e)}")
                fail += 1
            finally:
                if tmp and os.path.exists(tmp):
                    os.remove(tmp)
        if in_place:
            wb.save()
        else:
            wb.save(out_path)
        wb.close()
    return ok, fail, out_path


def insert_and_preview(excel_path: str, items: list[tuple], log, fit: bool = True, delete_slots: list[dict] = None) -> tuple[int, int, str, list]:
    """사진을 임시 작업 파일에 삽입/삭제하고, 삽입된 상태의 최신 인쇄 미리보기를 생성해 반환."""
    import tempfile
    tmpdir = tempfile.mkdtemp(prefix="ximg_staged_")
    ext = os.path.splitext(excel_path)[1]
    staged_excel = os.path.join(tmpdir, f"staged{ext}")
    staged_pdf = os.path.join(tmpdir, "staged.pdf")

    # 1. 임시 파일에 사진 삽입 및 삭제 실행
    ok, fail, _ = insert_images_by_slots(excel_path, items, staged_excel, log, fit, delete_slots=delete_slots)

    # 2. 사진이 삽입된 엑셀에서 PDF 생성 및 미리보기 추출
    previews = []
    if os.path.exists(staged_excel):
        try:
            with excel_session() as app:
                wb = app.books.open(staged_excel, read_only=True)
                try:
                    wb.api.ExportAsFixedFormat(0, staged_pdf, Quality=1, IncludeDocProperties=False, IgnorePrintAreas=False)
                finally:
                    wb.close()
            if os.path.exists(staged_pdf):
                import fitz
                doc = fitz.open(staged_pdf)
                for page_num in range(min(len(doc), 40)):
                    page = doc[page_num]
                    pix = page.get_pixmap(dpi=110)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    previews.append(img)
                doc.close()
                if log:
                    log(f"✓ 삽입된 엑셀 미리보기 {len(previews)}페이지 준비 완료")
        except Exception as pe:
            if log:
                log(f"  ⚠ 삽입 후 미리보기 생성 실패: {pe}")

    return ok, fail, staged_excel, previews


C = {
    "side": "#1B2440", "side_hi": "#27325A", "bg": "#F4F6FA", "card": "#FFFFFF",
    "line": "#E3E7EF", "text": "#1F2937", "mut": "#6B7280", "primary": "#4F6BED",
    "ok": "#16A34A", "bad": "#DC2626", "field": "#F9FAFC",
}
NAV_ICONS = {"quantity": "🔢", "copy": "📑", "image": "🖼", "merge": "🧩", "pdf": "📄"}
PAGE_INFO = {
    "quantity": ("수량 입력", "설비별 점검 수량을 입력하고 바로 시트를 복사합니다."),
    "copy": ("시트 복사", "수량표 엑셀 파일을 기준으로 설비별 양식을 복사합니다."),
    "image": ("이미지 삽입", "엑셀 안의 {{이미지:파일명}} 표식 위치에 사진을 넣습니다."),
    "merge": ("엑셀 통합", "폴더 안의 엑셀 파일을 하나의 보고서 파일로 합칩니다."),
    "pdf": ("PDF 변환", "엑셀 파일을 제출용 PDF로 저장합니다."),
}


TABS = [("image", "이미지 삽입", "#FF9800", "#F57C00"), ("link", "사진 연결", "#14B8A6", "#0F9488")]
NAV_ICONS["link"] = "🔗"
NAV_ICONS["env"] = "🩺"
PAGE_INFO["env"] = ("환경 점검", "다른 컴퓨터에서도 잘 실행되는지 확인하고, 부족한 항목을 안내합니다.")
HELP_TEXTS["env"] = (
    "환경 점검 사용방법",
    "프로그램이 동작하는 데 필요한 항목(Windows, Microsoft Excel, 구성요소, 양식 폴더)을 점검합니다.\n"
    "❌ 는 반드시 해결해야 하는 항목, ⚠ 는 일부 기능에만 필요한 항목입니다.\n"
    "[엑셀 연결 테스트]로 엑셀이 실제로 실행되는지 확인하고, [결과 복사]로 점검 결과를 복사해 전달할 수 있습니다.",
)
PAGE_INFO["link"] = ("사진 연결 (미리보기)", "사진을 보면서 엑셀의 칸에 직접 연결합니다. 파일 이름을 바꿀 필요가 없어요.")
HELP_TEXTS["link"] = (
    "사진 연결 사용방법",
    "1. 대상 엑셀을 고르면 {{이미지}} 또는 {{이미지:설명}} 표식이 있는 칸이 왼쪽 목록에 나타납니다.\n"
    "2. 사진 폴더를 고르면 오른쪽에 사진 미리보기가 나타납니다.\n"
    "   사진 위에 마우스를 올리면 큰 사진이 보이고, 사진 미리보기는 기본 보통 크기로 표시됩니다.\n"
    "3. 왼쪽에서 칸을 클릭한 뒤 오른쪽 사진을 클릭하면 연결되고, 다음 빈 칸이 자동으로 선택됩니다.\n"
    "   사진을 칸 위로 끌어다 놓아도 됩니다.\n"
    "4. '순서대로 자동 채우기'로 빈 칸을 한 번에 채운 뒤 틀린 곳만 고칠 수 있습니다.\n"
    "5. '삽입 실행'을 누르면 원본은 그대로 두고 '파일명_사진삽입.xlsx'가 새로 만들어집니다.\n\n"
    "※ 사진 폴더에 사진을 추가했다면 ‘새로고침’(F5)을 누르세요. 연결해 둔 내용은 그대로 유지되고, 새 사진에는 NEW가 표시됩니다.\n"
    "※ 엑셀에 이미 들어 있는 사진(직접 붙여 넣은 사진 포함)은 ‘기존 사진 있음’으로 표시되고 현재 사진이 미리보기로 보입니다.\n"
    "   새 사진을 연결하면 해당 슬롯의 기존 사진을 먼저 삭제한 뒤 새 사진을 삽입합니다. A1 회사 로고는 제외됩니다.\n"
    "※ ‘사진 비율 유지’를 끄면 사진을 칸 크기에 꽉 채워 넣습니다.\n"
    "※ 엑셀 파일과 사진 폴더는 창에 끌어다 놓아도 됩니다.\n"
    "※ 아이폰 HEIC 사진은 엑셀이 지원하지 않아 목록에 나타나지 않습니다.",
)


def shade(hex_color: str, factor: float) -> str:
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    if factor < 0:
        r, g, b = (int(v * (1 + factor)) for v in (r, g, b))
    else:
        r, g, b = (int(v + (255 - v) * factor) for v in (r, g, b))
    return f"#{r:02x}{g:02x}{b:02x}"


def round_rect(cv: tk.Canvas, x1, y1, x2, y2, r, **kw):
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return cv.create_polygon(pts, smooth=True, **kw)


def lbl(parent, text, size=10, bold=False, color=None, bg=None, **kw):
    return tk.Label(parent, text=text, font=(FONT, size, "bold" if bold else "normal"),
                    fg=color or C["text"], bg=bg or parent.cget("bg"), **kw)


def field_entry(parent, var, width=None, justify="left"):
    return tk.Entry(parent, textvariable=var, width=width or 20, justify=justify, relief="flat",
                    bg=C["field"], fg=C["text"], font=(FONT, 10), insertbackground=C["text"],
                    highlightthickness=1, highlightbackground=C["line"], highlightcolor=C["primary"])


def check(parent, text, var):
    bg = parent.cget("bg")
    return tk.Checkbutton(parent, text=text, variable=var, bg=bg, activebackground=bg, fg=C["text"],
                          font=(FONT, 10), relief="flat", cursor="hand2", selectcolor="#FFFFFF")


def card(parent) -> tk.Frame:
    return tk.Frame(parent, bg=C["card"], highlightthickness=1, highlightbackground=C["line"])


class RoundButton(tk.Canvas):
    """둥근 모서리 + 마우스 올림 효과가 있는 버튼"""
    STYLES = {"primary": (C["primary"], "#FFFFFF"), "ghost": ("#EEF1FB", C["primary"]),
              "danger": ("#FDECEC", C["bad"]), "success": (C["ok"], "#FFFFFF")}

    def __init__(self, master, text, command, kind="primary", width=None, height=34):
        fill, fg = self.STYLES[kind]
        font = (FONT, 10, "bold")
        w = width or tkfont.Font(font=font).measure(text) + 34
        super().__init__(master, width=w, height=height, bg=master.cget("bg"), highlightthickness=0, bd=0, cursor="hand2")
        hover = shade(fill, -0.08)
        shape = round_rect(self, 1, 1, w - 1, height - 1, 8, fill=fill, outline=fill)
        self._text_id = self.create_text(w / 2, height / 2, text=text, fill=fg, font=font)
        self.bind("<Enter>", lambda e: self.itemconfigure(shape, fill=hover, outline=hover))
        self.bind("<Leave>", lambda e: self.itemconfigure(shape, fill=fill, outline=fill))
        self.bind("<ButtonRelease-1>", lambda e: command())

    def set_text(self, text: str):
        self.itemconfigure(self._text_id, text=text)


class DropCard(tk.Canvas):
    """점선 테두리 드래그&드롭 영역 + 찾아보기 버튼"""

    def __init__(self, master, title: str, on_select=None, width=420, height=150):
        bg = master.cget("bg")
        super().__init__(master, width=width, height=height, bg=bg, highlightthickness=0, bd=0)
        self.title_text, self.on_select, self.path = title, on_select, tk.StringVar()
        round_rect(self, 2, 2, width - 2, height - 2, 14, fill="#FAFBFF", outline="#B9C4F3", dash=(6, 4), width=2)
        self.create_text(width / 2, 30, text="📄", font=(FONT, 20))
        self._main = self.create_text(width / 2, 62, text=f"{title} 파일을 여기에 끌어다 놓으세요",
                                      font=(FONT, 10, "bold"), fill=C["text"])
        self._sub = self.create_text(width / 2, 84, text="또는 아래 버튼으로 선택", font=(FONT, 9), fill=C["mut"])
        browse = RoundButton(self, "찾아보기", self._browse, kind="ghost", height=30)
        self.create_window(width / 2, height - 24, window=browse)
        self.path.trace_add("write", lambda *a: self._refresh())
        for w in (self, browse):
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<Drop>>", self._on_drop)

    def _refresh(self):
        p = self.path.get()
        if p:
            name = os.path.basename(p)
            name = name if len(name) <= 46 else name[:43] + "..."
            self.itemconfigure(self._main, text=f"✓ {name}", fill=C["ok"])
            self.itemconfigure(self._sub, text="다른 파일을 끌어다 놓으면 바꿀 수 있어요")
        else:
            self.itemconfigure(self._main, text=f"{self.title_text} 파일을 여기에 끌어다 놓으세요", fill=C["text"])
            self.itemconfigure(self._sub, text="또는 아래 버튼으로 선택")

    def _browse(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if path:
            self.set_path(path)

    def _on_drop(self, event):
        items = self.tk.splitlist(event.data)
        path = items[0] if items else ""
        if os.path.isfile(path):
            self.set_path(path)
        else:
            messagebox.showwarning("파일 형식 오류", f"올바른 {self.title_text} 파일을 선택해주세요.")

    def set_path(self, path: str):
        self.path.set(path)
        if self.on_select:
            self.on_select(path)


def make_scroll_area(parent, bg) -> tuple[tk.Canvas, tk.Frame]:
    canvas = tk.Canvas(parent, bg=bg, highlightthickness=0, bd=0, height=60)
    bar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
    inner = tk.Frame(canvas, bg=bg)
    win = canvas.create_window((0, 0), window=inner, anchor="nw")
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win, width=e.width))
    canvas.configure(yscrollcommand=bar.set)
    bar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    SCROLL_CANVASES.append(canvas)
    return canvas, inner


# ───────── v16 신규: 환경 점검 ─────────
def check_environment() -> list[dict]:
    """다른 컴퓨터에서도 실행되는지 항목별로 점검한다. status: ok / warn / bad"""
    frozen = getattr(sys, "frozen", False)
    is_win = platform.system() == "Windows"
    items: list[dict] = []

    def add(name, status, detail, fix=""):
        items.append({"name": name, "status": status, "detail": detail, "fix": fix})

    add("실행 방식", "ok", "exe (필요한 라이브러리가 모두 포함됨)" if frozen else f"파이썬 {platform.python_version()}")
    add("Windows", "ok" if is_win else "bad",
        platform.platform() if is_win else f"{platform.system()} 에서는 실행할 수 없습니다",
        "" if is_win else "이 프로그램은 Windows 전용입니다.")
    if is_win:
        try:
            import winreg
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CurVer") as k:
                    ver = winreg.QueryValue(k, None)
            except OSError:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"):
                    ver = "버전 확인 불가"
            add("Microsoft Excel", "ok", f"설치됨 ({ver})")
        except OSError:
            add("Microsoft Excel", "bad", "설치되어 있지 않거나 찾을 수 없습니다",
                "Microsoft Excel 2013 이상을 설치하세요. (웹 엑셀, 한컴셀, LibreOffice, 맥용 엑셀은 사용할 수 없습니다)")
    else:
        add("Microsoft Excel", "warn", "확인할 수 없습니다 (Windows가 아님)")
    source = "exe에 포함됨" if frozen else "설치됨"
    add("xlwings (엑셀 제어)", "ok", source)
    add("tkinterdnd2 (끌어다 놓기)", "ok", source)
    try:
        import pythoncom  # noqa: F401
        add("pywin32 (엑셀 연결)", "ok", source)
    except ImportError:
        add("pywin32 (엑셀 연결)", "bad" if is_win else "warn", "없음",
            "exe를 다시 만들어 주세요 (build_exe.py)." if frozen else "터미널에서 pip install pywin32")
    if PIL_OK:
        add("Pillow (사진 미리보기)", "ok", source)
    else:
        add("Pillow (사진 미리보기)", "warn", "없음 — 사진 연결 탭의 미리보기를 쓸 수 없습니다",
            "exe를 만드는 컴퓨터에서 build_exe.py를 다시 실행해 새 exe로 교체하세요. (사용하는 컴퓨터에서 설치할 필요는 없습니다)"
            if frozen else "사진 연결 탭의 [Pillow 설치] 버튼을 누르거나, 터미널에서 pip install pillow")

    n = len(list(TEMPLATE_DIR.glob("*.xlsx"))) if TEMPLATE_DIR.is_dir() else 0
    add("InspectionSheets 폴더", "ok" if n else "warn", f"{TEMPLATE_DIR}  (양식 {n}개)",
        "" if n else "프로그램(exe)과 같은 폴더에 InspectionSheets 폴더를 함께 복사하세요. (시트 복사 기능에 필요)")
    sn = len(list(SEOUL_TEMPLATE_DIR.glob("*.xlsx"))) if SEOUL_TEMPLATE_DIR.is_dir() else 0
    add("서울형 양식 폴더", "ok" if sn else "warn", f"양식 {sn}개" if sn else "없음",
        "" if sn else "서울형 양식을 쓸 때만 필요합니다. (InspectionSheets 안의 '서울형 기계설비성능점검보고서_설비별')")
    try:
        probe = BASE_DIR / ".write_test"
        probe.write_text("x")
        probe.unlink()
        add("폴더 쓰기 권한", "ok", str(BASE_DIR))
    except OSError:
        add("폴더 쓰기 권한", "warn", f"{BASE_DIR} 에 파일을 만들 수 없습니다",
            "'Program Files' 같은 보호된 폴더 대신 바탕화면이나 문서 폴더로 옮겨서 실행하세요.")
    return items


def test_excel_connection() -> str:
    """엑셀을 실제로 한 번 실행해 보고 버전을 돌려준다."""
    with excel_session() as app:
        return str(app.version)


ENV_GUIDE = [
    ("h", "■ 다른 컴퓨터에서 사용하는 방법"),
    ("", "1. 아래 두 가지를 같은 폴더에 넣어 통째로 복사합니다.\n"
         "     · ximgpro exe 파일\n"
         "     · InspectionSheets 폴더 (설비별 양식 엑셀)\n"
         "   ※ 파이썬, xlwings, Pillow 같은 프로그램은 따로 설치하지 않아도 됩니다. exe 안에 이미 들어 있습니다.\n"
         "2. 사용할 컴퓨터에 Microsoft Excel 2013 이상이 설치되어 있어야 합니다.\n"
         "   (웹 엑셀, 한컴셀, LibreOffice, 맥용 엑셀은 사용할 수 없습니다.)\n"
         "3. exe를 더블클릭해서 실행하고, 이 탭에서 모든 항목이 ✓ 인지 확인합니다.\n"
         "4. [엑셀 연결 테스트] 버튼으로 엑셀이 실제로 실행되는지 확인할 수 있습니다."),
    ("h", "■ 처음 실행할 때 파란 경고창이 뜨는 경우 (Windows SmartScreen)"),
    ("", "‘추가 정보’를 누른 뒤 ‘실행’을 누르세요. 처음 한 번만 나옵니다."),
    ("h", "■ 메일·인터넷으로 받은 exe가 실행되지 않는 경우"),
    ("", "exe 파일을 우클릭 → 속성 → 아래쪽 ‘차단 해제’에 체크 → 확인을 누른 뒤 다시 실행하세요."),
    ("h", "■ 백신 프로그램이 exe를 차단하는 경우"),
    ("", "직접 만든 프로그램은 백신이 의심하는 경우가 있습니다. 백신의 예외(허용) 목록에 이 exe가 있는 폴더를 등록하세요."),
    ("h", "■ 폴더 위치 주의"),
    ("", "‘C:\\Program Files’ 같은 보호된 폴더에 두면 설정과 결과 파일이 저장되지 않을 수 있습니다. 바탕화면이나 문서 폴더에 두세요."),
    ("h", "■ ‘Pillow가 없다’는 경고가 뜨는 경우"),
    ("", "사진 미리보기에 쓰는 구성요소가 exe에 포함되지 않은 것입니다. 사용하는 컴퓨터에서 설치하는 것이 아니라,\n"
         "exe를 만드는 컴퓨터에서 build_exe.py를 다시 실행해 새 exe로 교체해야 합니다."),
    ("h", "■ 엑셀 관련 오류가 나는 경우"),
    ("", "· 작업할 엑셀 파일은 엑셀에서 닫은 뒤 실행하세요. 열려 있는 파일은 오류가 날 수 있습니다.\n"
         "· 작업 도중 엑셀 창이 남아 있으면 작업 관리자에서 EXCEL.EXE를 종료한 뒤 다시 시도하세요.\n"
         "· [엑셀 연결 테스트]가 실패하면 엑셀을 직접 한 번 실행해 라이선스 인증이나 업데이트 창을 닫아 주세요."),
    ("h", "■ exe를 만드는 방법 (관리자용)"),
    ("", "build_exe.py와 프로그램 .py 파일을 같은 폴더에 두고, 그 폴더에서 명령 프롬프트를 열어 python build_exe.py 를 실행합니다.\n"
         "끝나면 ‘배포’ 폴더가 생기고, 그 안에 exe와 InspectionSheets 폴더가 함께 들어 있습니다. 이 폴더를 통째로 복사하면 됩니다."),
]


# ─────────────────────── 실시간 카메라 촬영 대화상자 ───────────────────────
class CameraCaptureDialog(tk.Toplevel):
    """실시간 웹캠/외장 카메라 사진 촬영 및 지정 폴더 저장 대화상자"""

    def __init__(self, master, folder: str, on_photo_saved=None, suggested_name="", auto_assign_var=None, get_next_name=None):
        super().__init__(master)
        self.title("📷 사진 촬영 및 저장")
        self.geometry("700x640")
        self.minsize(580, 540)
        self.configure(bg=C["bg"])
        self.folder = folder
        self.on_photo_saved = on_photo_saved
        self.suggested_name = suggested_name
        self.auto_assign_var = auto_assign_var
        self.get_next_name = get_next_name
        self.cap = None
        self.running = False
        self.cur_frame = None
        self.photo_image = None
        self.cam_index = 0

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<space>", lambda e: self._capture())
        self.bind("<Return>", lambda e: self._capture())

        self.focus_set()
        self.after(100, lambda: self._start_camera(0))

    def _build_ui(self):
        top_bar = tk.Frame(self, bg=C["card"], padx=14, pady=10)
        top_bar.pack(fill=tk.X)
        lbl(top_bar, "📷 실시간 카메라 촬영", 11, bold=True).pack(side=tk.LEFT)

        cam_ctl = tk.Frame(top_bar, bg=C["card"])
        cam_ctl.pack(side=tk.RIGHT)
        lbl(cam_ctl, "카메라 번호:", 9, color=C["mut"]).pack(side=tk.LEFT, padx=(0, 4))
        self.cam_spin = ttk.Spinbox(cam_ctl, from_=0, to=4, width=3, command=self._on_cam_change)
        self.cam_spin.set(0)
        self.cam_spin.pack(side=tk.LEFT, padx=(0, 6))
        RoundButton(cam_ctl, "전환", self._on_cam_change, kind="ghost", height=26).pack(side=tk.LEFT, padx=(0, 8))
        RoundButton(cam_ctl, "Windows 카메라 앱 열기", self._open_win_camera, kind="ghost", height=26).pack(side=tk.LEFT)

        video_wrap = tk.Frame(self, bg="#000000", highlightthickness=1, highlightbackground=C["line"])
        video_wrap.pack(fill=tk.BOTH, expand=True, padx=14, pady=(10, 6))

        self.video_lbl = tk.Label(video_wrap, text="카메라에 연결하는 중...\n(잠시만 기다려주세요)",
                                  bg="#000000", fg="#CCCCCC", font=(FONT, 11))
        self.video_lbl.pack(fill=tk.BOTH, expand=True)

        self.status_lbl = lbl(self, "스페이스바(Space) 또는 [촬영 및 저장] 버튼을 누르면 사진이 저장됩니다.", 9, color=C["mut"], anchor="w")
        self.status_lbl.pack(fill=tk.X, padx=16, pady=(0, 6))

        ctrl = tk.Frame(self, bg=C["card"], padx=16, pady=12, highlightthickness=1, highlightbackground=C["line"])
        ctrl.pack(fill=tk.X, padx=14, pady=(0, 14))

        r1 = tk.Frame(ctrl, bg=C["card"])
        r1.pack(fill=tk.X, pady=(0, 8))
        lbl(r1, "저장 파일명:", 9, bold=True).pack(side=tk.LEFT, padx=(0, 6))
        self.name_var = tk.StringVar(value=self.suggested_name or self._make_default_name())
        self.name_entry = tk.Entry(r1, textvariable=self.name_var, font=(FONT, 10), width=30)
        self.name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        RoundButton(r1, "현재 시간으로 생성", lambda: self.name_var.set(self._make_default_name()), kind="ghost", height=28).pack(side=tk.LEFT)

        r2 = tk.Frame(ctrl, bg=C["card"])
        r2.pack(fill=tk.X)
        if self.auto_assign_var is not None:
            tk.Checkbutton(r2, text="촬영 즉시 선택된 칸에 자동 연결 (다음 빈 칸으로 이동)",
                           variable=self.auto_assign_var, bg=C["card"], activebackground=C["card"],
                           font=(FONT, 9)).pack(side=tk.LEFT)

        btn_box = tk.Frame(r2, bg=C["card"])
        btn_box.pack(side=tk.RIGHT)
        RoundButton(btn_box, "📸 사진 촬영 및 저장 (Space)", self._capture, kind="primary", height=36).pack(side=tk.RIGHT, padx=(6, 0))
        RoundButton(btn_box, "닫기", self._on_close, kind="ghost", height=36).pack(side=tk.RIGHT)

    def _make_default_name(self) -> str:
        return datetime.datetime.now().strftime("IMG_%Y%m%d_%H%M%S.jpg")

    def _on_cam_change(self):
        try:
            idx = int(self.cam_spin.get())
        except Exception:
            idx = 0
        self._start_camera(idx)

    def _start_camera(self, idx: int):
        self._stop_camera()
        self.cam_index = idx
        try:
            import cv2
            cap = cv2.VideoCapture(idx, cv2.CAP_DSHOW)
            if not cap.isOpened():
                cap = cv2.VideoCapture(idx)
            if not cap.isOpened():
                self.status_lbl.config(text=f"❌ {idx}번 카메라를 열 수 없습니다. 다른 카메라 번호를 선택해보세요.", fg=C["bad"])
                self.video_lbl.config(image="", text=f"카메라 {idx}번을 열 수 없습니다.\nUSB 연결 확인 또는 상단의 [Windows 카메라 앱 열기]를 이용하세요.")
                return
            cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1920)
            cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)
            self.cap = cap
            self.running = True
            self.status_lbl.config(text=f"✓ 카메라 {idx}번 연결 완료 (스페이스바 또는 [촬영] 클릭)", fg=C["ok"])
            self._update_stream()
        except Exception as e:
            self.status_lbl.config(text=f"❌ 카메라 시작 실패: {e}", fg=C["bad"])
            self.video_lbl.config(image="", text=f"카메라 오류: {e}")

    def _update_stream(self):
        if not self.running or self.cap is None:
            return
        ret, frame = self.cap.read()
        if ret and frame is not None:
            self.cur_frame = frame
            try:
                import cv2
                rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                h, w = rgb.shape[:2]
                vw = max(100, self.video_lbl.winfo_width())
                vh = max(100, self.video_lbl.winfo_height())
                scale = min(vw / w, vh / h)
                nw, nh = max(1, int(w * scale)), max(1, int(h * scale))
                resized = cv2.resize(rgb, (nw, nh), interpolation=cv2.INTER_LINEAR)
                img = Image.fromarray(resized)
                self.photo_image = ImageTk.PhotoImage(img)
                self.video_lbl.config(image=self.photo_image, text="")
            except Exception:
                pass
        self.after(30, self._update_stream)

    def _stop_camera(self):
        self.running = False
        if self.cap is not None:
            try:
                self.cap.release()
            except Exception:
                pass
            self.cap = None

    def _capture(self):
        if self.cur_frame is None:
            messagebox.showwarning("촬영 안내", "카메라 화면이 아직 준비되지 않았거나 연결되지 않았습니다.", parent=self)
            return

        name = self.name_var.get().strip()
        # 윈도우 파일명 금지 문자 (\ / : * ? " < > |) 치환
        name = re.sub(r'[\\/*?:"<>|]', '_', name).strip(' .')
        if not name:
            name = self._make_default_name()
        elif not name.lower().endswith((".jpg", ".jpeg", ".png")):
            name += ".jpg"

        out_path = os.path.join(self.folder, name)
        p = Path(out_path)
        base = p.stem
        # 중복 파일이 존재하는 경우 끝에 _1, _2... 붙임
        if p.exists():
            n = 1
            while p.exists():
                p = p.with_name(f"{base}_{n}{p.suffix}")
                n += 1
        out_path = str(p.resolve())

        try:
            import cv2
            is_success, buffer = cv2.imencode(".jpg", self.cur_frame, [int(cv2.IMWRITE_JPEG_QUALITY), 95])
            if is_success:
                with open(out_path, "wb") as f:
                    f.write(buffer)
            else:
                cv2.imwrite(out_path, self.cur_frame)

            self.status_lbl.config(text=f"✓ 사진 촬영 및 저장 완료: {p.name}", fg=C["ok"])

            if self.on_photo_saved:
                self.on_photo_saved(out_path)

            if self.get_next_name:
                next_n = self.get_next_name()
                self.name_var.set(next_n or self._make_default_name())
            else:
                self.name_var.set(self._make_default_name())

        except Exception as e:
            messagebox.showerror("저장 오류", f"사진을 저장하지 못했습니다: {e}", parent=self)

    def _open_win_camera(self):
        try:
            subprocess.Popen("start microsoft.windows.camera:", shell=True)
            self.status_lbl.config(
                text="✓ Windows 카메라 앱을 실행했습니다. 촬영 후 '새로 고침'을 누르면 목록에 추가됩니다.",
                fg=C["ok"]
            )
        except Exception as e:
            messagebox.showerror("실행 실패", f"카메라 앱을 열 수 없습니다: {e}", parent=self)

    def _on_close(self):
        self._stop_camera()
        if hasattr(self.master, "_camera_dialog"):
            self.master._camera_dialog = None
        self.destroy()


# ───────────────────────────── 메인 앱 ─────────────────────────────
class ExcelPhotoOneStopApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"엑셀 사진 원스톱 v{VERSION}")
        self.geometry("1300x860")
        self.minsize(1050, 720)
        self.configure(bg=C["bg"])

        self._cam_auto_assign = tk.BooleanVar(value=True)
        self._ui_queue: queue.Queue = queue.Queue()
        self._busy = False
        self._last_folder = ""
        self.loaded_quantities: dict[str, int] = {}
        self.pages: dict[str, tk.Frame] = {}
        self.nav: dict[str, tuple] = {}
        self.current_tab = ""
        self.target_excel = tk.StringVar()
        self.sheet_folder_vars: dict[str, tk.StringVar] = {}
        self.quantity_rows: list[dict] = []
        # 사진 연결(v14) 상태
        self.link_excel, self.link_folder = tk.StringVar(), tk.StringVar()
        self.slots: list[dict] = []
        self.slot_widgets: list[dict] = []
        self.slot_assign: dict[int, str] = {}
        self.slot_delete: set[int] = set()  # 기존 사진 삭제 예약 슬롯 목록
        self.selected_slot = None
        self.photos: list[dict] = []
        self.photo_index: dict[str, int] = {}
        self.link_sort = "name"
        self._gen = 0
        self._drag = None
        self._slot_of_widget: dict[str, int] = {}
        self.thumb_size = SIZES["보통"]
        self._env_items: list[dict] = []
        self._thumb_cache: dict[str, tuple] = {}  # 새로고침 때 이미 만든 미리보기 재사용
        self._slot_thumb_cache: dict[str, object] = {}  # 슬롯 카드용 썸네일 캐시
        self._new_paths: set[str] = set()  # 새로고침으로 추가된 사진 (NEW 표시)
        self.link_fit = tk.BooleanVar(value=True)
        self._blank = tk.PhotoImage(width=1, height=1)
        self._big_cache: dict[str, object] = {}
        self._preview = None
        self._hover_job = self._hide_job = self._relayout_job = None
        self._hover_key = None
        # 3. 미리보기(시트 뷰어) 및 임시 삽입(작업본) 관련
        self.preview_images_raw: list = []
        self.preview_tk_images: list = []
        self._preview_relayout_job = None
        self._staged_excel_path: str = ""
        self._staged_items_hash = None

        self._setup_style()
        self.current_tab = "link"
        self._build_layout()
        self.bind_all("<F5>", lambda e: self._refresh_all())
        self.bind_all("<MouseWheel>", self._on_wheel)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.after(100, self._drain_ui_queue)
        self._load_initial_config()

    # ── 스타일 / 레이아웃 ──
    def _setup_style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("Vertical.TScrollbar", troughcolor=C["card"], background="#CBD2E1", bordercolor=C["card"],
                    lightcolor="#CBD2E1", darkcolor="#CBD2E1", arrowcolor=C["card"], relief="flat")
        s.configure("Accent.Horizontal.TProgressbar", troughcolor=C["bg"], background=C["primary"],
                    bordercolor=C["bg"], lightcolor=C["primary"], darkcolor=C["primary"], thickness=4)

    def _build_layout(self):
        # 별도 메뉴나 다른 도구 없이 사진 연결·삽입을 한 화면에서 처리
        content = tk.Frame(self, bg=C["bg"])
        content.pack(fill=tk.BOTH, expand=True)
        head = tk.Frame(content, bg=C["bg"])
        head.pack(fill=tk.X, padx=28, pady=(20, 4))
        titles = tk.Frame(head, bg=C["bg"])
        titles.pack(side=tk.LEFT)
        self.page_title = lbl(titles, "엑셀 사진 원스톱", 19, True, bg=C["bg"])
        self.page_title.pack(anchor="w")
        self.page_sub = lbl(titles, "엑셀과 사진 폴더를 선택하면 위치를 찾고 사진을 연결해 한 번에 저장합니다.", 10, color=C["mut"], bg=C["bg"])
        self.page_sub.pack(anchor="w", pady=(2, 0))
        RoundButton(head, "?  사용방법", self._show_help, kind="ghost").pack(side=tk.RIGHT)
        self.action_slot = tk.Frame(head, bg=C["bg"])
        self.progress_holder = tk.Frame(content, bg=C["bg"], height=4)
        self.progress_holder.pack(fill=tk.X, padx=28, pady=(6, 8))
        self.progress_holder.pack_propagate(False)
        self.progress = ttk.Progressbar(self.progress_holder, mode="indeterminate", style="Accent.Horizontal.TProgressbar")
        logbox = tk.Frame(head, bg="#111827", highlightthickness=1, highlightbackground="#273244")
        logbox.pack(side=tk.RIGHT, padx=(14, 0), pady=(0, 2))
        self.status_text = tk.Text(logbox, height=3, width=48, bg="#111827", fg="#E5E7EB", font=("Consolas", 8), relief="flat", wrap="word", padx=7, pady=5, highlightthickness=0)
        self.status_text.pack(fill=tk.BOTH, expand=True)
        self.status_text.insert("1.0", "진행 상황이 여기에 표시됩니다.")
        self.status_text.configure(state=tk.DISABLED)
        for tag, color in [("ok", "#6EE7A8"), ("bad", "#FCA5A5"), ("info", "#93C5FD"), ("sep", "#4B5563")]:
            self.status_text.tag_config(tag, foreground=color)
        self.body = tk.Frame(content, bg=C["bg"])
        self.body.pack(fill=tk.BOTH, expand=True, padx=28)
        self.pages["link"] = self._page_link()
        self.pages["link"].pack(fill=tk.BOTH, expand=True)

    def _nav_item(self, parent, key, icon, text, color):
        row = tk.Frame(parent, bg=C["side"], cursor="hand2")
        row.pack(fill=tk.X)
        bar = tk.Frame(row, width=4, bg=C["side"])
        bar.pack(side=tk.LEFT, fill=tk.Y)
        lab = tk.Label(row, text=f"{icon}   {text}", font=(FONT, 11), bg=C["side"], fg="#C7D0EA", anchor="w", padx=18, pady=12)
        lab.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.nav[key] = (row, bar, lab, color)
        for w in (row, bar, lab):
            w.bind("<Button-1>", lambda e, k=key: self._switch_tab(k))
            w.bind("<Enter>", lambda e, k=key: self._paint_nav(k, hover=True))
            w.bind("<Leave>", lambda e, k=key: self._paint_nav(k))

    def _paint_nav(self, key, hover=False):
        row, bar, lab, color = self.nav[key]
        active = key == self.current_tab
        bg = C["side_hi"] if (active or hover) else C["side"]
        row.config(bg=bg)
        bar.config(bg=color if active else bg)
        lab.config(bg=bg, fg="white" if active else "#C7D0EA", font=(FONT, 11, "bold" if active else "normal"))

    def _switch_tab(self, key: str):
        self.current_tab = key
        for k in self.nav:
            self._paint_nav(k)
        for page in self.pages.values():
            page.pack_forget()
        self.pages[key].pack(fill=tk.BOTH, expand=True)
        title, sub = PAGE_INFO[key]
        self.page_title.config(text=title)
        self.page_sub.config(text=sub)
        for w in self.action_slot.winfo_children():
            w.destroy()
        actions = {
            "quantity": ("설비별 시트 복사 실행  ▶", self._start_copy_from_quantity),
            "copy": ("설비별 시트 복사 시작  ▶", self._start_copy_from_file),
            "image": ("🚀  이미지 삽입 실행", self._start_image_insert),
            "merge": ("엑셀 통합 시작  ▶", self._start_merge),
            "pdf": ("PDF 변환 시작  ▶", self._start_pdf),
            "link": ("💾  저장", self._save_inserted_file),
            "env": ("다시 점검  ▶", self._run_env_check),
        }
        text, cmd = actions[key]
        RoundButton(self.action_slot, text, cmd, kind="success", height=36).pack()

    def _bind_shortcuts(self):
        keys = {"1": "image", "2": "link", "i": "image", "l": "link"}
        for key, tab in keys.items():
            for k in {key, key.upper()}:
                self.bind_all(f"<Control-{k}>", lambda e, t=tab: self._switch_tab(t))

    def _on_wheel(self, event):
        try:
            w = self.winfo_containing(event.x_root, event.y_root)
        except Exception:
            return
        while w is not None:
            if w in SCROLL_CANVASES:
                w.yview_scroll(int(-event.delta / 120), "units")
                return
            w = w.master

    def _show_help(self):
        messagebox.showinfo(
            "사용방법",
            "1. 대상 엑셀 파일을 선택합니다.\n"
            "2. 사진이 들어 있는 폴더를 선택합니다.\n"
            "3. 엑셀의 {{이미지}} 표식과 기존 사진 위치를 자동으로 찾습니다.\n"
            "4. {{이미지:파일명.jpg}}처럼 이름을 적어 두면 같은 이름의 사진을 우선 자동 연결합니다.\n"
            "5. 연결이 안 된 사진은 왼쪽 위치를 선택한 뒤 오른쪽 사진을 클릭하거나 드래그해 연결합니다.\n"
            "6. [사진 삽입 및 저장]을 누르면 새 엑셀 파일이 만들어집니다.\n\n"
            "원본 파일은 기본적으로 보존됩니다."
        )

    def _show_about(self):
        messagebox.showinfo("프로그램 정보", f"엑셀 이미지 자동화 전문가 v{VERSION}")

    def _open_any_file(self):
        path = filedialog.askopenfilename(title="파일 열기", filetypes=[
            ("모든 파일", "*.*"), ("Excel 파일", "*.xlsx *.xlsm *.xls"), ("PDF 파일", "*.pdf"),
            ("이미지 파일", "*.jpg *.jpeg *.png *.bmp *.gif")])
        if path:
            try:
                open_with_default_app(path)
            except Exception as e:
                messagebox.showerror("오류", f"파일을 열 수 없습니다: {e}")

    # ── 스레드 안전 UI 호출 / 로그 / 작업 실행 ──
    def _drain_ui_queue(self):
        try:
            while True:
                fn = self._ui_queue.get_nowait()
                try:
                    fn()
                except Exception:
                    import traceback
                    self._append_log_now("❌ 화면 처리 중 오류:\n" + traceback.format_exc())
        except queue.Empty:
            pass
        self.after(100, self._drain_ui_queue)

    def call_ui(self, fn):
        self._ui_queue.put(fn)

    def log(self, message: str):
        """어느 스레드에서 호출해도 안전합니다. ✓/❌ 줄은 색으로 구분됩니다."""
        self.call_ui(lambda: self._append_log_now(message))

    def _append_log_now(self, message: str):
        self.status_text.configure(state=tk.NORMAL)
        for line in message.split("\n"):
            s = line.strip()
            tag = ("ok" if s.startswith("✓") else "bad" if s.startswith("❌") else
                   "info" if s.startswith(("🔄", "[")) else "sep" if s.startswith("=") else "")
            self.status_text.insert(tk.END, line + "\n", tag)
        # 상단 로그는 화면을 차지하지 않도록 최근 3줄만 표시
        lines = self.status_text.get("1.0", "end-1c").splitlines()
        if len(lines) > 3:
            self.status_text.delete("1.0", f"{len(lines) - 2}.0")
        self.status_text.see(tk.END)
        self.status_text.configure(state=tk.DISABLED)

    def _set_busy(self, busy: bool):
        self._busy = busy
        if busy:
            self.progress.pack(fill=tk.X)
            self.progress.start(12)
        else:
            self.progress.stop()
            self.progress.pack_forget()

    def run_task(self, title: str, fn, *args, on_done=None):
        """엑셀(COM) 작업을 백그라운드에서 실행. 동시에 하나만 실행합니다."""
        if self._busy:
            messagebox.showinfo("작업 중", "이미 다른 작업이 실행 중입니다. 끝난 뒤 다시 시도하세요.")
            return
        self._set_busy(True)
        self.log("=" * 40 + f"\n[{title}] 시작...")

        def worker():
            pythoncom = None
            try:
                import pythoncom  # type: ignore
                pythoncom.CoInitialize()
            except ImportError:
                pass
            try:
                result = fn(*args)
                if on_done:
                    self.call_ui(lambda: on_done(result))
            except Exception as e:
                self.log(f"❌ {title} 중 오류: {e}")
            finally:
                if pythoncom:
                    pythoncom.CoUninitialize()
                self.call_ui(lambda: self._set_busy(False))

        threading.Thread(target=worker, daemon=True).start()

    # ── 공용 폼 조각 ──
    def _new_page(self):
        page = tk.Frame(self.body, bg=C["bg"])
        box = card(page)
        box.pack(fill=tk.BOTH, expand=True, pady=(0, 4))
        return page, box

    def _form_row(self, parent, label, var, browse=None):
        row = tk.Frame(parent, bg=C["card"])
        row.pack(fill=tk.X, padx=28, pady=7)
        lbl(row, label, 10, True).pack(anchor="w")
        line = tk.Frame(row, bg=C["card"])
        line.pack(fill=tk.X, pady=(5, 0))
        field_entry(line, var).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6)
        if browse:
            RoundButton(line, "찾아보기", browse, kind="ghost").pack(side=tk.LEFT, padx=(8, 0))

    # ── 1. 수량 입력 ──
    def _page_quantity(self):
        page, box = self._new_page()
        bar = tk.Frame(box, bg=C["card"])
        bar.pack(fill=tk.X, padx=22, pady=(16, 6))
        lbl(bar, "설비별 수량", 12, True).pack(side=tk.LEFT)
        self.qty_seoul_var = tk.BooleanVar(value=False)
        check(bar, "서울형 양식", self.qty_seoul_var).pack(side=tk.RIGHT)

        foot = tk.Frame(box, bg=C["card"])
        foot.pack(side=tk.BOTTOM, fill=tk.X, padx=22, pady=14)
        for text, cmd, kind in [("엑셀 불러오기", self._load_quantity_excel, "ghost"),
                                ("엑셀 저장", self._save_quantity_excel, "ghost"),
                                ("수량 비우기", self._clear_quantities, "danger"),
                                ("기본 목록으로", self._reset_quantity_list, "ghost")]:
            RoundButton(foot, text, cmd, kind=kind).pack(side=tk.LEFT, padx=(0, 8))
        RoundButton(foot, "설비별 시트 복사 실행  ▶", self._start_copy_from_quantity, height=38).pack(side=tk.RIGHT)
        area = tk.Frame(box, bg=C["card"])
        area.pack(fill=tk.BOTH, expand=True, padx=22)
        self.qty_canvas, inner = make_scroll_area(area, C["card"])
        self.qty_inner = inner
        inner.grid_columnconfigure((0, 1), weight=1, uniform="col")
        for col in (0, 1):
            head = tk.Frame(inner, bg=C["card"])
            head.grid(row=0, column=col, sticky="w", padx=(0, 18), pady=(0, 4))
            lbl(head, "설비명", 9, True, C["mut"], width=18, anchor="w").pack(side=tk.LEFT)
            lbl(head, "수량", 9, True, C["mut"]).pack(side=tk.LEFT)
        self.qty_columns = []
        for col in (0, 1):
            frame = tk.Frame(inner, bg=C["card"])
            frame.grid(row=1, column=col, sticky="nw", padx=(0, 18))
            self.qty_columns.append(frame)
        self._fill_quantity_rows([(name, "1") for name in DEFAULT_EQUIPMENT])

        return page

    def _fill_quantity_rows(self, items: list[tuple[str, str]]):
        for row in self.quantity_rows:
            row["frame"].destroy()
        self.quantity_rows.clear()
        half = (len(items) + 1) // 2
        for idx, (name, qty) in enumerate(items):
            self._create_quantity_row(self.qty_columns[0 if idx < half else 1], name, qty)

    def _create_quantity_row(self, column: tk.Frame, name: str, qty: str):
        frame = tk.Frame(column, bg=C["card"])
        frame.pack(fill=tk.X, pady=2)
        name_var, qty_var = tk.StringVar(value=name), tk.StringVar(value=qty)
        name_entry = field_entry(frame, name_var, width=18)
        name_entry.pack(side=tk.LEFT, padx=(0, 6), ipady=3)
        qty_entry = field_entry(frame, qty_var, width=5, justify="center")
        qty_entry.pack(side=tk.LEFT, padx=(0, 6), ipady=3)
        clear = tk.Label(frame, text="✕", font=(FONT, 10), fg="#B0B7C6", bg=C["card"], cursor="hand2", padx=4)
        clear.pack(side=tk.LEFT)
        clear.bind("<Button-1>", lambda e: qty_var.set(""))
        clear.bind("<Enter>", lambda e: clear.config(fg=C["bad"]))
        clear.bind("<Leave>", lambda e: clear.config(fg="#B0B7C6"))

        row = {"frame": frame, "name": name_var, "qty": qty_var, "qty_entry": qty_entry}
        self.quantity_rows.append(row)
        name_entry.bind("<Tab>", lambda e: (qty_entry.focus_set(), "break")[1])
        qty_entry.bind("<Tab>", lambda e: self._focus_quantity(row, +1))
        qty_entry.bind("<Shift-Tab>", lambda e: self._focus_quantity(row, -1))

    def _focus_quantity(self, row: dict, step: int):
        idx = (self.quantity_rows.index(row) + step) % len(self.quantity_rows)
        target = self.quantity_rows[idx]
        target["qty_entry"].focus_set()
        target["qty_entry"].select_range(0, tk.END)
        self.update_idletasks()
        total = max(1, self.qty_inner.winfo_height())
        self.qty_canvas.yview_moveto(max(0.0, min(1.0, target["frame"].winfo_y() / total)))
        return "break"

    def _collect_quantities(self) -> dict[str, int]:
        result: dict[str, int] = {}
        for row in self.quantity_rows:
            name, qty = row["name"].get().strip(), row["qty"].get().strip()
            if name and qty:
                try:
                    result[name] = int(qty)
                except ValueError:
                    self.log(f"❌ 수량이 숫자가 아닙니다: {name} - {qty}")
        return result

    def _clear_quantities(self):
        for row in self.quantity_rows:
            row["qty"].set("")
        self.log("✓ 모든 수량을 빈 칸으로 초기화했습니다.")

    def _reset_quantity_list(self):
        self._fill_quantity_rows([(name, "1") for name in DEFAULT_EQUIPMENT])
        self.log("✓ 수량 입력 목록을 기본값으로 되돌렸습니다.")

    def _load_quantity_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")], title="수량 데이터 불러오기")
        if not path:
            return

        def done(data: dict[str, int]):
            if not data:
                messagebox.showwarning("경고", "불러올 수량 데이터가 없습니다.")
                return
            self.loaded_quantities = dict(data)
            self._fill_quantity_rows([(n, str(q)) for n, q in data.items()])
            self.log(f"✓ 수량 데이터 불러오기 완료: {path}")
            messagebox.showinfo("완료", f"수량 데이터를 불러왔습니다.\n총 {len(data)}개 설비")

        self.run_task("수량 불러오기", read_quantity_file, path, on_done=done)

    def _save_quantity_excel(self):
        data = list(self._collect_quantities().items())
        if not data:
            messagebox.showwarning("경고", "저장할 데이터가 없습니다.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
                                            title="수량 데이터 저장")
        if path:
            self.run_task("수량 저장", save_quantity_file, path, data,
                          on_done=lambda _: self.log(f"✓ 수량 데이터 저장 완료: {path}"))

    def _start_copy_from_quantity(self):
        qty = self._collect_quantities()
        if not qty:
            messagebox.showerror("오류", "수량 데이터가 없습니다.\n\n수량을 입력한 후 다시 시도하세요.")
            return
        self.loaded_quantities = dict(qty)
        self.log(f"✓ 수량 입력에서 {len(qty)}개 설비 데이터 수집 완료")
        self._run_copy(qty, self.qty_seoul_var.get())

    # ── 2. 시트 복사 ──
    def _page_copy(self):
        page, box = self._new_page()
        inner = tk.Frame(box, bg=C["card"])
        inner.pack(expand=True)
        lbl(inner, "수량 입력 탭에서 이미 실행했다면 이 단계는 필요 없습니다.", 9, color=C["mut"]).pack(pady=(0, 14))
        self.copy_file_card = DropCard(inner, "대상점검및수량.xlsx")
        self.copy_file_card.pack()
        self.copy_seoul_var = tk.BooleanVar(value=False)
        check(inner, "서울형 양식 사용", self.copy_seoul_var).pack(pady=(14, 10))
        RoundButton(inner, "설비별 시트 복사 시작  ▶", self._start_copy_from_file, height=38).pack()
        return page

    def _start_copy_from_file(self):
        qty_file = self.copy_file_card.path.get().strip()
        if qty_file and os.path.isfile(qty_file):
            self.log(f"수량 파일: {qty_file}")
            self.run_task("수량 파일 읽기", read_quantity_file, qty_file,
                          on_done=lambda data: self._run_copy(data, self.copy_seoul_var.get()) if data
                          else self.log("❌ 설비 수량 정보를 찾을 수 없습니다."))
        elif self.loaded_quantities:
            self.log("✓ 수량 입력/불러오기 데이터를 사용합니다.")
            self._run_copy(self.loaded_quantities, self.copy_seoul_var.get())
        else:
            messagebox.showerror("오류", "수량 데이터가 없습니다.\n\n수량 입력 탭에서 데이터를 불러오거나,\n"
                                         "대상점검및수량.xlsx 파일을 선택하세요.")

    def _run_copy(self, equip_qty: dict[str, int], seoul: bool):
        folder = SEOUL_TEMPLATE_DIR if seoul else TEMPLATE_DIR
        if not folder.is_dir():
            hint = "폴더가 존재하거나 체크박스를 해제하세요." if seoul else "폴더가 존재하는지 확인하세요."
            messagebox.showerror("오류", f"점검 양식 폴더를 찾을 수 없습니다:\n{folder}\n\n{hint}")
            return
        self.log(f"[설비별 시트 복사] 폴더: {folder}\n✓ 설비 수량 정보: {equip_qty}")
        save_name = "설비별_점검_통합본_서울형.xlsx" if seoul else "설비별_점검_통합본.xlsx"

        def done(path):
            if path:
                self.log(f"\n✓ 모든 시트 복사 및 저장 완료: {path}")

        self.run_task("설비별 시트 복사", lambda: copy_sheets_by_quantity(equip_qty, folder, save_name, self.log), on_done=done)

    # ── 3. 이미지 삽입 ──
    def _page_image(self):
        page, box = self._new_page()
        top = tk.Frame(box, bg=C["card"])
        top.pack(fill=tk.X, padx=22, pady=(18, 10))
        self.target_card = DropCard(top, "대상 엑셀", self._on_target_selected, width=400, height=140)
        self.target_card.pack(side=tk.LEFT)
        tip = tk.Frame(top, bg=C["card"])
        tip.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(22, 0))
        lbl(tip, "이렇게 쓰세요", 10, True).pack(anchor="w")
        lbl(tip, "사진이 들어갈 엑셀 칸에 아래처럼 적어 두세요.\n병합된 칸도 괜찮고, 사진은 칸 크기에 맞춰 들어갑니다.",
            9, color=C["mut"], justify="left").pack(anchor="w", pady=(4, 8))
        tk.Label(tip, text="{{이미지:101.jpg}}", font=("Consolas", 11, "bold"), fg=C["primary"], bg="#EEF1FB",
                 padx=12, pady=6).pack(anchor="w")
        lbl(tip, "※ 원본은 보존되며, 결과는 파일명_사진삽입.xlsx로 새로 저장됩니다.", 9, color=C["bad"], justify="left").pack(anchor="w", pady=(10, 0))

        foot = tk.Frame(box, bg=C["card"])
        foot.pack(side=tk.BOTTOM, fill=tk.X, padx=22, pady=14)
        RoundButton(foot, "🚀  이미지 삽입 실행", self._start_image_insert, kind="success", height=38).pack(side=tk.RIGHT)
        lbl(box, "시트별 이미지 폴더", 10, True).pack(anchor="w", padx=22, pady=(6, 4))
        area = tk.Frame(box, bg=C["card"], highlightthickness=1, highlightbackground=C["line"])
        area.pack(fill=tk.BOTH, expand=True, padx=22)
        _, self.sheet_frame = make_scroll_area(area, C["card"])
        self.sheet_msg = lbl(self.sheet_frame, "대상 엑셀을 올리면 시트 목록이 여기에 나타납니다.", 9, color=C["mut"], anchor="w")
        self.sheet_msg.pack(anchor="w", padx=10, pady=10)

        return page

    def _on_target_selected(self, path: str):
        self.target_excel.set(path)
        self._refresh_sheet_rows(path)

    def _refresh_sheet_rows(self, path: str):
        self.sheet_msg.config(text="시트 목록을 읽는 중...")
        self.run_task("시트 목록 읽기", get_image_sheet_names, path, on_done=self._populate_sheet_rows)

    def _populate_sheet_rows(self, names: list[str]):
        for widget in self.sheet_frame.winfo_children():
            if widget is not self.sheet_msg:
                widget.destroy()
        self.sheet_folder_vars.clear()
        self.sheet_msg.config(text="" if names else "(대상 엑셀에서 이름에 ‘-’가 들어간 시트가 없습니다)")
        for name in names:
            var = tk.StringVar()
            self.sheet_folder_vars[name] = var
            row = tk.Frame(self.sheet_frame, bg=C["card"])
            row.pack(fill=tk.X, padx=10, pady=3)
            lbl(row, name, 10, True, width=18, anchor="w").pack(side=tk.LEFT)
            entry = field_entry(row, var)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4, padx=(0, 8))
            var.trace_add("write", lambda *a, e=entry: e.xview_moveto(1))
            RoundButton(row, "폴더 선택", lambda v=var: self._select_folder(v), kind="ghost", height=30).pack(side=tk.LEFT, padx=(0, 6))
            RoundButton(row, "초기화", lambda v=var: v.set(""), kind="danger", height=30).pack(side=tk.LEFT)

    def _select_folder(self, var: tk.StringVar):
        if self._last_folder:
            initial = os.path.dirname(self._last_folder)
        else:
            target = self.target_excel.get()
            initial = os.path.dirname(target) if target and os.path.exists(target) else ""
        folder = filedialog.askdirectory(initialdir=initial)
        if folder:
            var.set(folder)
            self._last_folder = folder

    def _start_image_insert(self):
        target = self.target_excel.get().strip()
        if not target or not os.path.isfile(target):
            messagebox.showerror("입력 오류", "대상 엑셀 파일을 먼저 선택해주세요.")
            return
        folder_map = {name: var.get().strip() for name, var in self.sheet_folder_vars.items()}
        target = os.path.abspath(os.path.normpath(target))
        p = Path(target)
        orig_dir = p.parent
        base = p.stem + "_사진삽입"
        out_p = orig_dir / f"{base}{p.suffix}"
        n = 2
        while out_p.exists():
            out_p = orig_dir / f"{base}_{n}{p.suffix}"
            n += 1
        out = str(out_p.resolve())

        def done(result):
            ok, fail = result
            self.log(f"\n✓ 작업 완료 — 성공: {ok}개" + (f", 실패: {fail}개" if fail else ""))
            self.log(f"✓ 원본은 보존하고 새 파일로 저장했습니다: {out}")
            try:
                open_with_default_app(out)
            except Exception as e:
                self.log(f"엑셀 파일 자동 오픈 실패: {e}")

        self.run_task("이미지 삽입", lambda: insert_images(target, folder_map, self.log, out), on_done=done)

    # ── 4. 엑셀 통합 ──
    def _page_merge(self):
        page, box = self._new_page()
        self.merge_company, self.merge_folder, self.merge_output = tk.StringVar(), tk.StringVar(), tk.StringVar()
        tk.Frame(box, bg=C["card"], height=12).pack()
        self._form_row(box, "고객사명", self.merge_company)
        self._form_row(box, "통합할 엑셀 파일이 있는 폴더", self.merge_folder, self._browse_merge_folder)
        self._form_row(box, "저장 파일명 (고객사명을 쓰면 자동 완성)", self.merge_output)
        foot = tk.Frame(box, bg=C["card"])
        foot.pack(fill=tk.X, padx=28, pady=18)
        RoundButton(foot, "엑셀 통합 시작  ▶", self._start_merge, height=38).pack(side=tk.RIGHT)
        self.merge_company.trace_add("write", lambda *a: self.merge_output.set(report_filename(self.merge_company.get(), "xlsx")))
        self.merge_output.set(report_filename("", "xlsx"))
        return page

    def _browse_merge_folder(self):
        folder = filedialog.askdirectory(title="엑셀 파일 폴더 선택")
        if folder:
            self.merge_folder.set(folder)

    def _start_merge(self):
        folder, output = self.merge_folder.get().strip(), self.merge_output.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("오류", "엑셀 파일 폴더를 올바르게 선택하세요.")
            return
        if not output:
            messagebox.showerror("오류", "저장 파일명을 입력하세요.")
            return
        out_path = output if os.path.isabs(output) else os.path.join(folder, output)

        def done(success):
            if success:
                self.log(f"✓ 통합 완료: {out_path}")
                try:
                    open_with_default_app(out_path)
                except Exception as e:
                    self.log(f"(엑셀 자동 오픈 실패: {e})")

        self.run_task("엑셀 통합", lambda: merge_workbooks(folder, out_path, self.log), on_done=done)

    # ── 5. PDF 변환 ──
    def _page_pdf(self):
        page, box = self._new_page()
        self.pdf_company, self.pdf_quality = tk.StringVar(), tk.BooleanVar(value=True)
        tk.Frame(box, bg=C["card"], height=12).pack()
        self._form_row(box, "고객사명", self.pdf_company)
        holder = tk.Frame(box, bg=C["card"])
        holder.pack(pady=(6, 6))
        self.pdf_card = DropCard(holder, "엑셀")
        self.pdf_card.pack()
        check(box, "고품질(고해상도)로 PDF 저장", self.pdf_quality).pack(pady=(4, 6))
        RoundButton(box, "PDF 변환 시작  ▶", self._start_pdf, height=38).pack(pady=(6, 0))
        return page

    def _start_pdf(self):
        raw = self.pdf_card.path.get().strip()
        excel_path = raw if os.path.isfile(raw) else unquote(raw)
        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showerror("오류", "엑셀 파일을 올바르게 선택하세요.")
            return
        excel_path = os.path.normpath(excel_path)
        pdf_path = os.path.normpath(os.path.join(os.path.dirname(excel_path),
                                                 report_filename(self.pdf_company.get(), "pdf")))
        high = self.pdf_quality.get()

        def done(_):
            self.log(f"✓ PDF 저장 완료: {pdf_path}")
            if os.path.exists(pdf_path):
                try:
                    open_with_default_app(pdf_path)
                except Exception as e:
                    self.log(f"(PDF 자동 열기 실패: {e})")

        self.log(f"[PDF 변환] {excel_path}")
        self.run_task("PDF 변환", export_pdf, excel_path, pdf_path, high, on_done=done)

    # ── 6. 사진 연결 (미리보기 방식, v14 신규) ──
    def _page_link(self):
        page, box = self._new_page()

        # 상단 버튼 바: 엑셀 선택, 폴더 선택, 새로 고침
        btn_bar = tk.Frame(box, bg=C["card"])
        btn_bar.pack(fill=tk.X, padx=12, pady=(10, 8))

        RoundButton(btn_bar, "엑셀 선택", self._browse_link_excel, kind="ghost", height=30).pack(side=tk.LEFT, padx=(0, 6))
        RoundButton(btn_bar, "폴더 선택", self._browse_link_folder, kind="ghost", height=30).pack(side=tk.LEFT, padx=(0, 6))
        RoundButton(btn_bar, "📷 사진 촬영", self._open_camera_capture, kind="ghost", height=30).pack(side=tk.LEFT, padx=(0, 6))
        RoundButton(btn_bar, "새로 고침", self._refresh_all, kind="ghost", height=30).pack(side=tk.LEFT, padx=(0, 14))

        self.link_path_info = lbl(btn_bar, "엑셀 파일과 사진 폴더를 선택하거나 여기에 끌어다 놓으세요", 9, color=C["mut"], anchor="w")
        self.link_path_info.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.pil_row = tk.Frame(box, bg=C["card"])
        if getattr(sys, "frozen", False):
            lbl(self.pil_row, "⚠ 이 exe에는 사진 미리보기(Pillow)가 포함되어 있지 않습니다. Pillow 포함 버전으로 다시 빌드하세요.", 9, True, C["bad"]).pack(side=tk.LEFT)
        else:
            lbl(self.pil_row, "⚠ 사진 미리보기에는 Pillow 라이브러리가 필요합니다.", 9, True, C["bad"]).pack(side=tk.LEFT)
            RoundButton(self.pil_row, "Pillow 설치", self._install_pillow, kind="danger", height=26).pack(side=tk.LEFT, padx=8)
        if not PIL_OK:
            self.pil_row.pack(fill=tk.X, padx=12, pady=2)

        for w in (box, btn_bar, self.link_path_info):
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<Drop>>", self._on_link_drop)

        # 하단 바
        foot = tk.Frame(box, bg=C["card"])
        foot.pack(side=tk.BOTTOM, fill=tk.X, padx=12, pady=8)
        self.link_summary = lbl(foot, "칸 0개 · 연결 0개 · 비어 있는 칸 0개 · 사진 0장", 9, color=C["mut"])
        self.link_summary.pack(side=tk.LEFT)
        RoundButton(foot, "저장", self._save_inserted_file, kind="success", height=36).pack(side=tk.RIGHT)
        RoundButton(foot, "사진 삽입", self._apply_photo_insert, kind="primary", height=36).pack(side=tk.RIGHT, padx=(0, 7))
        RoundButton(foot, "연결 초기화", self._clear_all_links, kind="danger", height=36).pack(side=tk.RIGHT, padx=(0, 7))
        RoundButton(foot, "자동 연결", self._auto_match_and_fill, kind="ghost", height=36).pack(side=tk.RIGHT, padx=(0, 7))

        # 메인 바디 (3분할: 좌측 상/하 + 우측 전체)
        body = tk.Frame(box, bg=C["card"])
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 4))
        body.grid_columnconfigure(0, weight=5, uniform="main_panes")
        body.grid_columnconfigure(1, weight=6, uniform="main_panes")
        body.grid_rowconfigure(0, weight=1)

        # 좌측 패널 (① 상: 엑셀 사진 위치, ② 하: 사진 목록)
        left_pane = tk.Frame(body, bg=C["card"])
        left_pane.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        left_pane.grid_rowconfigure(0, weight=1)
        left_pane.grid_rowconfigure(1, weight=1)
        left_pane.grid_columnconfigure(0, weight=1)

        # 좌측 상단: ① 엑셀 사진 위치
        top_left = tk.Frame(left_pane, bg=C["card"])
        top_left.grid(row=0, column=0, sticky="nsew", pady=(0, 4))
        top_left.grid_rowconfigure(1, weight=1)
        top_left.grid_columnconfigure(0, weight=1)
        lbl(top_left, "① 엑셀 사진 위치", 10, True).grid(row=0, column=0, sticky="w", pady=(0, 3))
        a1 = tk.Frame(top_left, bg=C["card"], highlightthickness=1, highlightbackground=C["line"])
        a1.grid(row=1, column=0, sticky="nsew")
        self.slot_canvas, self.slot_inner = make_scroll_area(a1, C["card"])
        self.slot_inner.grid_columnconfigure((0, 1), weight=1, uniform="slot_col")

        # 좌측 하단: ② 사진 (클릭하거나 끌어다 놓기)
        bottom_left = tk.Frame(left_pane, bg=C["card"])
        bottom_left.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        bottom_left.grid_rowconfigure(1, weight=1)
        bottom_left.grid_columnconfigure(0, weight=1)
        b_hdr = tk.Frame(bottom_left, bg=C["card"])
        b_hdr.grid(row=0, column=0, sticky="ew", pady=(0, 3))
        lbl(b_hdr, "② 사진 (클릭하거나 끌어다 놓기)", 10, True).pack(side=tk.LEFT)
        RoundButton(b_hdr, "📷 사진 촬영", self._open_camera_capture, kind="primary", height=24).pack(side=tk.RIGHT)
        a2 = tk.Frame(bottom_left, bg=C["card"], highlightthickness=1, highlightbackground=C["line"])
        a2.grid(row=1, column=0, sticky="nsew")
        self.photo_canvas, self.photo_inner = make_scroll_area(a2, C["card"])
        self.photo_canvas.bind("<Configure>", lambda e: self._schedule_relayout(), add="+")

        # 우측 패널: 3. 미리보기 (엑셀 시트/페이지 뷰어)
        right_pane = tk.Frame(body, bg=C["card"])
        right_pane.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        right_pane.grid_rowconfigure(1, weight=1)
        right_pane.grid_columnconfigure(0, weight=1)

        preview_head = tk.Frame(right_pane, bg=C["card"])
        preview_head.grid(row=0, column=0, sticky="ew", pady=(0, 3))
        lbl(preview_head, "3. 미리보기", 10, True).pack(side=tk.LEFT)
        self.preview_page_info = lbl(preview_head, "", 9, color=C["mut"])
        self.preview_page_info.pack(side=tk.RIGHT)

        a3 = tk.Frame(right_pane, bg=C["card"], highlightthickness=1, highlightbackground=C["line"])
        a3.grid(row=1, column=0, sticky="nsew")
        self.preview_canvas, self.preview_inner = make_scroll_area(a3, "#F3F4F6")
        self.preview_canvas.bind("<Configure>", lambda e: self._schedule_preview_relayout(), add="+")

        self._show_slot_msg("대상 엑셀을 선택하면 사진을 넣을 위치가 여기에 나타납니다.")
        self._show_photo_msg("사진 폴더를 선택하면 미리보기가 여기에 나타납니다.")
        self._show_preview_msg("대상 엑셀을 선택하면 시트 인쇄 미리보기가 여기에 나타납니다.")
        return page

    def _show_slot_msg(self, text):
        for w in self.slot_inner.winfo_children():
            w.destroy()
        lbl(self.slot_inner, text, 9, color=C["mut"], anchor="w", justify="left", wraplength=260).pack(anchor="w", padx=10, pady=10)

    def _show_photo_msg(self, text):
        for w in self.photo_inner.winfo_children():
            w.destroy()
        lbl(self.photo_inner, text, 9, color=C["mut"], anchor="w", justify="left", wraplength=260).grid(row=0, column=0, padx=10, pady=10, sticky="w")

    def _show_preview_msg(self, text):
        for w in self.preview_inner.winfo_children():
            w.destroy()
        lbl(self.preview_inner, text, 9, color=C["mut"], bg="#F3F4F6", anchor="center", justify="center").pack(padx=20, pady=40)

    def _update_link_paths_display(self):
        e = self.link_excel.get().strip()
        f = self.link_folder.get().strip()
        parts = []
        if e:
            parts.append(f"엑셀: {os.path.basename(e)}")
        if f:
            parts.append(f"사진: {os.path.basename(f) or f}")
        if parts:
            self.link_path_info.config(text="  |  ".join(parts), fg=C["primary"])
        else:
            self.link_path_info.config(text="엑셀 파일과 사진 폴더를 선택하거나 여기에 끌어다 놓으세요", fg=C["mut"])

    def _populate_previews(self, preview_images: list):
        self.preview_images_raw = preview_images
        self.preview_tk_images = []
        for w in self.preview_inner.winfo_children():
            w.destroy()
        if not preview_images:
            self._show_preview_msg("시트 미리보기를 생성할 수 없거나 문서가 비어 있습니다.")
            self.preview_page_info.config(text="")
            return

        self.preview_page_info.config(text=f"총 {len(preview_images)}페이지")
        self._render_previews_to_canvas()

    def _schedule_preview_relayout(self):
        if self._preview_relayout_job:
            self.after_cancel(self._preview_relayout_job)
        self._preview_relayout_job = self.after(150, self._render_previews_to_canvas)

    def _render_previews_to_canvas(self):
        self._preview_relayout_job = None
        if not self.preview_images_raw:
            return
        cw = max(240, self.preview_canvas.winfo_width() - 36)
        self.preview_tk_images = []
        for w in self.preview_inner.winfo_children():
            w.destroy()

        for idx, im in enumerate(self.preview_images_raw):
            w, h = im.size
            ratio = cw / w
            th = max(50, int(h * ratio))
            resized = im.resize((cw, th), Image.Resampling.LANCZOS)
            tk_img = ImageTk.PhotoImage(resized)
            self.preview_tk_images.append(tk_img)

            card_box = tk.Frame(self.preview_inner, bg="white", highlightthickness=1, highlightbackground="#D1D5DB")
            card_box.pack(padx=14, pady=10, fill=tk.X)

            img_lbl = tk.Label(card_box, image=tk_img, bg="white")
            img_lbl.pack(fill=tk.BOTH, expand=True)

            page_tag = tk.Label(card_box, text=f"- {idx + 1} 페이지 -", font=(FONT, 8), fg="#9CA3AF", bg="white")
            page_tag.pack(pady=(2, 4))

    def _browse_link_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if path:
            self._set_link_excel(path)

    def _browse_link_folder(self):
        folder = filedialog.askdirectory(title="점검 사진 폴더 선택")
        if folder:
            self._set_link_folder(folder)

    def _on_link_drop(self, event):
        for item in self.tk.splitlist(event.data):
            if os.path.isdir(item):
                self._set_link_folder(item)
            elif item.lower().endswith((".xlsx", ".xlsm", ".xls")):
                self._set_link_excel(item)

    def _set_link_excel(self, path: str):
        path = os.path.abspath(os.path.normpath(path))
        self.link_excel.set(path)
        cfg = load_config()
        cfg["link_excel_path"] = path
        save_config(cfg)
        self._update_link_paths_display()
        self.run_task("점검 칸 및 미리보기 찾기", scan_image_slots, path, self.log, on_done=self._populate_slots)

    def _set_link_folder(self, folder: str):
        self.link_folder.set(folder)
        cfg = load_config()
        cfg["link_folder_path"] = folder
        save_config(cfg)
        self._update_link_paths_display()
        self._thumb_cache.clear()
        self._new_paths.clear()
        self._load_photos()

    def _refresh_photos(self):
        """사진 폴더 목록만 다시 읽는다."""
        folder = self.link_folder.get()
        if not folder:
            messagebox.showinfo("안내", "먼저 사진 폴더를 선택하세요.")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("오류", f"폴더를 찾을 수 없습니다:\n{folder}")
            return
        old = {p["path"] for p in self.photos}
        top = self.photo_canvas.yview()[0]
        self._load_photos(quiet=True)
        now = {p["path"] for p in self.photos}
        self._new_paths = (now - old) | (self._new_paths & now)
        self._paint_photos()
        self.after(80, lambda: self.photo_canvas.yview_moveto(top))
        self.log(f"✓ 사진 목록 갱신: {len(now)}장 (추가 {len(now-old)}장, 삭제 {len(old-now)}장)")

    def _get_current_slot_suggested_name(self) -> str:
        """현재 선택된 칸의 엑셀 셀 마크업({{이미지:...}}) 또는 라벨을 바탕으로 저장 파일명을 추천한다.
        윈도우 파일명 금지 문자(: 등)는 '_'로 치환하고, 파일명이 중복되는 경우 끝에 _1, _2 등 숫자를 붙인다."""
        if self.selected_slot is None or self.selected_slot >= len(self.slots):
            return ""

        slot = self.slots[self.selected_slot]
        folder = self.link_folder.get().strip()

        # 1. 마크업/라벨에서 이름 추출 (예: {{이미지:2.기내압력 점검:현황}} -> 2.기내압력 점검:현황)
        raw = (slot.get("label") or "").strip()
        if not raw and slot.get("marker"):
            m = re.match(r"\{\{\s*이미지\s*(?::\s*([^}]*?))?\s*\}\}", str(slot.get("marker")).strip())
            if m and m.group(1):
                raw = m.group(1).strip()

        if not raw:
            raw = f"칸_{self.selected_slot + 1}"

        # 2. 윈도우 파일명 금지 문자 (\ / : * ? " < > |) 치환
        safe_name = re.sub(r'[\\/*?:"<>|]', '_', raw).strip(' .')
        if not safe_name:
            safe_name = f"칸_{self.selected_slot + 1}"

        # 확장자 분리
        if safe_name.lower().endswith((".jpg", ".jpeg", ".png")):
            stem = safe_name.rsplit(".", 1)[0]
            ext = "." + safe_name.rsplit(".", 1)[1]
        else:
            stem = safe_name
            ext = ".jpg"

        target_name = f"{stem}{ext}"

        # 3. 중복 확인 (사진 폴더 내 파일 및 이미 연결된 사진 기준)
        if folder and os.path.isdir(folder):
            existing_names = set()
            try:
                existing_names = {f.casefold() for f in os.listdir(folder)}
            except Exception:
                pass
            for assigned in self.slot_assign.values():
                if assigned:
                    existing_names.add(os.path.basename(assigned).casefold())

            # 중복이 있으면 끝에 _1, _2, _3 ... 숫자를 붙임
            if target_name.casefold() in existing_names:
                n = 1
                while f"{stem}_{n}{ext}".casefold() in existing_names:
                    n += 1
                target_name = f"{stem}_{n}{ext}"

        return target_name

    def _open_camera_capture(self):
        """실시간 카메라 촬영 대화상자를 연다."""
        folder = self.link_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            ans = messagebox.askyesno("사진 폴더 필요", "촬영한 사진을 저장할 폴더가 아직 선택되지 않았습니다.\n지금 사진 폴더를 선택하시겠습니까?")
            if ans:
                self._browse_link_folder()
                folder = self.link_folder.get().strip()
                if not folder or not os.path.isdir(folder):
                    return
            else:
                return

        sug_name = self._get_current_slot_suggested_name()
        self._camera_dialog = CameraCaptureDialog(
            self,
            folder,
            on_photo_saved=self._on_photo_captured,
            suggested_name=sug_name,
            auto_assign_var=self._cam_auto_assign,
            get_next_name=self._get_current_slot_suggested_name,
        )

    def _on_photo_captured(self, saved_path: str):
        """카메라로 새 사진이 촬영·저장되었을 때 사진 목록을 갱신하고 선택된 칸에 자동 연결한다."""
        self._refresh_photos()
        norm_saved = os.path.normpath(saved_path)
        idx = next((i for i, p in enumerate(self.photos) if os.path.normpath(p["path"]) == norm_saved), None)
        if idx is not None and self.selected_slot is not None and self._cam_auto_assign.get():
            cur_slot = self.selected_slot
            self._assign(cur_slot, idx, advance=True)
            self.log(f"📷 사진 촬영 완료 및 연결: {os.path.basename(saved_path)} → {cur_slot + 1}번 칸")
        else:
            self.log(f"📷 사진 촬영 및 폴더 저장 완료: {os.path.basename(saved_path)}")

    def _refresh_all(self):
        """사진 폴더와 엑셀을 함께 다시 읽는다. 기존 사용자 연결은 셀 주소 기준으로 보존한다."""
        excel = self.link_excel.get().strip()
        folder = self.link_folder.get().strip()
        if folder and os.path.isdir(folder):
            old_paths = {p["path"] for p in self.photos}
            self._load_photos(quiet=True)
            now = {p["path"] for p in self.photos}
            self._new_paths = (now - old_paths) | (self._new_paths & now)
            self._paint_photos()
        if not excel or not os.path.isfile(excel):
            if folder:
                self.log("✓ 사진 폴더를 갱신했습니다. 엑셀 파일을 선택하면 엑셀 목록도 갱신됩니다.")
            else:
                messagebox.showinfo("안내", "엑셀 파일 또는 사진 폴더를 먼저 선택하세요.")
            return
        saved = {}
        for idx, path in self.slot_assign.items():
            if idx < len(self.slots):
                sl = self.slots[idx]
                saved[(sl.get("sheet"), sl.get("addr"))] = path
        self.log("[새로 고침] 엑셀의 사진 위치와 사진 폴더를 다시 확인합니다...")
        def refreshed(result):
            self._populate_slots(result)
            key_to_idx = {(sl.get("sheet"), sl.get("addr")): i for i, sl in enumerate(self.slots)}
            for key, path in saved.items():
                idx = key_to_idx.get(key)
                if idx is not None and (not self.photos or path in self.photo_index):
                    self.slot_assign[idx] = path
                    self._paint_slot(idx)
            self._update_summary()
            self.log("✓ 엑셀 및 사진 폴더 새로 고침 완료 (기존 연결 유지)")
        self.run_task("엑셀 새로 고침", scan_image_slots, excel, self.log, on_done=refreshed)

    def _install_pillow(self):
        if getattr(sys, "frozen", False):
            messagebox.showinfo("안내", "exe 버전에서는 여기서 설치할 수 없습니다.\nbuild_exe.py 로 exe를 다시 만들어 주세요.")
            return
        self.log("[Pillow 설치 중] 잠시만 기다려 주세요...")

        def work():
            r = subprocess.run([sys.executable, "-m", "pip", "install", "pillow"], capture_output=True, text=True,
                               creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            self.call_ui(lambda: self._after_pillow_install(r))

        threading.Thread(target=work, daemon=True).start()

    def _after_pillow_install(self, r):
        global Image, ImageOps, ImageTk, PIL_OK, PIL_ERR
        if r.returncode != 0:
            self.log("❌ Pillow 설치 실패:\n" + (r.stderr or r.stdout)[-500:])
            return
        import importlib
        importlib.invalidate_caches()
        try:
            from PIL import Image, ImageOps, ImageTk
            PIL_OK, PIL_ERR = True, ""
        except Exception as e:
            self.log(f"❌ 설치는 되었지만 불러오지 못했습니다: {e}\n   프로그램을 껐다가 다시 실행해 보세요.")
            return
        self.pil_row.pack_forget()
        self.log("✓ Pillow 설치 완료")
        if self.link_folder.get():
            self._load_photos()

    # 칸 목록
    def _populate_slots(self, result):
        if isinstance(result, tuple) and len(result) == 2:
            slots, previews = result
        else:
            slots, previews = result, []
        self._slot_of_widget.clear()
        self.slots, self.slot_widgets, self.slot_assign, self.slot_delete, self.selected_slot = slots, [], {}, set(), None
        if not slots:
            self._show_slot_msg("이 엑셀에서 {{이미지}} 표식이나 기존 사진을 찾지 못했습니다.\n사진이 들어갈 칸에 {{이미지}} 라고 적어 주세요.")
            self._update_summary()
        else:
            for w in self.slot_inner.winfo_children():
                w.destroy()
            for i, s in enumerate(slots):
                self._build_slot_row(i, s)
            self._select_slot(0)
            self._paint_photos()
            self._update_summary()
            had = sum(1 for sl in slots if sl.get("existing"))
            self.log(f"✓ 사진을 넣을 칸 {len(slots)}개를 찾았습니다." + (f" (이미 사진이 있는 칸 {had}개 — 새 사진을 연결하면 교체됩니다)" if had else ""))
            if self.photos:
                self._auto_match_and_fill()
        if previews:
            self._populate_previews(previews)

    def _slot_photo(self, path: str):
        if not PIL_OK:
            return None
        cached = self._slot_thumb_cache.get(path)
        if cached is not None:
            return cached
        try:
            im = Image.open(path)
            im.draft("RGB", (80, 80))
            im = ImageOps.exif_transpose(im).convert("RGB")
            im.thumbnail((42, 42))
            ph = ImageTk.PhotoImage(im)
            self._slot_thumb_cache[path] = ph
            return ph
        except Exception:
            return None

    def _build_slot_row(self, i: int, s: dict):
        r, c = divmod(i, 2)
        row = tk.Frame(self.slot_inner, bg=C["card"], highlightthickness=1, highlightbackground=C["line"], cursor="hand2")
        row.grid(row=r, column=c, padx=3, pady=3, sticky="new")

        # 번호 배지
        num = tk.Label(row, text=str(i + 1), font=(FONT, 8, "bold"), fg="white", bg=C["primary"], width=3, pady=2)
        num.pack(side=tk.LEFT, padx=(5, 5), pady=6)

        # 중앙 텍스트 (라벨, 시트명, 연결 상태)
        txt = tk.Frame(row, bg=C["card"])
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=4)
        lbl_text = s["label"] or ("기존 사진" if s["kind"] == "pic" else f"위치 {i+1}")
        t1 = tk.Label(txt, text=lbl_text, font=(FONT, 8, "bold"), fg=C["text"], bg=C["card"], anchor="w")
        t2 = tk.Label(txt, text=f'{s["sheet"]}', font=(FONT, 7), fg=C["mut"], bg=C["card"], anchor="w")
        t3 = tk.Label(txt, text="비어 있음", font=(FONT, 7, "bold"), fg=C["mut"], bg=C["card"], anchor="w")
        t1.pack(fill=tk.X)
        t2.pack(fill=tk.X)
        t3.pack(fill=tk.X)

        # 오른쪽 썸네일 + 삭제 버튼
        right = tk.Frame(row, bg=C["card"])
        right.pack(side=tk.RIGHT, padx=(2, 4), pady=4)

        # 썸네일 고정 크기(40x40) 박스 - 이미지가 없어도 칸 높이가 커지지 않도록 방지
        thumb_box = tk.Frame(right, width=40, height=40, bg=C["card"])
        thumb_box.pack(side=tk.LEFT, padx=(0, 2))
        thumb_box.pack_propagate(False)

        thumb = tk.Label(thumb_box, bg=C["card"])
        thumb.pack(fill=tk.BOTH, expand=True)
        thumb.bind("<Enter>", lambda e, k=i: self._hover_enter_slot(k, e))
        thumb.bind("<Leave>", lambda e: self._hover_leave())

        clear = tk.Label(right, text="✕", font=(FONT, 9), fg="#B0B7C6", bg=C["card"], cursor="hand2", padx=3)
        clear.pack(side=tk.LEFT)
        clear.bind("<Button-1>", lambda e, k=i: self._clear_slot(k))
        clear.bind("<Enter>", lambda e, w=clear: w.config(fg=C["bad"]))
        clear.bind("<Leave>", lambda e, w=clear: w.config(fg="#B0B7C6"))

        widgets = (row, num, txt, t1, t2, t3, right, thumb_box, thumb, clear)
        for w in widgets:
            self._slot_of_widget[str(w)] = i
            if w is not clear:
                w.bind("<Button-1>", lambda e, k=i: self._select_slot(k))

        thumb.bind("<Button-1>", lambda e, k=i: self._click_preview_slot(k), add="+")
        self.slot_widgets.append({
            "row": row, "state": t3, "thumb": thumb,
            "bgs": [row, txt, t1, t2, t3, right, thumb_box, thumb, clear]
        })
        self._paint_slot(i)

    def _paint_slot(self, i: int):
        if i >= len(self.slot_widgets):
            return
        sw = self.slot_widgets[i]
        selected = i == self.selected_slot
        bg = "#EEF2FF" if selected else C["card"]
        for w in sw["bgs"]:
            w.config(bg=bg)
        sw["row"].config(highlightbackground=C["primary"] if selected else C["line"])
        path = self.slot_assign.get(i)
        if path:
            photo = self._slot_photo(path)
            fname = os.path.basename(path)
            if len(fname) > 12:
                fname = fname[:10] + "…"
            sw["state"].config(text=f"→ {fname}", fg=C["ok"])
            sw["thumb"].config(image=photo or "", text="" if photo else "🖼")
            sw["thumb"].image = photo
        elif i in self.slot_delete:
            sw["state"].config(text="🗑 기존 사진 삭제 예정", fg=C["bad"])
            sw["thumb"].config(image="", text="❌", fg=C["bad"], font=(FONT, 9, "bold"))
            sw["thumb"].image = None
        else:
            if self.slots[i].get("existing"):
                sw["state"].config(text="📷 기존 사진", fg=C["primary"])
            else:
                sw["state"].config(text="비어 있음", fg=C["mut"])
            old = self._old_photo(i) if self.slots[i].get("existing") else None
            sw["thumb"].config(image=old or "", text="")
            sw["thumb"].image = old

    def _select_slot(self, i: int):
        old, self.selected_slot = self.selected_slot, i
        if old is not None:
            self._paint_slot(old)
        self._paint_slot(i)
        self._ensure_visible(i)
        if getattr(self, "_camera_dialog", None) is not None:
            try:
                if self._camera_dialog.winfo_exists():
                    new_name = self._get_current_slot_suggested_name()
                    if new_name:
                        self._camera_dialog.name_var.set(new_name)
            except Exception:
                self._camera_dialog = None

    def _ensure_visible(self, i: int):
        self.update_idletasks()
        total = max(1, self.slot_inner.winfo_height())
        top, bottom = self.slot_canvas.yview()
        frac = self.slot_widgets[i]["row"].winfo_y() / total
        if frac < top or frac > bottom - 0.08:
            self.slot_canvas.yview_moveto(max(0.0, frac - 0.03))

    # 사진 목록
    def _load_photos(self, quiet=False):
        folder = self.link_folder.get()
        self._cancel_hover()
        self._hide_preview()
        self._gen += 1
        gen = self._gen
        self.photos, self.photo_index = [], {}
        try:
            paths = [os.path.normpath(os.path.join(folder, f)) for f in os.listdir(folder) if f.lower().endswith(IMG_EXT)]
        except OSError as e:
            self._show_photo_msg(f"폴더를 읽을 수 없습니다: {e}")
            return
        paths.sort(key=lambda p: natural_key(os.path.basename(p)))
        self.slot_assign = {k: v for k, v in self.slot_assign.items() if v in set(paths)}
        if not paths:
            self._show_photo_msg("이 폴더에 사진(jpg, png, bmp, gif, tif)이 없습니다.")
            self._repaint_all()
            return
        for w in self.photo_inner.winfo_children():
            w.destroy()
        for idx, p in enumerate(paths):
            self._build_photo_cell(idx, p)
        self.photo_index = {p["path"]: i for i, p in enumerate(self.photos)}
        self._repaint_all()
        if self.slots:
            self._auto_match_and_fill()
        if not quiet:
            self.log(f"✓ 사진 {len(paths)}장을 불러왔습니다.")
        if not PIL_OK:
            self.pil_row.pack(fill=tk.X, pady=3)
            if getattr(sys, "frozen", False):
                self.log("❌ 미리보기를 만들 수 없습니다. 이 exe에는 Pillow가 포함되어 있지 않습니다.\n"
                         "   exe를 만드는 컴퓨터에서 build_exe.py를 다시 실행해 새 exe로 교체하세요.")
            else:
                self.log("❌ 미리보기를 만들 수 없습니다 (Pillow 불러오기 실패: " + PIL_ERR + ")\n"
                         "   위쪽의 [Pillow 설치] 버튼을 누르거나, 터미널에서 아래 명령을 실행하세요.\n"
                         f'   "{sys.executable}" -m pip install pillow')
            return
        self.pil_row.pack_forget()

        def work():
            reported = False
            for i, p in enumerate(paths):
                if gen != self._gen:
                    return
                try:
                    mtime = os.path.getmtime(p)
                except OSError:
                    mtime = 0
                cached = self._thumb_cache.get(p)
                if cached and cached[0] == mtime:  # 바뀌지 않은 사진은 다시 만들지 않음
                    self.call_ui(lambda i=i, im=cached[1], g=gen: self._set_thumb(g, i, im))
                    continue
                try:
                    im = Image.open(p)
                    im.draft("RGB", (THUMB * 2, THUMB * 2))  # 큰 JPEG를 빠르게 축소 로딩
                    im = ImageOps.exif_transpose(im).convert("RGB")
                    im.thumbnail((THUMB, THUMB))
                except Exception as e:
                    im = None
                    if not reported:
                        reported = True
                        self.log(f"❌ 미리보기 생성 실패 ({os.path.basename(p)}): {e}")
                if im is not None:
                    self._thumb_cache[p] = (mtime, im)
                self.call_ui(lambda i=i, im=im, g=gen: self._set_thumb(g, i, im))

        threading.Thread(target=work, daemon=True).start()

    def _cols(self) -> int:
        w = self.photo_canvas.winfo_width()
        w = w if w > 60 else 380
        return max(1, (w - 8) // (self.thumb_size + 16))

    def _short(self, name: str) -> str:
        n = max(9, self.thumb_size // 6)
        return name if len(name) <= n else name[: n // 2 - 1] + "…" + name[-(n // 2):]

    def _build_photo_cell(self, idx: int, path: str):
        r, c = divmod(idx, self._cols())
        size = self.thumb_size
        cell = tk.Frame(self.photo_inner, bg=C["card"], highlightthickness=2, highlightbackground=C["line"], cursor="hand2")
        cell.grid(row=r, column=c, padx=4, pady=4, sticky="n")
        name = os.path.basename(path)
        img = tk.Label(cell, image=self._blank, compound="center", text="…" if PIL_OK else "미리보기\n없음",
                       bg="#F3F4F8", fg=C["mut"], width=size, height=size, font=(FONT, 8))
        img.pack()
        nm = tk.Label(cell, text=self._short(name), font=(FONT, 8), fg=C["text"], bg=C["card"])
        nm.pack()
        badge = tk.Label(cell, text=" ", font=(FONT, 8, "bold"), fg=C["ok"], bg=C["card"])
        badge.pack()
        for w in (cell, img, nm, badge):
            w.bind("<ButtonPress-1>", lambda e, i=idx: self._photo_press(e, i))
            w.bind("<B1-Motion>", self._photo_motion)
            w.bind("<ButtonRelease-1>", self._photo_release)
            w.bind("<Enter>", lambda e, i=idx: self._hover_enter(i, e))
            w.bind("<Leave>", lambda e: self._hover_leave())
        self.photos.append({"path": path, "name": name, "cell": cell, "img": img, "nm": nm, "badge": badge,
                            "pil": None, "tk": None, "tk_size": None, "small": None})

    def _photo_tk(self, p: dict):
        """현재 선택한 크기에 맞는 화면용 이미지"""
        if p["pil"] is None:
            return None
        if p["tk_size"] != self.thumb_size:
            im = p["pil"].copy()
            im.thumbnail((self.thumb_size, self.thumb_size))
            p["tk"], p["tk_size"] = ImageTk.PhotoImage(im), self.thumb_size
        return p["tk"]

    def _set_thumb(self, gen: int, i: int, im):
        if gen != self._gen or i >= len(self.photos):
            return
        p = self.photos[i]
        if im is None:
            p["img"].config(text="미리보기\n불가")
            return
        p["pil"] = im
        p["img"].config(image=self._photo_tk(p), text="", width=self.thumb_size, height=self.thumb_size)
        for k, v in self.slot_assign.items():
            if v == p["path"]:
                self._paint_slot(k)

    def _old_photo(self, i: int):
        """칸에 이미 들어 있는 기존 사진의 작은 미리보기"""
        sl = self.slots[i]
        if not PIL_OK or sl.get("old_pil") is None:
            return None
        if "old_tk" not in sl:
            im = sl["old_pil"].copy()
            im.thumbnail((38, 38))
            sl["old_tk"] = ImageTk.PhotoImage(im)
        return sl["old_tk"]

    def _small_photo(self, path: str):
        i = self.photo_index.get(path)
        if i is None or not PIL_OK:
            return None
        p = self.photos[i]
        if p["small"] is None and p["pil"] is not None:
            small = p["pil"].copy()
            small.thumbnail((56, 56))
            p["small"] = ImageTk.PhotoImage(small)
        return p["small"]

    # 사진 크기 바꾸기 / 창 크기에 맞춰 다시 배치
    def _paint_size_chips(self):
        for name, chip in self.size_chips.items():
            on = SIZES[name] == self.thumb_size
            chip.config(bg=C["primary"] if on else "#EEF1FB", fg="white" if on else C["primary"])

    def _set_size(self, name: str):
        self.thumb_size = SIZES[name]
        self._paint_size_chips()
        for p in self.photos:
            p["img"].config(width=self.thumb_size, height=self.thumb_size)
            if p["pil"] is not None:
                p["img"].config(image=self._photo_tk(p))
            p["nm"].config(text=self._short(p["name"]))
        self._relayout_photos()

    def _schedule_relayout(self):
        if self._relayout_job:
            self.after_cancel(self._relayout_job)
        self._relayout_job = self.after(150, self._relayout_photos)

    def _relayout_photos(self):
        self._relayout_job = None
        cols = self._cols()
        for idx, p in enumerate(self.photos):
            r, c = divmod(idx, cols)
            p["cell"].grid(row=r, column=c, padx=4, pady=4, sticky="n")

    # 마우스를 올리면 큰 미리보기
    def _hover_enter(self, idx: int, e):
        if idx < len(self.photos):
            self._start_hover(("photo", idx), self.photos[idx]["path"], e)

    def _hover_enter_slot(self, k: int, e):
        path = self.slot_assign.get(k)
        if path:
            self._start_hover(("slot", k, path), path, e)
        elif PIL_OK and self.slots[k].get("old_pil") is not None:  # 연결 전이면 현재 들어 있는 사진을 크게 보여 줌
            self._start_hover(("old", k), f"@old{k}", e)

    def _click_preview_slot(self, k: int):
        """엑셀에 이미 있는 사진을 클릭하면 큰 미리보기 창으로 표시한다."""
        if not PIL_OK or k >= len(self.slots):
            return
        path = self.slot_assign.get(k)
        if path:
            key, preview_path = ("slot", k, path), path
        elif self.slots[k].get("old_pil") is not None:
            key, preview_path = ("old", k), f"@old{k}"
        else:
            return
        self._cancel_hover()
        self._cancel_hide()
        self._show_preview(key, preview_path, self.winfo_pointerx(), self.winfo_pointery())

    def _start_hover(self, key, path: str, e):
        self._cancel_hide()
        if self._preview is not None and self._hover_key == key:
            return
        self._cancel_hover()
        x, y = e.x_root, e.y_root
        self._hover_job = self.after(250, lambda: self._show_preview(key, path, x, y))

    def _hover_leave(self):
        self._cancel_hover()
        self._cancel_hide()
        self._hide_job = self.after(80, self._hide_preview)

    def _cancel_hover(self):
        if self._hover_job:
            self.after_cancel(self._hover_job)
            self._hover_job = None

    def _cancel_hide(self):
        if self._hide_job:
            self.after_cancel(self._hide_job)
            self._hide_job = None

    def _hide_preview(self):
        self._hide_job = None
        if self._preview is not None:
            self._preview.destroy()
            self._preview, self._hover_key = None, None

    def _big_photo(self, path: str):
        if path in self._big_cache:
            return self._big_cache[path]
        try:
            target_size = max(420, min(BIG, self.winfo_screenheight() - 100))  # 큰 미리보기 목표 크기 (화면 높이에 맞춰 자동 조절)
            if path.startswith("@old"):
                index = int(path[4:])
                im = self.slots[index]["old_pil"].copy()
            else:
                im = Image.open(path)
                im.draft("RGB", (target_size * 2, target_size * 2))
                im = ImageOps.exif_transpose(im).convert("RGB")

            # 가로/세로 비율을 유지하면서 목표 크기에 맞춰 확대/축소 (작은 이미지도 크게 확대)
            w, h = im.size
            if w > 0 and h > 0:
                scale = target_size / max(w, h)
                new_w, new_h = max(1, int(w * scale)), max(1, int(h * scale))
                im = im.resize((new_w, new_h), Image.Resampling.LANCZOS)
            ph = ImageTk.PhotoImage(im)
        except Exception as e:
            self.log(f"❌ 큰 미리보기 생성 실패 ({os.path.basename(path)}): {e}")
            return None
        if len(self._big_cache) >= 16:  # 오래된 것부터 정리
            self._big_cache.pop(next(iter(self._big_cache)))
        self._big_cache[path] = ph
        return ph

    def _show_preview(self, key, path: str, x: int, y: int):
        self._hover_job = None
        if not PIL_OK:
            return
        ph = self._big_photo(path)
        if ph is None:
            return
        self._hide_preview()
        pv = tk.Toplevel(self)
        pv.overrideredirect(True)
        try:
            pv.attributes("-topmost", True)
        except tk.TclError:
            pass
        frame = tk.Frame(pv, bg="#111827", padx=3, pady=3)
        frame.pack()
        tk.Label(frame, image=ph, bg="#111827").pack()
        tk.Label(frame, text="현재 들어 있는 사진" if path.startswith("@old") else os.path.basename(path), fg="white", bg="#111827", font=(FONT, 9)).pack(pady=(2, 0))
        pv.update_idletasks()
        w, h = pv.winfo_reqwidth(), pv.winfo_reqheight()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        px = x - w - 24 if x > sw / 2 else x + 24  # 커서가 있는 쪽의 반대편에 표시
        pv.geometry(f"+{max(0, min(px, sw - w))}+{max(0, min(y - h // 2, sh - h))}")
        self._preview, self._hover_key = pv, key

    def _paint_photos(self):
        used = {v: k for k, v in self.slot_assign.items()}
        for p in self.photos:
            k = used.get(p["path"])
            if k is not None:
                p["badge"].config(text=f"✓ {k + 1}번 칸", fg=C["ok"])
            elif p["path"] in self._new_paths:
                p["badge"].config(text="NEW", fg=C["primary"])
            else:
                p["badge"].config(text=" ")
            p["cell"].config(highlightbackground=C["ok"] if k is not None else C["line"])

    def _repaint_all(self):
        for k in range(len(self.slot_widgets)):
            self._paint_slot(k)
        self._paint_photos()
        self._update_summary()

    def _update_summary(self):
        n, m, total = len(self.slots), len(self.slot_assign), len(self.photos)
        del_cnt = len(self.slot_delete)
        del_str = f" · 🗑 삭제 {del_cnt}개" if del_cnt else ""
        self.link_summary.config(text=f"칸 {n}개 · 연결 {m}개{del_str} · 비어 있는 칸 {n - m}개 · 사진 {total}장 (안 쓴 사진 {total - m}장)")

    # 연결 동작
    def _assign(self, slot: int, photo_idx: int, advance=False):
        path = self.photos[photo_idx]["path"]
        self.slot_delete.discard(slot)  # 새 사진 연결 시 삭제 예약 자동 해제
        for k in [k for k, v in self.slot_assign.items() if v == path and k != slot]:
            del self.slot_assign[k]  # 같은 사진은 한 칸에만: 다른 칸의 연결은 해제
            self._paint_slot(k)
        self.slot_assign[slot] = path
        self._paint_slot(slot)
        self._paint_photos()
        self._update_summary()
        if advance:
            order = list(range(slot + 1, len(self.slots))) + list(range(0, slot))
            nxt = next((k for k in order if k not in self.slot_assign and k not in self.slot_delete), None)
            if nxt is not None:
                self._select_slot(nxt)

    def _clear_slot(self, i: int):
        # 1. 새 사진이 연결되어 있던 칸이면 -> 새 사진 연결 해제
        if i in self.slot_assign:
            self.slot_assign.pop(i, None)
        # 2. 기존 사진이 있는 칸이면 -> 삭제 예약 토글 (클릭할 때마다 삭제 예약 <-> 기존 유지 전환)
        elif self.slots[i].get("existing"):
            if i in self.slot_delete:
                self.slot_delete.remove(i)
                self.log(f"[안내] {i + 1}번 칸의 기존 사진 삭제를 취소했습니다.")
            else:
                self.slot_delete.add(i)
                self.log(f"[안내] {i + 1}번 칸 기존 사진 삭제 예약 (다시 ✕를 누르면 취소됩니다)")
        self._paint_slot(i)
        self._paint_photos()
        self._update_summary()

    def _clear_all_links(self):
        self.slot_assign.clear()
        self.slot_delete.clear()
        self._repaint_all()

    def _auto_match_and_fill(self):
        """표식에 파일명이 지정된 칸은 이름으로 연결하고, 이름 없는 칸은 남은 사진을 순서대로 연결."""
        if not self.slots or not self.photos:
            messagebox.showinfo("안내", "대상 엑셀과 사진 폴더를 먼저 선택하세요.")
            return
        by_name = {}
        for photo in self.photos:
            name = os.path.basename(photo["path"]).casefold()
            stem = os.path.splitext(name)[0]
            by_name[name] = photo["path"]
            by_name.setdefault(stem, photo["path"])
        matched = 0
        used = set(self.slot_assign.values())
        for i, slot in enumerate(self.slots):
            if i in self.slot_assign:
                continue
            label = (slot.get("label") or "").strip().replace("\\", "/").split("/")[-1]
            if not label:
                continue
            key = label.casefold()
            path = by_name.get(key) or by_name.get(os.path.splitext(key)[0])
            if path and path not in used:
                self.slot_assign[i] = path
                used.add(path)
                matched += 1
        free = iter([p["path"] for p in self.photos if p["path"] not in used])
        sequential = 0
        for i, slot in enumerate(self.slots):
            if i in self.slot_assign or (slot.get("label") or "").strip():
                continue
            path = next(free, None)
            if path is None:
                break
            self.slot_assign[i] = path
            sequential += 1
        self._repaint_all()
        self.log(f"✓ 통합 자동 연결 완료: 파일명 일치 {matched}개, 순서 연결 {sequential}개")

    # 끌어다 놓기 (클릭과 구분)
    def _photo_press(self, e, idx: int):
        self._cancel_hover()
        self._hide_preview()
        self._drag = {"idx": idx, "x": e.x_root, "y": e.y_root, "ghost": None}

    def _photo_motion(self, e):
        d = self._drag
        if not d:
            return
        if d["ghost"] is None and abs(e.x_root - d["x"]) + abs(e.y_root - d["y"]) > 8:
            d["ghost"] = self._make_ghost(d["idx"])
        if d["ghost"] is not None:
            d["ghost"].geometry(f"+{e.x_root + 12}+{e.y_root + 12}")

    def _make_ghost(self, idx: int):
        g = tk.Toplevel(self)
        g.overrideredirect(True)
        try:
            g.attributes("-topmost", True)
            g.attributes("-alpha", 0.85)
        except tk.TclError:
            pass
        p = self.photos[idx]
        if p["tk"] is not None:
            tk.Label(g, image=p["tk"], bd=2, relief="solid").pack()
        else:
            tk.Label(g, text=p["name"], bd=2, relief="solid", padx=8, pady=4).pack()
        return g

    def _photo_release(self, e):
        d, self._drag = self._drag, None
        if not d:
            return
        if d["ghost"] is not None:  # 끌어다 놓기
            d["ghost"].destroy()
            try:
                w = self.winfo_containing(e.x_root, e.y_root)
            except Exception:
                w = None
            while w is not None and str(w) not in self._slot_of_widget:
                w = w.master
            if w is not None:
                self._assign(self._slot_of_widget[str(w)], d["idx"])
        elif self.selected_slot is not None:  # 클릭 → 선택된 칸에 연결 후 다음 빈 칸으로
            self._assign(self.selected_slot, d["idx"], advance=True)
        else:
            self.log("[안내] 먼저 왼쪽에서 사진을 넣을 칸을 선택하세요.")

    # ── 사진 삽입 (미리보기 갱신) 및 저장 ──
    def _apply_photo_insert(self):
        excel = self.link_excel.get().strip()
        if not excel or not os.path.isfile(excel):
            messagebox.showerror("입력 오류", "대상 엑셀 파일을 먼저 선택해주세요.")
            return
        if not self.slot_assign and not self.slot_delete:
            messagebox.showerror("입력 오류", "연결되거나 삭제할 사진이 없습니다.\n사진을 연결하거나 ✕를 눌러 삭제를 지정한 뒤 [사진 삽입]을 누르세요.")
            return

        empty = sum(1 for i in range(len(self.slots)) if i not in self.slot_assign and not self.slots[i].get("existing") and i not in self.slot_delete)
        action_desc = []
        if self.slot_assign:
            action_desc.append(f"연결된 {len(self.slot_assign)}개 삽입")
        if self.slot_delete:
            action_desc.append(f"지정된 {len(self.slot_delete)}개 삭제")
        desc_str = " 및 ".join(action_desc)

        if empty and not messagebox.askyesno("확인", f"비어 있는 칸이 {empty}개 있습니다.\n{desc_str} 작업을 먼저 진행하고 미리보기를 갱신할까요?"):
            return

        items = [(self.slots[i], path) for i, path in sorted(self.slot_assign.items())]
        delete_slots = [self.slots[i] for i in sorted(self.slot_delete)]
        excel = os.path.normpath(excel)
        fit = self.link_fit.get()

        def done(result):
            ok, fail, staged_excel, previews = result
            self._staged_excel_path = staged_excel
            self._staged_items_hash = (hash(tuple(sorted(self.slot_assign.items()))), hash(tuple(sorted(self.slot_delete))))
            msg_parts = []
            if self.slot_assign:
                msg_parts.append(f"사진 {len(self.slot_assign)}개 삽입")
            if self.slot_delete:
                msg_parts.append(f"사진 {len(self.slot_delete)}개 삭제")
            self.log(f"\n✓ {' / '.join(msg_parts)} 완료! 우측 [3. 미리보기]에 반영되었습니다." + (f" (실패: {fail}개)" if fail else ""))
            self.log("👉 우측 미리보기를 확인하신 후, 만족스러우면 [저장] 버튼을 누르면 최종 엑셀 파일로 저장됩니다.")
            if previews:
                self._populate_previews(previews)

        self.run_task("사진 삽입/삭제 및 미리보기 갱신", lambda: insert_and_preview(excel, items, self.log, fit, delete_slots=delete_slots), on_done=done)

    def _save_inserted_file(self):
        excel = self.link_excel.get().strip()
        if not excel or not os.path.isfile(excel):
            messagebox.showerror("입력 오류", "대상 엑셀 파일을 먼저 선택해주세요.")
            return
        if not self.slot_assign and not self.slot_delete:
            messagebox.showerror("입력 오류", "연결되거나 삭제할 사진이 없습니다.\n사진을 칸에 연결하거나 삭제할 칸을 선택한 뒤 저장하세요.")
            return

        current_hash = (hash(tuple(sorted(self.slot_assign.items()))), hash(tuple(sorted(self.slot_delete))))
        if not self._staged_excel_path or not os.path.exists(self._staged_excel_path) or self._staged_items_hash != current_hash:
            ans = messagebox.askyesno("저장 안내", "사진 삽입/삭제 작업이 아직 실행되지 않았거나 변경되었습니다.\n지금 작업을 적용하고 바로 저장할까요?")
            if not ans:
                return
            self._direct_insert_and_save()
            return

        self._finish_save(self._staged_excel_path)

    def _direct_insert_and_save(self):
        excel = os.path.abspath(os.path.normpath(self.link_excel.get().strip()))
        orig_p = Path(excel)
        orig_dir = orig_p.parent.resolve()
        items = [(self.slots[i], path) for i, path in sorted(self.slot_assign.items())]
        delete_slots = [self.slots[i] for i in sorted(self.slot_delete)]
        fit = self.link_fit.get()
        replacing = sum(1 for i in self.slot_assign if self.slots[i].get("existing"))
        deleting = len(delete_slots)
        base = orig_p.stem if orig_p.stem.endswith("_사진삽입") else orig_p.stem + "_사진삽입"
        out_p = orig_dir / f"{base}{orig_p.suffix}"
        n = 2
        while out_p.exists():
            out_p = orig_dir / f"{base}_{n}{orig_p.suffix}"
            n += 1
        out = str(out_p.resolve())
        if replacing or deleting:
            warn_msg = []
            if replacing:
                warn_msg.append(f"기존 사진 {replacing}개 교체")
            if deleting:
                warn_msg.append(f"기존 사진 {deleting}개 삭제")
            ans = messagebox.askyesnocancel(
                "저장 방식", f"{', '.join(warn_msg)}를 진행합니다.\n\n[예]  원본 파일에 덮어써서 교체\n"
                            f"[아니요]  새 파일로 저장 (원본과 같은 폴더에 생성)\n    → {out_p.name}\n[취소]  중단")
            if ans is None:
                return
            if ans:
                out = str(orig_p.resolve())

        def done(res):
            ok, fail, out_path = res
            self.log(f"\n✓ 저장 완료 — 성공 {ok}개" + (f", ❌ 실패 {fail}개" if fail else ""))
            self.log(f"✓ 원본 파일과 동일한 위치에 저장되었습니다:")
            self.log(f"    → {out_path}")
            try:
                open_with_default_app(out_path)
            except Exception as e:
                self.log(f"(자동 열기 실패: {e})")
            messagebox.showinfo("저장 완료", f"원본 엑셀과 같은 폴더에 성공적으로 저장되었습니다!\n\n저장 경로:\n{out_path}")

        self.run_task("사진 삽입/삭제 및 저장", lambda: insert_images_by_slots(excel, items, out, self.log, fit, delete_slots=delete_slots), on_done=done)

    def _finish_save(self, staged_path: str):
        import shutil
        excel = os.path.abspath(os.path.normpath(self.link_excel.get().strip()))
        orig_p = Path(excel)
        orig_dir = orig_p.parent.resolve()
        replacing = sum(1 for i in self.slot_assign if self.slots[i].get("existing"))
        deleting = len(self.slot_delete)
        base = orig_p.stem if orig_p.stem.endswith("_사진삽입") else orig_p.stem + "_사진삽입"
        out_p = orig_dir / f"{base}{orig_p.suffix}"
        n = 2
        while out_p.exists():
            out_p = orig_dir / f"{base}_{n}{orig_p.suffix}"
            n += 1
        out = str(out_p.resolve())
        if replacing or deleting:
            warn_msg = []
            if replacing:
                warn_msg.append(f"기존 사진 {replacing}개 교체")
            if deleting:
                warn_msg.append(f"기존 사진 {deleting}개 삭제")
            ans = messagebox.askyesnocancel(
                "저장 방식", f"{', '.join(warn_msg)}가 적용된 상태입니다.\n\n[예]  원본 파일에 덮어써서 교체\n"
                            f"[아니요]  새 파일로 저장 (원본과 같은 폴더에 생성)\n    → {out_p.name}\n[취소]  중단")
            if ans is None:
                return
            if ans:
                out = str(orig_p.resolve())

        try:
            shutil.copyfile(staged_path, out)
            self.log(f"\n✓ 엑셀 저장 완료 (원본과 동일 폴더):")
            self.log(f"    → {out}")
            try:
                open_with_default_app(out)
            except Exception as e:
                self.log(f"(자동 열기 실패: {e})")
            messagebox.showinfo("저장 완료", f"원본 엑셀과 같은 폴더에 성공적으로 저장되었습니다!\n\n저장 경로:\n{out}")
        except Exception as e:
            self.log(f"❌ 파일 저장 실패: {e}")
            messagebox.showerror("저장 오류", f"파일을 저장할 수 없습니다: {e}")

    # ── 7. 환경 점검 (v16 신규) ──
    def _page_env(self):
        page, box = self._new_page()
        area = tk.Frame(box, bg=C["card"])
        area.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        _, inner = make_scroll_area(area, C["card"])
        head = tk.Frame(inner, bg=C["card"])
        head.pack(fill=tk.X, padx=22, pady=(16, 6))
        self.env_summary = lbl(head, "점검 중...", 12, True)
        self.env_summary.pack(side=tk.LEFT)
        self.env_frame = tk.Frame(inner, bg=C["card"])
        self.env_frame.pack(fill=tk.X, padx=22)
        btns = tk.Frame(inner, bg=C["card"])
        btns.pack(fill=tk.X, padx=22, pady=12)
        RoundButton(btns, "다시 점검", self._run_env_check, kind="ghost").pack(side=tk.LEFT, padx=(0, 8))
        RoundButton(btns, "엑셀 연결 테스트", self._test_excel, kind="ghost").pack(side=tk.LEFT, padx=(0, 8))
        RoundButton(btns, "결과 복사", self._copy_env, kind="ghost").pack(side=tk.LEFT)
        lbl(inner, "상세 안내 — 다른 컴퓨터에서 사용하려면", 11, True).pack(anchor="w", padx=22, pady=(10, 4))
        guide = tk.Text(inner, height=32, wrap="word", bg=C["field"], fg=C["text"], relief="flat", font=(FONT, 10),
                        padx=14, pady=10, highlightthickness=1, highlightbackground=C["line"])
        guide.pack(fill=tk.X, padx=22, pady=(0, 20))
        guide.tag_config("h", font=(FONT, 10, "bold"), foreground=C["primary"], spacing1=10, spacing3=2)
        for tag, text in ENV_GUIDE:
            guide.insert(tk.END, text + "\n", tag)
        guide.config(state="disabled")
        return page

    def _render_env(self, items: list[dict]):
        for w in self.env_frame.winfo_children():
            w.destroy()
        colors = {"ok": C["ok"], "bad": C["bad"], "warn": "#D97706"}
        marks = {"ok": "✓", "bad": "❌", "warn": "⚠"}
        for it in items:
            row = tk.Frame(self.env_frame, bg=C["card"])
            row.pack(fill=tk.X, pady=3)
            tk.Label(row, text=marks[it["status"]], fg=colors[it["status"]], bg=C["card"],
                     font=(FONT, 11, "bold"), width=3).pack(side=tk.LEFT, anchor="n")
            col = tk.Frame(row, bg=C["card"])
            col.pack(side=tk.LEFT, fill=tk.X, expand=True)
            line = tk.Frame(col, bg=C["card"])
            line.pack(fill=tk.X)
            lbl(line, it["name"], 10, True, width=24, anchor="w").pack(side=tk.LEFT)
            lbl(line, it["detail"], 9, color=C["mut"], anchor="w", justify="left", wraplength=420).pack(side=tk.LEFT, fill=tk.X, expand=True)
            if it["fix"]:
                lbl(col, "→ " + it["fix"], 9, color=colors[it["status"]], anchor="w", justify="left", wraplength=560).pack(fill=tk.X, pady=(2, 0))
        bad = sum(i["status"] == "bad" for i in items)
        warn = sum(i["status"] == "warn" for i in items)
        if bad:
            self.env_summary.config(text=f"❌ 해결이 필요한 항목 {bad}개" + (f" · ⚠ 확인 {warn}개" if warn else ""), fg=C["bad"])
        elif warn:
            self.env_summary.config(text=f"⚠ 일부 기능에 필요한 항목 {warn}개를 확인하세요", fg="#D97706")
        else:
            self.env_summary.config(text="✓ 모든 항목이 정상입니다", fg=C["ok"])

    def _startup_env_check(self):
        """실행할 때 한 번 점검해서, 필수 항목이 부족하면 환경 점검 탭을 먼저 보여 준다."""
        self._env_items = check_environment()
        self._render_env(self._env_items)
        if any(i["status"] == "bad" for i in self._env_items):
            self._switch_tab("env")
            self.log("❌ 필수 항목이 부족합니다. '환경 점검' 탭의 안내를 확인하세요.")

    def _run_env_check(self):
        self._env_items = check_environment()
        self._render_env(self._env_items)
        marks = {"ok": "✓", "bad": "❌", "warn": "⚠"}
        for it in self._env_items:
            if it["status"] != "ok":
                self.log(f'{marks[it["status"]]} {it["name"]}: {it["detail"]}')
        if all(i["status"] == "ok" for i in self._env_items):
            self.log("✓ 환경 점검: 모든 항목이 정상입니다.")

    def _test_excel(self):
        self.run_task("엑셀 연결 테스트", test_excel_connection,
                      on_done=lambda v: self.log(f"✓ 엑셀 연결 성공 (엑셀 버전 {v})"))

    def _copy_env(self):
        marks = {"ok": "OK", "bad": "NG", "warn": "주의"}
        lines = [f'[{marks[i["status"]]}] {i["name"]} - {i["detail"]}' + (f' / 해결: {i["fix"]}' if i["fix"] else "")
                 for i in self._env_items]
        self.clipboard_clear()
        self.clipboard_append(f"ximgpro v{VERSION} 환경 점검\n" + "\n".join(lines))
        self.log("✓ 점검 결과를 복사했습니다. 메신저나 메일에 붙여넣어 전달할 수 있어요.")

    # ── 시작 시 설정 복원 ──
    def _load_initial_config(self):
        config = load_config()
        if not config:
            return
        path = config.get("target_excel_path", "") or config.get("link_excel_path", "")
        folder = config.get("link_folder_path", "")
        self.log("✓ 최근 설정을 불러왔습니다.")
        if path and os.path.exists(path):
            self._set_link_excel(path)
        if folder and os.path.isdir(folder):
            self._set_link_folder(folder)


if __name__ == "__main__":
    ExcelPhotoOneStopApp().mainloop()
