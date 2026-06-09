"""
xlsx_scanner.py — MetadataCompiler V1
Core parsing logic: Excel sheet → raw JSON

Layout thực tế của file spec:
  - ファイル名  : nằm ở col B (col=2), dữ liệu ngay dưới keyword row.
  - 出力条件   : keyword ở col E (col=5). Header ở dòng TIẾP THEO.
                  Header scan từ col B đến khi gặp ô trống (thực tế: col 5–13).
  - 通常処理   : keyword ở col E. Header CÙNG DÒNG keyword, bắt đầu từ col E.
                  col D (col=4) = 開発用チェック, col E = DB_Field, col F = Description.
  - 特殊処理   : keyword ở col E. Header ở dòng TIẾP THEO.
                  Đọc data cho đến khi 50 dòng liên tiếp trống.

Required 4 keywords (tất cả đều ở col E):
  出力条件, 通常処理, 特殊処理, ファイル名  ← ファイル名 thực ra ở col B
"""

import os
import json
import openpyxl
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────
REQUIRED_KEYWORDS = {"ファイル名", "出力条件", "通常処理", "特殊処理"}

COL_B = 2
COL_D = 4
COL_E = 5
COL_F = 6
COL_G = 7
COL_N = 14
COL_O = 15

EMPTY_STREAK_LIMIT = 50


# ─────────────────────────────────────────────────────────
# Helper utilities
# ─────────────────────────────────────────────────────────

def _cv(ws, row: int, col: int) -> str:
    """Return stripped string value of a cell, '' if None."""
    v = ws.cell(row=row, column=col).value
    return str(v).strip() if v is not None else ""


def _bool_str(s: str) -> bool:
    return s.upper() in ("TRUE", "1", "○", "〇", "✓", "✔")


GREY_ARGB = {
    "FFA9A9A9", "FF808080", "FFC0C0C0", "FFD3D3D3",
    "FFB0B0B0", "FF969696", "FFBFBFBF", "FF7F7F7F",
    "FFD9D9D9", "FFE0E0E0", "FFF2F2F2",
}

# Hex threshold — nếu cần thêm dải màu xám khác chỉ cần bổ sung vào tập trên.
# Hàm helper kiểm tra toàn bộ vùng cột của một row trong 特殊処理 có bị xám không.
def _is_special_row_grey(ws, row: int, col_start: int = COL_E, col_end: int = COL_N) -> bool:
    """Return True nếu ít nhất một ô trong vùng [col_start, col_end] của row được tô xám.
    
    Dùng cho 特殊処理: kiểm tra xem một dòng có phải bị tô xám hay không.
    Chỉ cần 1 ô chủ chốt (col E = col 5 = tên item) bị xám là đủ.
    """
    for col in range(col_start, col_end + 1):
        if _is_grey(ws.cell(row=row, column=col)):
            return True
    return False


def _is_grey(cell) -> bool:
    """Return True if the cell has an explicit grey background fill.
    
    NOTE: openpyxl quirk — when a cell has NO fill, fill_type may still be
    'solid' with fgColor.rgb = 'FF000000' (Excel internal default). This must
    NOT be treated as grey. Only known grey ARGB values or high-tint theme
    colours count as grey.
    """
    try:
        fill = cell.fill
        if fill is None or fill.fill_type in (None, "none"):
            return False
        fg = fill.fgColor
        if fg is None:
            return False
        if fg.type == "rgb":
            argb = fg.rgb.upper()
            # Explicitly exclude Excel defaults that mean "no real fill"
            if argb in ("00000000", "FFFFFFFF", "FF000000", "00FFFFFF"):
                return False
            if argb in GREY_ARGB:
                return True
            if len(argb) == 8 and int(argb[0:2], 16) != 0:
                r = int(argb[2:4], 16)
                g = int(argb[4:6], 16)
                b = int(argb[6:8], 16)
                # Grey = R≈G≈B, not white (avg<210), not black (avg>20)
                avg = (r + g + b) / 3
                diff = max(r, g, b) - min(r, g, b)
                return diff <= 20 and 20 < avg < 210
        if fg.type == "theme" and fg.tint <= -0.35:
            return True
    except Exception:
        pass
    return False


# ─────────────────────────────────────────────────────────
# Keyword discovery — scan any column 1-5
# ─────────────────────────────────────────────────────────

def _find_keywords(ws) -> dict:
    """
    Return {keyword: (row, col)} for the first occurrence of each keyword.
    Scan cols 1–5 for 出力条件/通常処理/特殊処理, col 1-3 for ファイル名.
    """
    found = {}
    max_row = min(ws.max_row, 2000)
    for row in range(1, max_row + 1):
        for col in range(1, 6):
            v = _cv(ws, row, col)
            if v in REQUIRED_KEYWORDS and v not in found:
                found[v] = (row, col)
        if len(found) == 4:
            break
    return found


def sheet_is_valid(ws) -> bool:
    kw = _find_keywords(ws)
    return REQUIRED_KEYWORDS.issubset(kw.keys())


# ─────────────────────────────────────────────────────────
# Section parsers
# ─────────────────────────────────────────────────────────

def _parse_filename_section(ws, kw_row: int, kw_col: int, stop_row: int) -> list:
    """
    ファイル名 は keyword のすぐ下のセルから始まる (keyword列そのまま)。
    kw_col列, kw_row+1 行から stop_row-1 まで読む。
    空行50連続で止まる。
    """
    results = []
    priority = 1
    empty_streak = 0

    for row in range(kw_row + 1, stop_row):
        v = _cv(ws, row, kw_col)
        if v:
            results.append({
                "key": f"ファイル名_{priority}",
                "priority": priority,
                "file_name": v,
            })
            priority += 1
            empty_streak = 0
        else:
            empty_streak += 1
            if empty_streak >= EMPTY_STREAK_LIMIT:
                break
    return results


def _parse_output_condition_section(ws, kw_row: int, stop_row: int) -> dict:
    """
    出力条件:
      - keyword row に keyword がある (col E)
      - Header は次の行 (kw_row + 1) の col 5 以降 (実際は col E~M)
      - Data は kw_row+2 から stop_row-1 まで
      - Grouping rules が適用される
    """
    header_row = kw_row + 1
    col_headers: dict[int, str] = {}
    # Scan from col B(2) to col N(14) for headers
    for col in range(COL_B, COL_N + 1):
        h = _cv(ws, header_row, col)
        if h:
            col_headers[col] = h

    data_start = header_row + 1
    data_end   = stop_row - 1

    raw_rows = []
    for row in range(data_start, data_end + 1):
        row_data = {}
        has_any = False
        for col, hdr in col_headers.items():
            v = _cv(ws, row, col)
            if v:
                row_data[hdr] = v
                has_any = True
        if has_any:
            raw_rows.append(row_data)

    sentences = _group_output_condition(raw_rows)

    return {
        "headers": {str(c): h for c, h in col_headers.items()},
        "raw_rows": raw_rows,
        "sentences": sentences,
    }


def _group_output_condition(raw_rows: list) -> list:
    """
    Grouping rules for 出力条件:

    R1 : Consecutive 紐づけ rows → start a group.
    R2 : After 紐づけ, if 処理==条件分岐 AND 備考==AND → keep accumulating
         until that condition breaks, then CLOSE the group immediately.
         The next row starts a new sentence (fresh 紐づけ block or standalone).
    R1b: If the group ended WITHOUT R2 (no 条件分岐+AND rows), absorb the
         very next non-紐づけ row to complete the sentence.
    R3 : Any other row → standalone sentence.

    Example (アカウント（新規オーナー）):
      [紐づけ, 紐づけ, 条件分岐+AND, 条件分岐+AND]  → grouped  (R2 closes)
      [紐づけ, 紐づけ, 条件分岐]                    → grouped  (R1b absorbs)
      [出力]                                        → single   (R3)

    Returns list of {type: "grouped"|"single", rows: [...]}
    """
    shori_key = "処理"
    biko_key  = "備考"
    sentences = []
    i = 0
    n = len(raw_rows)

    while i < n:
        row = raw_rows[i]
        shori = row.get(shori_key, "")

        if shori == "紐づけ":
            group = [row]
            j = i + 1
            closed_by_r2 = False  # True when a 条件分岐+AND run ended the group

            # Outer accumulation: 紐づけ rows first, then check R2
            while j < n:
                nr = raw_rows[j]
                ns = nr.get(shori_key, "")
                nb = nr.get(biko_key, "")

                if ns == "紐づけ":
                    # R1 — keep collecting 紐づけ rows
                    group.append(nr)
                    j += 1

                elif ns == "条件分岐" and nb.upper() == "AND":
                    # R2 — collect all consecutive 条件分岐+AND rows
                    group.append(nr)
                    j += 1
                    while j < n:
                        nr2 = raw_rows[j]
                        ns2 = nr2.get(shori_key, "")
                        nb2 = nr2.get(biko_key, "")
                        if ns2 == "条件分岐" and nb2.upper() == "AND":
                            group.append(nr2)
                            j += 1
                        else:
                            break
                    # 条件分岐+AND run ended → close this group immediately.
                    # The next row starts a fresh sentence.
                    closed_by_r2 = True
                    break  # exit outer accumulation loop

                else:
                    break  # neither 紐づけ nor 条件分岐+AND → stop

            # R1b — only when NOT closed by R2.
            # Absorb the very next non-紐づけ row to complete the sentence.
            if not closed_by_r2:
                if j < n and raw_rows[j].get(shori_key, "") != "紐づけ":
                    group.append(raw_rows[j])
                    j += 1

            sentences.append({"type": "grouped", "rows": group})
            i = j

        else:
            # R3 — standalone sentence
            sentences.append({"type": "single", "rows": [row]})
            i += 1

    return sentences



def _parse_normal_processing_section(ws, kw_row: int, stop_row: int) -> dict:
    """
    通常処理:
      - keyword row = kw_row, keyword は col E
      - Headers SAME ROW as keyword, starting from col E:
          col D(4) は事実上 開発用チェック (実際のデータによる)
          col E(5) = DB_Field (forced)
          col F(6) = Description (forced)
          col G(7) 以降は実際の header text を読む
      - 実際の観察: col D=開発用チェック, col E=DB_Field, col F=DB_Name,
                    col G=項目名/出力内容_1, col H=_2, col I=_3, col J=処理, ...
      - Data rows: kw_row+1 ~ stop_row-1
      - Include only if:
          開発用チェック == TRUE
          AND (項目名/出力内容_1 is not empty / not "0")
          AND 処理 is not empty
    """
    # Build col_headers from the keyword row itself.
    # Observed layout:
    #   col D(4)  = 開発用チェック (raw boolean values in data rows, no header text on kw_row)
    #   col E(5)  = DB_Field (forced)
    #   col F(6)  = DB_Name / Description (use actual text if present)
    #   col G(7)+ = actual headers read from kw_row
    col_headers: dict[int, str] = {}
    col_headers[COL_D] = "開発用チェック"   # forced — raw TRUE/FALSE in data
    col_headers[COL_E] = "DB_Field"         # forced
    h_f = _cv(ws, kw_row, COL_F)
    col_headers[COL_F] = h_f if h_f else "Description"
    for col in range(COL_G, COL_O + 1):
        h = _cv(ws, kw_row, col)
        if h:
            col_headers[col] = h
        else:
            break   # stop at first empty beyond col G

    dev_chk_col = COL_D   # always col D by design

    # Determine helper column indexes by header name
    shori_col = next((c for c, h in col_headers.items() if h == "処理"), None)
    item1_col = next((c for c, h in col_headers.items() if h == "項目名 / 出力内容_1"), None)

    data_start = kw_row + 1
    data_end   = stop_row - 1
    rows = []

    for row in range(data_start, data_end + 1):
        # Skip grey rows
        if _is_grey(ws.cell(row=row, column=COL_E)):
            continue

        # 開発用チェック is a raw boolean in col D (not a string)
        dev_cell = ws.cell(row=row, column=dev_chk_col).value
        if dev_cell is None:
            continue
        # Accept Python bool True or string "True"
        is_active = (dev_cell is True) or _bool_str(str(dev_cell))
        if not is_active:
            continue

        # Read all defined header cols (skip col D = dev_check)
        row_data: dict = {}
        has_any = False
        for col, hdr in col_headers.items():
            if col == dev_chk_col:
                continue
            v = _cv(ws, row, col)
            if v:
                row_data[hdr] = v
                has_any = True

        if not has_any:
            continue

        row_data["_dev_check"] = True

        # Include row only when both 項目名/出力内容_1 AND 処理 are present.
        # Rows with only dev_check=True but missing these belong to 特殊処理.
        item1 = row_data.get("項目名 / 出力内容_1", "").strip()
        shori = row_data.get("処理", "")

        if not shori and not item1:
            continue

        rows.append(row_data)

    return {
        "headers": {str(c): h for c, h in col_headers.items()},
        "dev_check_col": dev_chk_col,
        "rows": rows,
    }


def _parse_special_processing_section(ws, kw_row: int) -> dict:
    """
    特殊処理:
      - keyword row = kw_row, keyword は col E
      - Header は次の行 (kw_row + 1), col E(5) ~ N(14)
      - Data から kw_row+2 以降
      - 50行連続で全列空なら停止

    Grey-row rules:
      - Nếu MỘT dòng bị tô xám → bỏ qua dòng đó (không thêm vào rows/block)
      - Nếu TẤT CẢ dòng trong một sentence-block đều bị xám → bỏ qua cả block
    """
    header_row = kw_row + 1
    col_headers: dict[int, str] = {}
    for col in range(COL_E, COL_N + 1):
        h = _cv(ws, header_row, col)
        if h:
            col_headers[col] = h

    # Values that indicate a repeated section keyword/header row to skip
    section_keywords_set = REQUIRED_KEYWORDS
    header_value_set = set(col_headers.values())

    data_start = header_row + 1
    max_row    = ws.max_row

    # ── Pass 1: thu thập tất cả raw data rows kèm row-index và grey flag ──
    raw_collected: list[dict] = []   # {"row_data": dict, "excel_row": int, "grey": bool}
    empty_streak = 0
    current_block_meta: list[dict] = []   # tạm chứa rows của block hiện tại
    blocks_meta: list[list[dict]] = []    # danh sách các block

    for row in range(data_start, max_row + 1):
        row_data: dict = {}
        has_any = False
        for col in range(COL_E, COL_N + 1):
            v = _cv(ws, row, col)
            hdr = col_headers.get(col, get_column_letter(col))
            if v:
                row_data[hdr] = v
                has_any = True

        if not has_any:
            empty_streak += 1
            if empty_streak >= EMPTY_STREAK_LIMIT:
                break
            continue

        # Skip rows that repeat the section keyword (e.g. another 特殊処理 header block)
        e_val = _cv(ws, row, COL_E)
        if e_val in section_keywords_set:
            # We hit a new block! Save current block rows to blocks_meta
            if current_block_meta:
                blocks_meta.append(current_block_meta)
                current_block_meta = []

            # This is a repeated section header — reset and re-detect headers
            new_header_row = row + 1
            col_headers = {}
            for col in range(COL_E, COL_N + 1):
                h = _cv(ws, new_header_row, col)
                if h:
                    col_headers[col] = h
            header_value_set = set(col_headers.values())
            empty_streak = 0
            continue

        # Skip rows that are header repetitions (all values match header names)
        row_values = set(row_data.values())
        if row_values and row_values.issubset(header_value_set):
            empty_streak = 0
            continue

        # Detect if this Excel row is grey
        is_grey_row = _is_special_row_grey(ws, row)

        entry = {"row_data": row_data, "excel_row": row, "grey": is_grey_row}
        raw_collected.append(entry)
        current_block_meta.append(entry)
        empty_streak = 0

    if current_block_meta:
        blocks_meta.append(current_block_meta)

    # ── Pass 2: áp dụng grey rules ──
    #
    # Rule A: Nếu TẤT CẢ entries trong một block đều là grey → bỏ qua cả block
    # Rule B: Nếu CHỈ MỘT SỐ entries trong block bị grey → giữ block nhưng
    #         lọc bỏ các grey entries đó

    rows: list[dict] = []
    sentences: list[dict] = []

    for block in blocks_meta:
        all_grey = all(e["grey"] for e in block)
        if all_grey:
            # Rule A — bỏ qua cả block
            continue

        # Rule B — chỉ giữ lại các dòng không bị xám
        visible_rows = [e["row_data"] for e in block if not e["grey"]]

        if not visible_rows:
            continue

        rows.extend(visible_rows)
        sentences.append({
            "type": "grouped",
            "rows": visible_rows,
        })

    return {
        "headers": {str(c): h for c, h in col_headers.items()},
        "rows": rows,
        "sentences": sentences,
    }


# ─────────────────────────────────────────────────────────
# Master sheet parser
# ─────────────────────────────────────────────────────────

def parse_sheet(ws) -> dict | None:
    """
    Parse a single worksheet.
    Returns structured dict or None if the sheet lacks required keywords.
    """
    kw_map = _find_keywords(ws)

    if not REQUIRED_KEYWORDS.issubset(kw_map.keys()):
        return None

    row_fn,  col_fn  = kw_map["ファイル名"]
    row_out, _       = kw_map["出力条件"]
    row_nor, _       = kw_map["通常処理"]
    row_spe, _       = kw_map["特殊処理"]

    # ファイル名 is a sub-keyword inside the 出力条件 block (col B).
    # Its data runs from row_fn+1 until the 通常処理 section (exclusive).
    filename_stop = row_nor   # exclusive

    return {
        "ファイル名": _parse_filename_section(ws, row_fn, col_fn, filename_stop),
        "出力条件":   _parse_output_condition_section(ws, row_out, row_nor),
        "通常処理":   _parse_normal_processing_section(ws, row_nor, row_spe),
        "特殊処理":   _parse_special_processing_section(ws, row_spe),
    }


# ─────────────────────────────────────────────────────────
# Scan a single xlsx file
# ─────────────────────────────────────────────────────────

def scan_xlsx(file_path: str) -> dict:
    """
    Open an xlsx file and parse all sheets.
    Returns:
      {
        "valid_sheets":   [sheet_name, ...],
        "skipped_sheets": [sheet_name, ...],
        "data":           {sheet_name: parsed_data, ...},
      }
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    valid_sheets   = []
    skipped_sheets = []
    data           = {}

    for ws in wb.worksheets:
        name   = ws.title
        parsed = parse_sheet(ws)
        if parsed is not None:
            data[name] = parsed
            valid_sheets.append(name)
        else:
            kw_map = _find_keywords(ws)
            missing = [k for k in ["ファイル名", "出力条件", "通常処理", "特殊処理"] if k not in kw_map]
            if 0 < len(missing) <= 2:
                skipped_sheets.append(f"{name} --> thiếu KEYWORD: {', '.join(missing)}")
            else:
                skipped_sheets.append(name)

    wb.close()
    return {
        "valid_sheets":   valid_sheets,
        "skipped_sheets": skipped_sheets,
        "data":           data,
    }
