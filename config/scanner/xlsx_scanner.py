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


def _get_cell_fill_argb(cell) -> str | None:
    """Return ARGB hex string of cell fill if colored, None if default/white/black."""
    try:
        fill = cell.fill
        if not fill or fill.fill_type in (None, "none"):
            return None
        fg = fill.fgColor
        if not fg:
            return None
        if fg.type == "rgb":
            argb = fg.rgb.upper() if fg.rgb else None
            # Exclude transparent, white, and black header
            if argb in ("00000000", "FFFFFFFF", "00FFFFFF", "FF000000"):
                return None
            return argb
        if fg.type == "theme":
            return f"theme_{fg.theme}_{fg.tint}"
    except Exception:
        pass
    return None



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


def _parse_output_condition_section(ws, kw_row: int, kw_col: int, stop_row: int) -> dict:
    """
    出力条件:
      - keyword row に keyword がある (col E)
      - Header は次の行 (kw_row + 1) の col 5 以降 (実際は col E~M)
      - Data は kw_row+2 から stop_row-1 まで
      - Grouping rules が適用される
    """
    header_row = kw_row + 1
    col_headers: dict[int, str] = {}
    # Scan from kw_col to col N(14) for headers (chỉ lấy từ cột chứa 出力条件 trở sang phải)
    for col in range(kw_col, COL_N + 1):
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
            # R3 / R4 — Check if it's a standalone sentence or just a context header
            shori_val = row.get(shori_key, "")
            joken_val = row.get("条件", "")
            mokuteki_val = row.get("目的", "")
            gaitou_val = row.get("該当項目名_1", "")

            if not (shori_val or joken_val or mokuteki_val or gaitou_val):
                # R4 — Dòng tiêu đề phụ mang tính ngữ cảnh (context)
                sentences.append({"type": "header", "rows": [row]})
            else:
                # R3 — Standalone sentence
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


def _is_special_output_row(r: dict) -> bool:
    """Return True if row in 特殊処理 is an output/assignment row targeting a field."""
    fld = r.get("項目名", "")
    shori = r.get("処理", "")
    if fld and shori not in ("紐づけ", "条件分岐"):
        return True
    if shori in ("出力", "固定値出力", "そのまま出力", "指定文字削除出力", "削除"):
        return True
    return False


def _is_special_prep_row(r: dict) -> bool:
    """Return True if row is a data preparation step (e.g. creating buffer column)."""
    shori = r.get("処理", "")
    cond = r.get("条件", "")
    if shori in ("結合", "ハイフン付出力") and ("新たな項目として追加" in cond or "新たな項目として" in cond):
        return True
    return False


def _extract_case_name(rows: list[dict], case_idx: int) -> str:
    """Extract a human-readable title for a special case block."""
    SUB_PURPOSES = {"契約形態の確認", "メールアドレスの情報を出力するため"}
    for r in rows:
        m = r.get("目的")
        if m and m.strip() not in SUB_PURPOSES:
            return m.strip().split("\n")[0]
    for r in rows:
        m = r.get("目的")
        if m:
            return m.strip().split("\n")[0]
    for r in rows:
        f = r.get("項目名")
        if f:
            return f.strip()
    return f"Case_{case_idx}"


def _group_case_sentences(case_items: list) -> list[dict]:
    """
    Lớp 2: Phân loại các rows bên trong 1 case thành các sentences.
    - Tiền xử lý (Pre-processing/tạo cột đệm): Tách thành sentence riêng (Slice 2).
    - Kế thừa ngữ cảnh (Context Inheritance): Nếu có các dòng 紐づけ chung cho nhiều cột output độc lập,
      nhân bản các dòng 紐づけ vào từng sentence con tương ứng (Slice 3).
    - Cây điều kiện (SWITCH/Branching): Gom các dòng điều kiện và output cùng trường vào 1 sentence.
    - Ranh giới thị giác (Visual & Blank Delimiters): Tách sentence con khi đổi màu nền hoặc dòng trống + mục đích mới.
    """
    if not case_items:
        return []

    def _r(item):
        return item["row_data"] if isinstance(item, dict) and "row_data" in item else item

    if len(case_items) == 1:
        return [{"type": "single", "rows": [_r(case_items[0])]}]

    sentences = []

    # 1. Tách các dòng tiền xử lý tạo cột đệm ở đầu case (Pre-processing separation)
    idx = 0
    while idx < len(case_items) and _is_special_prep_row(_r(case_items[idx])):
        sentences.append({"type": "single", "rows": [_r(case_items[idx])]})
        idx += 1

    remaining = case_items[idx:]
    if not remaining:
        return sentences

    # 2. Kiểm tra trường hợp Kế thừa ngữ cảnh (Context Inheritance):
    # Các dòng 紐づけ đi kèm nhiều dòng output độc lập cho các trường đích khác nhau (không có 条件分岐 xen kẽ)
    join_rows = []
    i = 0
    while i < len(remaining) and _r(remaining[i]).get("処理") == "紐づけ":
        join_rows.append(_r(remaining[i]))
        i += 1

    output_rows = [_r(x) for x in remaining[i:]]
    distinct_target_fields = set()
    has_cond = any(r.get("処理") == "条件分岐" for r in output_rows)
    all_simple_outputs = True

    for r in output_rows:
        sh = r.get("処理", "")
        f = r.get("項目名", "")
        if sh in ("そのまま出力", "出力", "固定値出力", "指定文字削除出力"):
            if f:
                distinct_target_fields.add(f)
        else:
            all_simple_outputs = False

    if join_rows and len(output_rows) > 1 and len(distinct_target_fields) > 1 and not has_cond and all_simple_outputs:
        # Context Inheritance: Nhân bản join_rows cho từng output riêng biệt
        for out_r in output_rows:
            sentences.append({"type": "grouped", "rows": join_rows + [out_r]})
    else:
        # Chuẩn gom nhóm cho cây điều kiện / single output
        current_sentence_rows = []
        seen_output = False
        current_fill = None

        for item in remaining:
            r = _r(item)
            f_col = item.get("fill_color") if isinstance(item, dict) and "row_data" in item else None
            had_blank = item.get("had_blank_before", False) if isinstance(item, dict) and "row_data" in item else False
            shori = r.get("処理", "")
            is_out = _is_special_output_row(r)
            has_mokuteki = bool(r.get("目的"))

            # Check boundaries:
            # 1. Join boundary
            is_join_boundary = (seen_output and shori == "紐づけ")
            # 2. Visual boundary: Color shift OR (Blank gap + New purpose)
            is_color_shift = (seen_output and f_col and current_fill and f_col != current_fill)
            is_blank_section = (seen_output and had_blank and has_mokuteki)

            if is_join_boundary or is_color_shift or is_blank_section:
                if current_sentence_rows:
                    s_type = "single" if len(current_sentence_rows) == 1 else "grouped"
                    sentences.append({"type": s_type, "rows": current_sentence_rows})
                    current_sentence_rows = []
                    seen_output = False
                    current_fill = f_col

            if current_fill is None and f_col is not None:
                current_fill = f_col

            current_sentence_rows.append(r)
            if is_out:
                seen_output = True

        if current_sentence_rows:
            s_type = "single" if len(current_sentence_rows) == 1 else "grouped"
            sentences.append({"type": s_type, "rows": current_sentence_rows})

    return sentences


def _parse_special_processing_section(ws, kw_row: int) -> dict:
    """
    特殊処理:
      - keyword row = kw_row, keyword は col E
      - Header は次の行 (kw_row + 1), col E(5) ~ N(14)
      - Data から kw_row+2 以降
      - 50行連続で全列空なら停止

    Cải tiến 2 lớp:
      - Lớp 1 (Cases): Gom các ý theo từng dòng header phân tách thành các trường hợp đặc biệt riêng biệt (cases).
      - Lớp 2 (Sentences): Trong mỗi case, phân loại thành các sentences dựa theo dòng output kèm điều kiện liên kết,
        màu nền và dòng trống phân cách.

    Grey-row rules:
      - Nếu MỘT dòng bị tô xám → bỏ qua dòng đó (không thêm vào rows/block)
      - Nếu TẤT CẢ dòng trong một case đều bị xám → bỏ qua cả case
    """
    header_row = kw_row + 1
    col_headers: dict[int, str] = {}
    for col in range(COL_E, COL_N + 1):
        h = _cv(ws, header_row, col)
        if h:
            col_headers[col] = h

    section_keywords_set = REQUIRED_KEYWORDS
    header_value_set = set(col_headers.values())

    data_start = header_row + 1
    max_row    = ws.max_row

    # Kiểm tra xem sheet có nhiều dòng header lặp lại (như Contact_Table) hay không
    pure_headers_count = 0
    for r in range(data_start, max_row + 1):
        v_e_chk = _cv(ws, r, COL_E)
        v_f_chk = _cv(ws, r, COL_E + 1)
        if v_e_chk == "項目名" and v_f_chk == "目的":
            pure_headers_count += 1

    has_multiple_headers = (pure_headers_count > 0)

    # ── Pass 1: Thu thập raw data rows theo từng Case (phân tách bởi dòng header) ──
    raw_collected: list[dict] = []
    empty_streak = 0
    current_case_entries: list[dict] = []
    cases_raw: list[list[dict]] = []

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

        had_blank_before = (empty_streak > 0)
        empty_streak = 0

        # Nhận diện dòng header phân tách
        v_e = _cv(ws, row, COL_E)
        v_f = _cv(ws, row, COL_E + 1)
        row_values = set(row_data.values())

        # Kiểm tra xem dòng có chứa action keyword nghiệp vụ hay không (từ cột H trở đi)
        ACTION_KEYWORDS = {
            "そのまま出力", "出力", "固定値出力", "指定文字削除出力", "削除",
            "条件分岐", "結合", "結合（条件付き）", "ハイフン付出力", "紐づけ", "重複チェック"
        }
        has_action = any(
            _cv(ws, row, c) in ACTION_KEYWORDS
            for c in range(COL_E + 3, COL_N + 1)
        )

        is_pure_header = (
            (v_e in section_keywords_set
             or (len(row_values) >= 2 and row_values.issubset(header_value_set))
             or (v_e == "項目名" and v_f == "目的"))
            and not has_action
        )

        is_hybrid_row = (v_e == "項目名" and v_f == "目的") and has_action

        # Adaptive boundary cho các sheet không lặp lại header (ví dụ Account.xlsx):
        # Tách case khi:
        # 1. Có khoảng trống trước đó (had_blank_before) và dòng mới có 目的 (v_f) nhưng không có 項目名 (v_e)
        # 2. Hoặc bước tiền xử lý chuyển đổi mục đích (ví dụ Row 39: 入居者情報作成に必要な項目の追加)
        is_blank_case_boundary = False
        if not has_multiple_headers and current_case_entries and not is_pure_header:
            if had_blank_before and v_f and not v_e:
                is_blank_case_boundary = True
            elif v_f and not v_e and "追加" in str(v_f):
                is_blank_case_boundary = True

        if is_pure_header or is_blank_case_boundary:
            # Gặp dòng header thuần túy hoặc ranh giới case -> Gom các ý phía trên thành 1 case
            if current_case_entries:
                cases_raw.append(current_case_entries)
                current_case_entries = []

            if is_pure_header:
                # Nếu dòng header lặp lại keyword section, đọc lại header ở dòng kế tiếp
                if v_e in section_keywords_set:
                    new_header_row = row + 1
                    col_headers = {}
                    for col in range(COL_E, COL_N + 1):
                        h = _cv(ws, new_header_row, col)
                        if h:
                            col_headers[col] = h
                    header_value_set = set(col_headers.values())

                empty_streak = 0
                continue

        c_e = ws.cell(row=row, column=COL_E)
        c_f = ws.cell(row=row, column=COL_E + 1)
        fill_color = _get_cell_fill_argb(c_e) or _get_cell_fill_argb(c_f)

        if is_hybrid_row:
            # Dòng chứa dữ liệu nhưng bị dính nhãn header ở cột E, F (ví dụ Row 119)
            for dummy_k in ("項目名", "目的", "該当ファイル名", "開発用チェック"):
                if row_data.get(dummy_k) == dummy_k:
                    del row_data[dummy_k]
            # Kế thừa trường đích của case nếu có
            last_case_field = ""
            for e in current_case_entries:
                f_val = e["row_data"].get("項目名")
                if f_val and f_val != "項目名":
                    last_case_field = f_val
            if last_case_field and not row_data.get("項目名"):
                row_data["項目名"] = last_case_field

            is_grey_row = _is_special_row_grey(ws, row)
            entry = {
                "row_data": row_data,
                "excel_row": row,
                "grey": is_grey_row,
                "fill_color": fill_color,
                "had_blank_before": had_blank_before,
            }
            raw_collected.append(entry)
            current_case_entries.append(entry)
            # Vì dòng này cũng đóng vai trò phân tách case phía trước với case phía sau
            cases_raw.append(current_case_entries)
            current_case_entries = []
            empty_streak = 0
            continue

        # Detect if this Excel row is grey
        is_grey_row = _is_special_row_grey(ws, row)

        entry = {
            "row_data": row_data,
            "excel_row": row,
            "grey": is_grey_row,
            "fill_color": fill_color,
            "had_blank_before": had_blank_before,
        }
        raw_collected.append(entry)
        current_case_entries.append(entry)
        empty_streak = 0

    if current_case_entries:
        cases_raw.append(current_case_entries)

    # ── Pass 2: Áp dụng grey rules & phân loại 2 lớp ──
    cases: list[dict] = []
    rows: list[dict] = []
    sentences: list[dict] = []

    for idx, case_entries in enumerate(cases_raw, 1):
        all_grey = all(e["grey"] for e in case_entries)
        if all_grey:
            # Rule A — bỏ qua cả case nếu tất cả đều xám
            continue

        # Rule B — chỉ giữ lại các dòng không bị xám
        visible_entries = [e for e in case_entries if not e["grey"]]
        if not visible_entries:
            continue

        visible_rows = [e["row_data"] for e in visible_entries]

        case_sentences = _group_case_sentences(visible_entries)
        case_name = _extract_case_name(visible_rows, idx)

        cases.append({
            "case_index": idx,
            "case_name": case_name,
            "rows": visible_rows,
            "sentences": case_sentences,
        })

        sentences.extend(case_sentences)

    return {
        "headers": {str(c): h for c, h in col_headers.items()},
        "cases": cases,
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
    row_out, col_out = kw_map["出力条件"]
    row_nor, _       = kw_map["通常処理"]
    row_spe, _       = kw_map["特殊処理"]

    # ファイル名 is a sub-keyword inside the 出力条件 block (col B).
    # Its data runs from row_fn+1 until the 通常処理 section (exclusive).
    filename_stop = row_nor   # exclusive

    return {
        "ファイル名": _parse_filename_section(ws, row_fn, col_fn, filename_stop),
        "出力条件":   _parse_output_condition_section(ws, row_out, col_out, row_nor),
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
