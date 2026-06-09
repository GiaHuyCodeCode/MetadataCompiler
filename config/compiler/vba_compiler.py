"""
vba_compiler.py — MetadataCompiler V1
Sinh mã VBA (.bas) từ sheet_raw.json theo sk-architect + sk-standard SKILL.
"""

import os
import json
import re
from typing import Optional


# ─────────────────────────────────────────────────────────
# Condition value parser
# ─────────────────────────────────────────────────────────

def _parse_condition_value(cond_text: str) -> str:
    """Extract actual comparison value from condition text.

    Japanese condition column (条件 / 項目マッピング) uses patterns like:
        '"個人"の場合'  → '個人'
        '"1"の場合'     → '1'
        '解約予定''     → '解約予定'
        '入居有り''     → '入居有り'
        '"-"の場合'     → '-'
        '"退去済"'      → '退去済'
    """
    text = cond_text.strip()
    # Strip trailing の場合
    text = re.sub(r'の場合$', '', text)
    # Strip trailing single quote (Excel artifact)
    text = text.rstrip("'")
    # Clean up prefixes
    text = re.sub(r'^(すべて|いずれも|上記以外で|上記以外かつ|上記以外)', '', text).strip()
    # Strip surrounding escaped double quotes: \"X\" → X
    text = text.replace('\\"', '"')
    # Strip surrounding double quotes
    text = text.strip('"').strip()
    return text


def _is_hyphen_null_mapping(row: dict) -> bool:
    """Check if 項目マッピング indicates hyphen-to-null conversion."""
    mapping = row.get("項目マッピング", "") or row.get("条件 / 項目マッピング", "")
    return ("-" in mapping and "ブランク" in mapping) or mapping == '"-"の場合はブランク出力'


def _detect_condition_type(cond_text: str) -> str:
    """Classify condition text into pattern type.

    Returns one of:
        EQUALS       - 'X'の場合 → field = 'X'
        LIKE         - を含む / 含まれている → LIKE '%X%'
        NOT_LIKE     - を含まない / 含まれていない → NOT LIKE '%X%'
        AND_ALL      - すべて"X"の場合 → multi-field AND
        ELSE         - 上記...該当しない / いずれも → True branch
        OR_LIST      - AまたはB → IN ('A','B')
        HYPHEN_NULL  - "-"の場合 + 出力なし → IIF null
        BLANK_CHECK  - ブランクの場合 → Nz check
        LEFT_JOIN    - 紐づかない場合
        POST_UPDATE  - 上記で出力したデータの...
        UNKNOWN
    """
    if not cond_text:
        return "UNKNOWN"
    if "含まれていない" in cond_text or "を含まない" in cond_text:
        return "NOT_LIKE"
    if "含まれている" in cond_text or "を含む" in cond_text or "含まれる" in cond_text:
        return "LIKE"
    if "紐づかない" in cond_text:
        return "LEFT_JOIN"
    if "すべて" in cond_text:
        return "AND_ALL"
    if "該当しない" in cond_text or "いずれも" in cond_text:
        return "ELSE"
    if "または" in cond_text:
        return "OR_LIST"
    if "ブランクの場合" in cond_text:
        return "BLANK_CHECK"
    if "上記で出力した" in cond_text:
        return "POST_UPDATE"
    if "存在しない" in cond_text or "NOT EXISTS" in cond_text:
        return "NOT_EXISTS"
    return "EQUALS"


def _extract_like_keyword(cond_text: str) -> str:
    """Extract keyword from LIKE/NOT LIKE condition text.

    Examples:
        '"会社"が含まれていない場合' → '会社'
        '連絡不要'が含まれていない場合' → '連絡不要'
    """
    # Try \"X\" pattern first
    m = re.search(r'["\\"]+([^"\\]+)["\\"]+', cond_text)
    if m:
        return m.group(1)
    # Try 'X' pattern (with leading single quote artifact)
    m = re.search(r"'?([^'が含]+)'?が含", cond_text)
    if m:
        return m.group(1).rstrip("'")
    # Fallback: remove known suffixes
    text = re.sub(r'(が含まれて(いない|いる)場合|を含(む|まない).*|の場合)$', '', cond_text)
    return text.strip('"\\\' ')


# ─────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────

# Map tên sheet → (target_table, sub_name, id_prefix)
SHEET_TARGET_MAP = {
    "Account":    ("Account",    "set_account",    None),
    "Building":   ("Building",   "set_building",   None),
    "Property":   ("Property",   "set_property",   None),
    "Contract":   ("Contract",   "set_contract",   None),
    "Contractor": ("Contractor", "set_contractor", None),
    "Owner":      ("Owner",      "set_owner",      None),
}

# Keyword → target table (từ tên sheet tiếng Nhật)
KEYWORD_TABLE_MAP = [
    (["アカウント", "account"], "Account",    "set_account"),
    (["建物",       "building"],"Building",   "set_building"),
    (["部屋",       "room"],    "Property",   "set_property"),
    (["契約",       "contract"],"Contract",   "set_contract"),
    (["オーナー",   "owner"],   "Owner",      "set_owner"),
]

def _infer_target(sheet_name: str) -> tuple[str, str]:
    """(target_table, sub_name) từ tên sheet."""
    sl = sheet_name.lower()
    for kws, tbl, sub in KEYWORD_TABLE_MAP:
        if any(k in sl or k in sheet_name for k in kws):
            return tbl, sub
    return "TargetTable", "set_data"


def _infer_prefix(sheet_name: str) -> str:
    """ID prefix từ tên sheet."""
    # Lấy suffix như _N, _P, _E, _S nếu có
    m = re.search(r'_([A-Z]+)$', sheet_name)
    suffix = m.group(1) if m else ""

    sl = sheet_name.lower()
    if "オーナー" in sheet_name or "owner" in sl:
        return "ON"
    if "建物" in sheet_name or "building" in sl:
        return "B"
    if "部屋" in sheet_name or "room" in sl:
        return "BN"
    if suffix:
        return suffix
    # default: viết tắt từ chữ cái đầu
    return "X"


def _clean_table_name(raw: str) -> str:
    """Chuẩn hóa tên bảng: bỏ 【GMO用】, bỏ .csv, bỏ GMO prefix."""
    name = raw.replace("【GMO用】", "").replace(".csv", "").strip()
    name = re.sub(r'^GMO\s+', '', name)
    return name


def _primary_source_table(filenames: list) -> str:
    """Lấy bảng nguồn chính (priority=1)."""
    if not filenames:
        return "SourceTable"
    first = filenames[0].get("file_name", "")
    return _clean_table_name(first)


def _secondary_source_tables(filenames: list) -> list[str]:
    return [_clean_table_name(f.get("file_name", "")) for f in filenames[1:]]


def _sql(lines: list[str]) -> str:
    """Build SQL block string từ list các dòng."""
    out = '    SQL = ""\n'
    for line in lines:
        escaped = line.replace('"', '""')
        escaped = re.sub(r'""\s*&\s*con(\d+)\s*&\s*""', r'" & con\1 & "', escaped)
        out += f'    SQL = SQL & "{escaped}"\n'
    return out


def _db_execute_log(log_msg: str) -> str:
    return f'    db.Execute SQL\n    log_write "{log_msg}"\n'


# ─────────────────────────────────────────────────────────
# Section generators
# ─────────────────────────────────────────────────────────

def _gen_header(sheet_name: str, target_table: str, sub_name: str) -> str:
    return (
        f'Attribute VB_Name = "{sheet_name}"\n'
        f'Option Compare Database\n'
        f'Option Explicit\n\n'
        f'Sub {sub_name}()\n'
        f'    log_write "{sub_name}:in"\n\n'
        f'    Dim t As Single\n'
        f'    t = Timer\n\n'
        f'    Dim db As ADODB.Connection\n'
        f'    Dim SQL As String\n\n'
        f'    Set db = CurrentProject.Connection\n\n'
    )


def _gen_flg_init(source_table: str, flg_reset_val: str = "0") -> str:
    return (
        f"    '出力条件\n"
        f'    AddNewFieldToTable "{source_table}", "FLG", "TEXT(1)"\n\n'
        f"    '■FLGリセット\n"
        f'    db.Execute "UPDATE {source_table} SET {source_table}.[FLG] = \'{flg_reset_val}\';"\n\n'
    )


def _gen_flg_cleanup(source_table: str) -> str:
    return (
        f"\n    '■項目削除\n"
        f'    DeleteFieldInTable "{source_table}", "FLG"\n'
    )


def _gen_footer(target_table: str, sub_name: str) -> str:
    return (
        f"\n    chk_required\n\n"
        f'    Debug.Print "{target_table}:" & Timer - t\n'
        f'    log_write "{sub_name}:out"\n\n'
        f'End Sub\n'
    )


def _gen_delete_existing(target_table: str, prefix: str) -> str:
    return (
        f'    db.Execute "DELETE FROM {target_table} WHERE ID LIKE \'{prefix}_%\';"\n\n'
        f'    If DCount("ID", "{target_table}", "ID LIKE \'{prefix}_*\'") = 0 Then\n\n'
    )


def _gen_section_sep(sheet_name: str) -> str:
    return (
        f"    '{'='*80}\n"
        f"    '{sheet_name}\n"
        f"    '{'='*80}\n\n"
    )


# ─────────────────────────────────────────────────────────
# Sentence Classifier
# ─────────────────────────────────────────────────────────

def _classify_sentence(rows: list) -> tuple[str, float]:
    """
    Returns (pattern_name, confidence).
    pattern_name: DEDUP | JOIN_FILTER_LEFT | JOIN_FILTER | IN_LIST_FILTER | BLANK_CHECK | NO_FILTER | UNKNOWN
    """
    if len(rows) == 1 and rows[0].get("処理") == "出力" and not (rows[0].get("条件", "") or rows[0].get("条件 / 項目マッピング", "")):
        return "NO_FILTER", 1.0

    shori_vals = [r.get("処理", "") for r in rows]
    cond_vals  = [r.get("条件", "") + r.get("条件 / 項目マッピング", "") for r in rows]
    all_cond   = " ".join(cond_vals)

    has_dedup        = any("重複チェック" in s for s in shori_vals)
    has_join         = any("紐づけ" in s for s in shori_vals)
    has_not_linked   = any("紐づかない" in c for c in cond_vals)
    has_or           = "または" in all_cond
    has_blank_delete = any(s == "削除" for s in shori_vals)
    has_multi_flg    = any("FLG='2'" in c or "FLG=2" in c for c in cond_vals)

    if has_dedup:
        return "DEDUP", 1.0
    if has_join and has_not_linked:
        return "JOIN_FILTER_LEFT", 0.95
    if has_join:
        return "JOIN_FILTER", 0.9
    if has_or:
        return "IN_LIST_FILTER", 0.85
    if has_blank_delete:
        return "BLANK_CHECK", 0.8
    if has_multi_flg:
        return "MULTI_FLG_FLOW", 0.75
    return "UNKNOWN", 0.3

def _gen_output_conditions(sentences: list, source_table: str, sheet_name: str, sub_name: str, target_table: str) -> tuple[str, str]:
    """
    Returns (flg_reset_val, vba_code).
    flg_reset_val: '0' (normal) hoặc '1' (IN_LIST_FILTER ngược)
    """
    code = ""
    flg_reset = "0"
    prefix = _infer_prefix(sheet_name)
    sheet_context = {
        "sheet_name": sheet_name,
        "source_table": source_table,
        "target_table": target_table,
        "prefix": prefix,
        "sub_name": sub_name,
    }

    for sentence in sentences:
        rows = sentence.get("rows", [])
        if not rows:
            continue

        pattern, confidence = _classify_sentence(rows)

        if pattern == "NO_FILTER":
            code += "    ' 出力 (No filter required)\n"

        elif pattern == "DEDUP":
            dedup_row = next(r for r in rows if "重複チェック" in r.get("処理", ""))
            key_field = dedup_row.get("該当項目名_1", "ID")
            src = _clean_table_name(dedup_row.get("該当ファイル名", source_table))
            code += (
                f"    '重複チェック {key_field}をキーに重複を削除\n"
                + _sql([
                    f"UPDATE {src} ",
                    f"SET {src}.FLG = '1' ",
                    f"WHERE {src}.ID NOT IN ( ",
                    f"    SELECT MIN(ID) ",
                    f"    FROM {src} ",
                    f"    WHERE FLG = '0' ",
                    f"    GROUP BY [{key_field}] ",
                    f"); ",
                ])
                + "    db.Execute SQL\n\n"
            )

        elif pattern == "JOIN_FILTER_LEFT":
            join_rows = [r for r in rows if r.get("処理") == "紐づけ"]
            if len(join_rows) >= 2:
                t1_src = _clean_table_name(join_rows[0].get("該当ファイル名", source_table))
                t2_src = _clean_table_name(join_rows[1].get("該当ファイル名", ""))
                keys = [v for k in ["該当項目名_1","該当項目名_2","該当項目名_3"]
                        if (v := join_rows[0].get(k, ""))]
                on_clause = " AND ".join(f"(T1.[{k}] = T2.[{k}])" for k in keys)
                code += (
                    "    'JOIN filter → 紐づかないデータを抽出\n"
                    + _sql([
                        f"UPDATE {t1_src} AS T1 ",
                        f"LEFT JOIN {t2_src} AS T2 ",
                        f"ON {on_clause} ",
                        f"SET T1.FLG = '0' ",
                        f"WHERE T2.[{keys[0]}] IS NULL; ",
                    ])
                    + "    db.Execute SQL\n\n"
                )
            flg_reset = "1"

        elif pattern == "JOIN_FILTER":
            join_rows = [r for r in rows if r.get("処理") == "紐づけ"]
            if len(join_rows) >= 2:
                t1_src = _clean_table_name(join_rows[0].get("該当ファイル名", source_table))
                t2_src = _clean_table_name(join_rows[1].get("該当ファイル名", ""))
                keys = [v for k in ["該当項目名_1","該当項目名_2","該当項目名_3"]
                        if (v := join_rows[0].get(k, ""))]
                on_clause = " AND ".join(f"(T1.[{k}] = T2.[{k}])" for k in keys)
                code += (
                    "    'JOIN filter\n"
                    + _sql([
                        f"UPDATE {t1_src} AS T1 ",
                        f"INNER JOIN {t2_src} AS T2 ",
                        f"ON {on_clause} ",
                        f"SET T1.FLG = '1'; ",
                    ])
                    + "    db.Execute SQL\n\n"
                )

        elif pattern == "IN_LIST_FILTER":
            flg_reset = "1"
            # Lấy field và values từ condition text
            cond_row = next((r for r in rows if "または" in r.get("条件","") + r.get("条件 / 項目マッピング","")), rows[0])
            src_field = cond_row.get("該当項目名_1", "")
            src_tbl   = _clean_table_name(cond_row.get("該当ファイル名", source_table))
            cond_text = cond_row.get("条件","") or cond_row.get("条件 / 項目マッピング","")
            # Parse "AまたはB" → ['A','B']
            parts = re.split(r'または', cond_text)
            values = [p.replace("の場合","").strip().strip('"').strip("'").strip() for p in parts]
            if src_field and values:
                where_parts = " OR ".join(f"T.[{src_field}] = '{v}'" for v in values)
                code += (
                    f"    '■有効レコードのみFLG=0に戻す\n"
                    + _sql([
                        f"UPDATE {src_tbl} AS T ",
                        f"SET T.FLG = '0' ",
                        f"WHERE {where_parts}; ",
                    ])
                    + "    db.Execute SQL\n\n"
                )

        elif pattern == "BLANK_CHECK":
            del_row = next((r for r in rows if r.get("処理") == "削除"), rows[0])
            src_field = del_row.get("該当項目名_1", "")
            src_tbl   = _clean_table_name(del_row.get("該当ファイル名", source_table))
            if src_field:
                code += (
                    f"    'Blank check — {src_field}\n"
                    + _sql([
                        f"UPDATE {src_tbl} AS T ",
                        f"SET T.FLG = '1' ",
                        f"WHERE Nz(T.[{src_field}], '') = ''; ",
                    ])
                    + "    db.Execute SQL\n\n"
                )

        else:
            # Gọi AI fallback để sinh code thực tế cho output condition không nhận diện được
            from compiler.ai_fallback import ai_fallback_generate
            code += ai_fallback_generate(sentence, sheet_context)

    log_msg = f"{sub_name}:{sheet_name} → 不要行を削除するため(FLG=1更新)"
    code += f'    log_write "{log_msg}"\n\n'
    return flg_reset, code


def Nz_check(r):
    return False  # placeholder


# ─────────────────────────────────────────────────────────
# 通常処理 generator
# ─────────────────────────────────────────────────────────

def _gen_normal_processing(rows: list, source_table: str, target_table: str,
                            prefix: str, sheet_name: str, sub_name: str) -> str:
    if not rows:
        return ""

    fields = []
    fields.append(f"    '{prefix}_' & T.ID as ID, ".replace("{prefix}", prefix))

    for row in rows:
        field = row.get("DB_Field", "")
        shori = row.get("処理", "")
        val1  = row.get("項目名 / 出力内容_1", "").rstrip("'")
        val2  = row.get("項目名 / 出力内容_2", "")
        val3  = row.get("項目名 / 出力内容_3", "")
        moji  = row.get("指定削除文字列", "")
        mapping = row.get("項目マッピング", "")

        if not field or not shori:
            continue

        # ── Check 項目マッピング for HYPHEN_TO_NULL override ──
        # When 項目マッピング contains "-"の場合はブランク出力,
        # override そのまま出力 to use IIF(T.[field]='-', NULL, T.[field])
        has_hyphen_null = _is_hyphen_null_mapping(row)

        if shori == "固定値出力":
            v = val1.strip("'\" ")
            fields.append(f"    '{v}' as {field}, ")
        elif shori == "そのまま出力" and has_hyphen_null:
            # 項目マッピング says: if value is "-", output NULL
            if val1:
                fields.append(f"    IIF(T.[{val1}] = '-', NULL, T.[{val1}]) as {field}, ")
            else:
                fields.append(f"    NULL as {field}, ")
        elif shori == "そのまま出力":
            if val1:
                fields.append(f"    T.[{val1}] as {field}, ")
            else:
                fields.append(f"    NULL as {field}, ")
        elif shori == "yyyy-mm-dd形式":
            fields.append(f"    IIF(IsDate(T.[{val1}]), Format(T.[{val1}], 'yyyy-mm-dd'), NULL) as {field}, ")
        elif shori == "yyyymm形式":
            fields.append(f"    Format(T.[{val1}], 'yyyymm') as {field}, ")
        elif shori == "ハイフン付出力":
            parts = [f"T.[{v}]" for v in [val1, val2, val3] if v]
            fields.append(f"    {' & \"-\" & '.join(parts)} as {field}, ")
        elif shori == "マージ":
            parts = [f"T.[{v}]" for v in [val1, val2, val3] if v]
            fields.append(f"    {' & '.join(parts)} as {field}, ")
        elif shori in ("ハイフンをNull変換", "指定文字をNull変換"):
            fields.append(f"    IIF(T.[{val1}] = '-', NULL, T.[{val1}]) as {field}, ")
        elif shori == "指定文字削除出力":
            chars = [c.strip() for c in moji.split(",") if c.strip()]
            expr = f"T.[{val1}]"
            for ch in chars:
                expr = f"Replace({expr}, '{ch}', '')"
            fields.append(f"    Val({expr}) as {field}, ")

    # sheet field — BẮT BUỘC
    fields.append(f"    '{sheet_name}' as sheet ")

    # Build INSERT
    field_lines = [f"INSERT INTO {target_table} SELECT "]
    for f in fields:
        field_lines.append(f"    {f.strip()}")
    field_lines.append(f"FROM {source_table} AS T ")
    field_lines.append(f"WHERE T.FLG = '0'; ")

    code = "    '通常処理\n" + _sql(field_lines)
    code += f'    db.Execute SQL\n'
    code += f'    log_write "{sub_name}:{sheet_name} → 通常処理"\n\n'
    return code


def parse_rows_into_blocks(rows: list[dict]) -> list[dict]:
    # Propagate field names (項目名 / DB_Field) and source fields (該当項目名_1)
    # for output actions that don't specify them (due to Excel cell merges).
    last_field = ""
    last_src_field = ""
    for r in rows:
        f = r.get("項目名") or r.get("DB_Field") or ""
        if f:
            last_field = f
        else:
            if last_field and r.get("処理") in ("出力", "固定値出力", "そのまま出力", "削除"):
                r["項目名"] = last_field

        sf = r.get("該当項目名_1") or ""
        if sf:
            last_src_field = sf
        else:
            if last_src_field and r.get("処理") in ("出力", "固定値出力", "そのまま出力", "削除"):
                r["該当項目名_1"] = last_src_field

    blocks = []
    current_block_branches = []
    current_branch_conds = []
    current_branch_outs = []
    seen_outputs = False
    
    for r in rows:
        cond_text = r.get("条件 / 項目マッピング", "") or r.get("条件", "") or ""
        
        # Check for end of block
        if "上記の分岐完了後" in cond_text or "上記の分岐完了後" in r.get("条件", ""):
            if current_branch_conds or current_branch_outs:
                current_block_branches.append({
                    "conditions": current_branch_conds,
                    "outputs": current_branch_outs
                })
            if current_block_branches:
                blocks.append({"branches": current_block_branches})
            current_block_branches = []
            current_branch_conds = []
            current_branch_outs = []
            seen_outputs = False
            continue
            
        # A row is an output row if its action (処理) is an output/delete action
        is_out = r.get("処理") in ("出力", "固定値出力", "そのまま出力", "削除")
        
        # A row is a condition if it's explicitly a condition or not an output row
        is_cond = r.get("処理") == "条件分岐" or not is_out
        
        if is_cond and is_out:
            # Both condition and output (e.g. kind_id mapping or tag mapping)
            if current_branch_conds or current_branch_outs:
                current_block_branches.append({
                    "conditions": current_branch_conds,
                    "outputs": current_branch_outs
                })
            current_block_branches.append({
                "conditions": [r],
                "outputs": [r]
            })
            current_branch_conds = []
            current_branch_outs = []
            seen_outputs = False
        elif is_cond:
            if seen_outputs:
                # Start new branch in current block
                current_block_branches.append({
                    "conditions": current_branch_conds,
                    "outputs": current_branch_outs
                })
                current_branch_conds = [r]
                current_branch_outs = []
                seen_outputs = False
            else:
                current_branch_conds.append(r)
        elif is_out:
            current_branch_outs.append(r)
            seen_outputs = True
            
    if current_branch_conds or current_branch_outs:
        current_block_branches.append({
            "conditions": current_branch_conds,
            "outputs": current_branch_outs
        })
    if current_block_branches:
        blocks.append({"branches": current_block_branches})
        
    return blocks


def build_branch_cond_expression(cond_rows: list[dict]) -> str:
    parts = []
    for cr in cond_rows:
        cond_str = cr.get("条件 / 項目マッピング", "") or cr.get("条件", "") or cr.get("cond", "")
        cond_type = _detect_condition_type(cond_str)
        if cond_type == "ELSE":
            continue
            
        # Collect all fields
        fields = []
        for raw_key, fallback_key in [("該当項目名_1", "src_field"), ("該当項目名_2", "src_field2"), ("該当項目名_3", "src_field3")]:
            val = cr.get(raw_key) or cr.get(fallback_key)
            if val:
                fields.append(val)
                
        val_parsed = _parse_condition_value(cond_str)
        # Clean up prefixes
        val_parsed = re.sub(r'^(すべて|いずれも|上記以外で|上記以外かつ|上記以外)', '', val_parsed).strip()
        val_parsed = val_parsed.strip('"\'')
        
        for f in fields:
            if cond_type == "LIKE":
                kw = _extract_like_keyword(cond_str)
                parts.append(f"T1.[{f}] LIKE '%{kw}%'")
            elif cond_type == "NOT_LIKE":
                kw = _extract_like_keyword(cond_str)
                parts.append(f"T1.[{f}] NOT LIKE '%{kw}%'")
            elif cond_type == "EQUALS" and val_parsed == "-":
                parts.append(f"Nz(Replace(T1.[{f}], '-', ''), '') = ''")
            else:
                parts.append(f"T1.[{f}] = '{val_parsed}'")
                
    if not parts:
        return "True"
    return " AND ".join(parts)


def compile_post_update_block(block: dict, prefix: str, target_table: str, sub_name: str, sheet_name: str) -> str:
    code = ""
    for branch in block["branches"]:
        if not branch["conditions"]:
            continue
        cond_row = branch["conditions"][0]
        cond_text = cond_row.get("条件 / 項目マッピング", "") or cond_row.get("条件", "") or cond_row.get("cond", "")
        
        # Extract the value (e.g., '個人' or '法人')
        val_match = re.search(r'["\'門](個人|法人)["\'門]?', cond_text)
        val = val_match.group(1) if val_match else ""
        if not val:
            if "個人" in cond_text:
                val = "個人"
            elif "法人" in cond_text:
                val = "法人"
        if not val:
            continue
            
        # Build SET clause
        set_parts = []
        for out_row in branch["outputs"]:
            f = out_row.get("項目名") or out_row.get("DB_Field")
            if not f:
                continue
            shori = out_row.get("処理", "")
            if shori == "削除":
                set_parts.append(f"T.[{f}] = ''")
            elif shori == "固定値出力":
                out_val = out_row.get("該当項目名_1", "").strip('\'"')
                set_parts.append(f"T.[{f}] = '{out_val}'")
            else:
                out_val = out_row.get("該当項目名_1", "").strip('\'"')
                set_parts.append(f"T.[{f}] = '{out_val}'")
                
        if not set_parts:
            continue
            
        set_str = ", ".join(set_parts)
        
        sql_lines = [
            f"UPDATE {target_table} AS T ",
            f"SET {set_str} ",
            f"WHERE T.ID LIKE '%{prefix}_%' AND T.kind_id = '{val}'; "
        ]
        code += _sql(sql_lines)
        code += f"    db.Execute SQL\n"
        code += f'    log_write "{sub_name}:{sheet_name} → {cond_text}"\n\n'
    return code

def _gen_special_processing(sentences: list, source_table: str, target_table: str,
                             prefix: str, sheet_name: str, sub_name: str) -> str:
    if not sentences:
        return ""

    sheet_context = {
        "sheet_name": sheet_name,
        "source_table": source_table,
        "target_table": target_table,
        "prefix": prefix,
        "sub_name": sub_name,
    }

    # Parse all sentences ONCE to preserve branch object identities
    parsed_sentences = []
    for sentence in sentences:
        rows = sentence.get("rows", [])
        blocks = parse_rows_into_blocks(rows) if rows else []
        parsed_sentences.append((sentence, blocks))

    # ── Helper: sắp xếp 契約者N theo thứ tự ưu tiên 1→3→2 ──────────────
    def _con_priority(r: dict) -> int:
        f = r.get("該当項目名_1", "")
        if "契約者1" in f: return 0
        if "契約者3" in f: return 1
        if "契約者2" in f: return 2
        return 9

    def _build_sql_condition_generic(cond_row: dict, con_idx_map: dict[str, str]) -> str:
        cond_str = cond_row.get("条件 / 項目マッピング", "") or cond_row.get("条件", "") or cond_row.get("cond", "")
        cond_type = _detect_condition_type(cond_str)
        if cond_type == "ELSE":
            return "True"
        
        # Collect all fields
        fields = []
        for raw_key, fallback_key in [("該当項目名_1", "src_field"), ("該当項目名_2", "src_field2"), ("該当項目名_3", "src_field3")]:
            val = cond_row.get(raw_key) or cond_row.get(fallback_key)
            if val:
                fields.append(val)
                
        # If the first field has a con index, use that
        if fields and fields[0] in con_idx_map:
            idx = con_idx_map[fields[0]]
            return f'" & con{idx} & "'
            
        val_parsed = _parse_condition_value(cond_str)
        # Clean up prefixes
        val_parsed = re.sub(r'^(すべて|いずれも|上記以外で|上記以外かつ|上記以外)', '', val_parsed).strip()
        val_parsed = val_parsed.strip('"\'')
        
        # Generate field expressions
        parts = []
        for f in fields:
            if cond_type == "LIKE":
                kw = _extract_like_keyword(cond_str)
                parts.append(f"T1.[{f}] LIKE '%{kw}%'")
            elif cond_type == "NOT_LIKE":
                kw = _extract_like_keyword(cond_str)
                parts.append(f"T1.[{f}] NOT LIKE '%{kw}%'")
            elif cond_type == "EQUALS" and val_parsed == "-":
                parts.append(f"Nz(Replace(T1.[{f}], '-', ''), '') = ''")
            else:
                parts.append(f"T1.[{f}] = '{val_parsed}'")
                
        if not parts:
            return "True"
            
        if cond_type == "OR_LIST":
            return " OR ".join(parts)
        elif cond_type == "AND_ALL":
            return " AND ".join(parts)
        else:
            return " AND ".join(parts)

    # ── PASS 1: Scan tất cả sentences để xác định shared Dim con ──────────
    priority_sentence = None
    priority_block = None
    priority_branches = []
    
    for sentence, blocks in parsed_sentences:
        for block in blocks:
            p_branches = [
                b for b in block["branches"]
                if any("入居" in cr.get("該当項目名_1", "") for cr in b["conditions"])
            ]
            if len(p_branches) >= 2:
                priority_sentence = sentence
                priority_block = block
                priority_branches = p_branches
                break
        if priority_block:
            break

    # ── Build Dim con declarations ────────────────────────────────────────
    dim_con_code = ""
    con_idx_map: dict[str, str] = {}   # branch_index -> con_idx
    
    if priority_branches:
        dim_con_code += "    ' Dim con — điều kiện ưu tiên 契約者1→3→2 (khai báo 1 lần)\n"
        
        # Sort and map variables
        # Contractor 1 -> con1
        # Contractor 3 -> con3
        # Contractor 2 -> con2
        for b in priority_branches:
            conds_str = " ".join(cr.get("該当項目名_1", "") for cr in b["conditions"])
            if "契約者1" in conds_str:
                idx = "1"
            elif "契約者3" in conds_str:
                idx = "3"
            elif "契約者2" in conds_str:
                idx = "2"
            else:
                continue
                
            cond_expr = build_branch_cond_expression(b["conditions"])
            dim_con_code += f"    Dim con{idx} As String\n"
            dim_con_code += f'    con{idx} = "({cond_expr})"\n'
            
            # Map this branch identity to its con{idx}
            con_idx_map[id(b)] = idx
            
        dim_con_code += "\n"

    # ── PASS 2: Sinh code cho từng sentence ───────────────────────────────
    body = ""
    post_update_branches = []
    target_fields_all = []

    for sentence, blocks in parsed_sentences:
        rows = sentence.get("rows", [])
        if not rows:
            continue
            
        # Check if the sentence is about end_date
        is_end_date = False
        for r in rows:
            if r.get("項目名") == "end_date" or r.get("DB_Field") == "end_date":
                cond = r.get("条件 / 項目マッピング", "") or r.get("条件", "") or r.get("該当項目名_1", "") or ""
                if "処理対象日" in cond:
                    is_end_date = True
                    break
                    
        if is_end_date:
            if "end_date" not in target_fields_all:
                target_fields_all.append("end_date")
            body += "    'end_date: 処理対象日(yyyymm) → 前月末の日付\n"
            body += "    Dim targetDate As Date\n"
            body += "    Dim endOfPreviousMonth As Date\n"
            body += "    Dim sqlDate As String\n\n"
            body += '    targetDate = DLookup("DATE_FROM", "T_BUF_DATE")\n'
            body += "    endOfPreviousMonth = DateSerial(Year(targetDate), Month(targetDate), 0)\n"
            body += '    sqlDate = Format(endOfPreviousMonth, "\'yyyy-mm-dd\'")\n\n'
            sql_lines = [
                f"UPDATE {target_table} As T ",
                f"SET T.[end_date] = \" & sqlDate & \" ",
                f"WHERE T.ID LIKE '{prefix}_%'; "
            ]
            body += _sql(sql_lines)
            body += f"    db.Execute SQL\n"
            body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: end_date"\n\n'
            continue

        # Check if it is post-update sentence
        is_post_update = False
        for block in blocks:
            if any(
                any("上記で出力したデータ" in cr.get("条件 / 項目マッピング", "") or "上記で出力したデータ" in cr.get("条件", "") or 
                    "出力された" in cr.get("条件 / 項目マッピング", "") or "出力された" in cr.get("条件", "") or
                    "完了後" in cr.get("条件 / 項目マッピング", "") or "完了後" in cr.get("条件", "")
                    for cr in b["conditions"])
                for b in block["branches"]
            ):
                is_post_update = True
                break
                
        if is_post_update:
            for block in blocks:
                post_update_branches.extend(block["branches"])
            continue

        # Check if the sentence should be handled as parsed blocks
        if len(blocks) > 0 and (len(blocks[0]["branches"]) >= 2 or any(b.get("処理") == "削除" for b in rows)):
            for block in blocks:
                # Standard block compilation
                src_tbl = _clean_table_name(
                    next((r.get("該当ファイル名", source_table) for r in rows if r.get("該当ファイル名")), source_table)
                )
                
                # Collect all target fields in this block
                target_fields = []
                for b in block["branches"]:
                    for out_row in b["outputs"]:
                        f = out_row.get("項目名") or out_row.get("DB_Field")
                        if f and f not in target_fields:
                            target_fields.append(f)
                            if f not in target_fields_all:
                                target_fields_all.append(f)
                            
                # Check for end_date special pattern
                if "end_date" in target_fields:
                    target_fields.remove("end_date")
                    if not target_fields:
                        continue
                        
                if len(block["branches"]) == 1:
                    b = block["branches"][0]
                    cond_expr = build_branch_cond_expression(b["conditions"])
                    
                    set_parts = []
                    for f in target_fields:
                        out_row = next((r for r in b["outputs"] if r.get("項目名") == f or r.get("DB_Field") == f), None)
                        if not out_row:
                            continue
                            
                        out_field = out_row.get("該当項目名_1", "")
                        if not out_field and f == "tag":
                            for cr in b["conditions"]:
                                cond_field = cr.get("該当項目名_1", "")
                                if cond_field:
                                    out_field = cond_field
                                    break
                                    
                        out_expr = f"T1.[{out_field}]"
                        if out_row.get("処理") == "固定値出力":
                            out_expr = f"'{out_field.strip('\'\"')}'"
                            
                        if _is_hyphen_null_mapping(out_row):
                            out_expr = f"IIF(Nz(Replace({out_expr}, '-', ''), '') = '', NULL, {out_expr})"
                            
                        set_parts.append(f"T.[{f}] = {out_expr}")
                        
                    if set_parts:
                        set_clause = ",\n    ".join(set_parts)
                        sql_lines = [
                            f"UPDATE {target_table} AS T ",
                            f"INNER JOIN {src_tbl} AS T1 ON T.ID = ('{prefix}_' & T1.ID) ",
                            f"SET {set_clause} ",
                            f"WHERE {cond_expr}; "
                        ]
                        body += _sql(sql_lines)
                        body += f"    db.Execute SQL\n"
                        fields_log = ",".join(target_fields)
                        body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: {fields_log}"\n\n'
                        
                else:
                    for f in target_fields:
                        switch_lines = [
                            f"UPDATE {target_table} AS T ",
                            f"INNER JOIN {src_tbl} AS T1 ON T.ID = ('{prefix}_' & T1.ID) ",
                            f"SET T.[{f}] = SWITCH( "
                        ]
                        
                        has_default_branch = False
                        for b in block["branches"]:
                            # Determine condition expression (conX variable or raw SQL)
                            if id(b) in con_idx_map:
                                idx = con_idx_map[id(b)]
                                cond_expr = f'" & con{idx} & "'
                            else:
                                cond_expr = build_branch_cond_expression(b["conditions"])
                            if cond_expr == "True":
                                has_default_branch = True
                                
                            # Find output row for this field
                            out_row = next((r for r in b["outputs"] if r.get("項目名") == f or r.get("DB_Field") == f), None)
                            if not out_row:
                                out_expr = "NULL"
                            else:
                                out_field = out_row.get("該当項目名_1", "")
                                if out_row.get("処理") == "固定値出力":
                                    val_clean = out_field.strip('\'"')
                                    out_expr = f"'{val_clean}'"
                                else:
                                    out_expr = f"T1.[{out_field}]"
                                    
                                if _is_hyphen_null_mapping(out_row):
                                    out_expr = f"IIF(Nz(Replace({out_expr}, '-', ''), '') = '', NULL, {out_expr})"
                                    
                            switch_lines.append(f'    {cond_expr}, {out_expr}, ')
                            
                        if not has_default_branch:
                            switch_lines.append("    True, NULL ); ")
                        else:
                            switch_lines[-1] = switch_lines[-1].rstrip(", ") + " ); "
                            
                        body += _sql(switch_lines)
                        body += f"    db.Execute SQL\n"
                        body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: {f}"\n\n'
            continue

        # --- Fallback: parse conditions list ---
        field = ""
        conditions = []
        for row in rows:
            f = row.get("項目名", "")
            if f:
                field = f
            conditions.append({
                "field":      f or field,
                "shori":      row.get("処理", ""),
                "cond":       row.get("条件 / 項目マッピング", "") or row.get("条件", ""),
                "src_field":  row.get("該当項目名_1", ""),
                "src_field2": row.get("該当項目名_2", ""),
                "src_field3": row.get("該当項目名_3", ""),
                "src_file":   _clean_table_name(row.get("該当ファイル名", source_table)),
            })

        if not field or not conditions:
            from compiler.ai_fallback import ai_fallback_generate
            body += ai_fallback_generate(sentence, sheet_context)
            continue

        if field not in target_fields_all:
            target_fields_all.append(field)

        switch_conds = [c for c in conditions if c["shori"] == "条件分岐"]
        fixed_vals   = [c for c in conditions if c["shori"] == "固定値出力"]
        direct_vals  = [c for c in conditions if c["shori"] == "そのまま出力"]
        null_vals    = [c for c in conditions if c["shori"] == "出力なし"]

        # ── Check for "-" của trường hợp 出力なし pair (HYPHEN_TO_NULL in 特殊処理) ──
        hyphen_null_pairs = []
        clean_switch_conds = []
        clean_results = []

        result_list_all = []
        for c in conditions:
            if c["shori"] == "条件分岐":
                result_list_all.append({"cond": c, "result": None})
            elif c["shori"] in ("固定値出力", "そのまま出力", "出力なし") and result_list_all:
                for entry in reversed(result_list_all):
                    if entry["result"] is None:
                        entry["result"] = c
                        break

        for entry in result_list_all:
            cond_entry = entry["cond"]
            result_entry = entry["result"]
            parsed_val = _parse_condition_value(cond_entry["cond"])
            if parsed_val == "-" and result_entry and result_entry["shori"] == "出力なし":
                hyphen_null_pairs.append(entry)
            elif result_entry:
                clean_switch_conds.append(cond_entry)
                clean_results.append(result_entry)

        # Identify fields that have hyphen-null handling
        hyphen_null_fields = {
            entry["cond"]["src_field"]
            for entry in hyphen_null_pairs
            if entry["cond"].get("src_field")
        }

        # If we only have hyphen-null pairs and exactly one direct output,
        # generate IIF pattern instead of SWITCH
        if hyphen_null_pairs and not clean_switch_conds and len(direct_vals) == 1:
            src_f = direct_vals[0]["src_field"]
            sql_lines = [
                f"UPDATE {target_table} AS T ",
                f"INNER JOIN {source_table} AS T1 ON T.ID = ('{prefix}_' & T1.ID) ",
                f"SET T.[{field}] = IIF(Nz(Replace(T1.[{src_f}], '-', ''), '') = '', NULL, T1.[{src_f}]) ",
            ]
            body += _sql(sql_lines)
            body += f'    db.Execute SQL\n'
            body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: {field}"\n\n'
            continue

        if clean_switch_conds and clean_results:
            switch_lines = [
                f"UPDATE {target_table} AS T ",
                f"INNER JOIN {source_table} AS T1 ON T.ID = ('{prefix}_' & T1.ID) ",
                f"SET T.[{field}] = SWITCH( ",
            ]
            for sc, fv in zip(clean_switch_conds, clean_results):
                cond_expr = _build_sql_condition_generic(sc, {})
                result_v = fv["src_field"].strip('"').strip("'")

                if fv["shori"] == "固定値出力":
                    switch_lines.append(f"    ({cond_expr}), '{result_v}', ")
                else:
                    if result_v in hyphen_null_fields:
                        switch_lines.append(f"    ({cond_expr}), IIF(Nz(Replace(T1.[{result_v}], '-', ''), '') = '', NULL, T1.[{result_v}]), ")
                    else:
                        switch_lines.append(f"    ({cond_expr}), T1.[{result_v}], ")

            switch_lines.append("    True, NULL ); ")
            body += _sql(switch_lines)
            body += f'    db.Execute SQL\n'
            body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: {field}"\n\n'

        elif switch_conds and len(switch_conds) == 1:
            sc = switch_conds[0]
            where = _build_sql_condition_generic(sc, {})

            result_rows = [c for c in conditions if c["shori"] == "そのまま出力"]
            result_field = result_rows[0]["src_field"] if result_rows else sc["src_field"]

            if result_field in hyphen_null_fields:
                set_clause = f"T.[{field}] = IIF(Nz(Replace(T1.[{result_field}], '-', ''), '') = '', NULL, T1.[{result_field}])"
            else:
                set_clause = f"T.[{field}] = T1.[{result_field}]"

            sql_lines = [
                f"UPDATE {target_table} AS T ",
                f"INNER JOIN {source_table} AS T1 ON T.ID = ('{prefix}_' & T1.ID) ",
                f"SET {set_clause} ",
                f"WHERE {where}; ",
            ]
            body += _sql(sql_lines)
            body += f'    db.Execute SQL\n'
            body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: {field}"\n\n'

        else:
            from compiler.ai_fallback import ai_fallback_generate
            body += ai_fallback_generate(sentence, sheet_context)

    # ── PASS 3: Generate consolidated post-update logic ───────────────────
    if post_update_branches:
        # Group by '個人' or '法人'
        grouped_outputs = {"個人": [], "法人": []}
        for b in post_update_branches:
            cond_text = ""
            if b["conditions"]:
                cond_text = b["conditions"][0].get("条件 / 項目マッピング", "") or b["conditions"][0].get("条件", "") or ""
            val_match = re.search(r'["\'門](個人|法人)["\'門]?', cond_text)
            val = val_match.group(1) if val_match else ""
            if not val:
                if "個人" in cond_text:
                    val = "個人"
                elif "法人" in cond_text:
                    val = "法人"
            if val in grouped_outputs:
                grouped_outputs[val].extend(b["outputs"])
                
        for val in ["個人", "法人"]:
            outputs = grouped_outputs[val]
            if not outputs:
                continue
            set_parts = []
            seen_fields = set()
            for out_row in outputs:
                f = out_row.get("項目名") or out_row.get("DB_Field")
                if not f or f in seen_fields:
                    continue
                seen_fields.add(f)
                shori = out_row.get("処理", "")
                if shori == "削除":
                    set_parts.append(f"T.[{f}] = ''")
                elif shori in ("固定値出力", "そのまま出力"):
                    out_val = out_row.get("該当項目名_1", "").strip('\'"')
                    set_parts.append(f"T.[{f}] = '{out_val}'")
                else:
                    out_val = out_row.get("該当項目名_1", "").strip('\'"')
                    set_parts.append(f"T.[{f}] = '{out_val}'")
            if set_parts:
                set_str = ", ".join(set_parts)
                sql_lines = [
                    f"UPDATE {target_table} AS T ",
                    f"SET {set_str} ",
                    f"WHERE T.ID LIKE '%{prefix}_%' AND T.kind_id = '{val}'; "
                ]
                body += _sql(sql_lines)
                body += f"    db.Execute SQL\n"
                body += f'    log_write "{sub_name}:{sheet_name} → 上記で出力したデータの個人・法人区分が\'{val}\'の場合"\n\n'

    # ── PASS 4: Generate legacy_id duplicate deletion ─────────────────────
    if "legacy_id" in target_fields_all:
        dedup_sql = [
            f"DELETE FROM {target_table} ",
        ]
        if target_table == "Account":
            dedup_sql.append(f"WHERE ID LIKE '{prefix}_%' ")
            dedup_sql.append(f"AND ID NOT IN ( ")
            dedup_sql.append(f"   SELECT MIN(ID) ")
            dedup_sql.append(f"   FROM {target_table} ")
            dedup_sql.append(f"   WHERE ID LIKE '{prefix}_%' ")
            dedup_sql.append(f"   GROUP BY legacy_id ")
            dedup_sql.append(f") ")
        else:
            dedup_sql.append(f"WHERE ID NOT IN ( ")
            dedup_sql.append(f"   SELECT MIN(ID) ")
            dedup_sql.append(f"   FROM {target_table} ")
            dedup_sql.append(f"   GROUP BY legacy_id ")
            dedup_sql.append(f") ")
            
        body += "\n    ' duplicate legacy_id\n"
        body += _sql(dedup_sql)
        body += f"    db.Execute SQL\n"
        body += f'    log_write "{sub_name}:{sheet_name} → delete duplicate"\n\n'

    return "    '特殊処理\n\n" + dim_con_code + body


def compile_sheet(sheet_raw: dict) -> str:
    """Sinh mã VBA từ một sheet_raw dict."""
    meta = sheet_raw.get("_meta", {})
    sheet_name = meta.get("source_sheet", "Sheet")
    sections   = sheet_raw.get("sections", {})

    filenames    = sections.get("ファイル名", [])
    output_cond  = sections.get("出力条件", {})
    normal_proc  = sections.get("通常処理", {})
    special_proc = sections.get("特殊処理", {})

    source_table          = _primary_source_table(filenames)
    target_table, sub_name = _infer_target(sheet_name)
    prefix                = _infer_prefix(sheet_name)

    oc_sentences = output_cond.get("sentences", [])
    np_rows      = normal_proc.get("rows", [])
    sp_sentences = special_proc.get("sentences", [])

    # Lấy flg_reset và OC code từ classifier (không tự detect nữa)
    flg_reset, oc_code = _gen_output_conditions(
        oc_sentences, source_table, sheet_name, sub_name, target_table
    )

    # Build code
    code  = _gen_header(sheet_name, target_table, sub_name)
    code += _gen_delete_existing(target_table, prefix)
    code += _gen_section_sep(sheet_name)
    code += _gen_flg_init(source_table, flg_reset)
    code += oc_code
    code += _gen_normal_processing(np_rows, source_table, target_table, prefix, sheet_name, sub_name)
    code += _gen_special_processing(sp_sentences, source_table, target_table, prefix, sheet_name, sub_name)
    code += _gen_flg_cleanup(source_table)
    code += "\n    End If\n"
    code += _gen_footer(target_table, sub_name)

    return code



def compile_from_file(json_path: str) -> str:
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)
    return compile_sheet(data)
