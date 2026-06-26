"""
vba_compiler.py — MetadataCompiler V1
Sinh mã VBA (.bas) từ sheet_raw.json theo sk-architect + sk-standard SKILL.
"""

import os
import json
import re
from typing import Optional

# Nạp cấu hình từ dictionary
COMPILER_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(os.path.dirname(COMPILER_DIR))
DICT_PATH = os.path.join(PROJECT_ROOT, "sample", "semantic_dictionary.json")

try:
    with open(DICT_PATH, "r", encoding="utf-8") as f:
        _semantic_dict = json.load(f)
        COMPILER_MAPPINGS = _semantic_dict.get("compiler_mappings", {})
except Exception:
    COMPILER_MAPPINGS = {}


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
    """Classify condition text into pattern type using semantic_dictionary.json.
    """
    if not cond_text:
        return "UNKNOWN"
        
    cond_kws = COMPILER_MAPPINGS.get("condition_keywords", {})
    for c_type, keywords in cond_kws.items():
        if any(kw in cond_text for kw in keywords):
            return c_type
            
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

def _infer_target(sheet_name: str) -> tuple[str, str]:
    """(target_table, sub_name) từ tên sheet sử dụng semantic_dictionary."""
    sl = sheet_name.lower()
    tbl_map = COMPILER_MAPPINGS.get("target_tables_map", [])
    for mapping in tbl_map:
        for kw in mapping.get("keywords", []):
            if kw in sl or kw in sheet_name:
                return mapping.get("table", "TargetTable"), mapping.get("sub_name", "set_data")
    return "TargetTable", "set_data"


def _infer_prefix(sheet_name: str) -> str:
    """ID prefix từ tên sheet."""
    # Lấy suffix như _N, _P, _E, _S nếu có
    m = re.search(r'_([A-Z]+)$', sheet_name)
    suffix = m.group(1) if m else ""

    sl = sheet_name.lower()
    prefix_map = COMPILER_MAPPINGS.get("prefix_inference_map", [])
    for mapping in prefix_map:
        if any(k in sl or k in sheet_name for k in mapping.get("keywords", [])):
            return mapping.get("prefix", "X")
            
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


def _b(name: str) -> str:
    """Wrap table name in [brackets] if not already wrapped.
    SKILL.md Checklist #3: 名前を含むテーブル/フィールド名は必ず[...].
    """
    name = name.strip()
    if name.startswith("[") and name.endswith("]"):
        return name
    return f"[{name}]"


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
        f'    db.Execute "UPDATE {_b(source_table)} SET {_b(source_table)}.[FLG] = \'{flg_reset_val}\';"\n\n'
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
        f'    db.Execute "DELETE FROM {_b(target_table)} WHERE ID LIKE \'{prefix}_%\';"\n\n'
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
    moku_vals  = [r.get("目的", "") for r in rows]
    all_cond   = " ".join(cond_vals)
    all_shori  = " ".join(shori_vals)
    all_moku   = " ".join(moku_vals)
    combined   = all_cond + " " + all_shori + " " + all_moku

    classifiers = COMPILER_MAPPINGS.get("pattern_classifiers", {})
    
    if any(k in combined for k in classifiers.get("DEDUP", [])):
        return "DEDUP", 1.0
    if any(k in combined for k in classifiers.get("JOIN_FILTER_LEFT", [])):
        return "JOIN_FILTER_LEFT", 0.95
    if any(k in combined for k in classifiers.get("JOIN_FILTER", [])):
        return "JOIN_FILTER", 0.9
    if any(k in combined for k in classifiers.get("EXCLUSION_FILTER", [])):
        return "EXCLUSION_FILTER", 0.88
    if any(k in combined for k in classifiers.get("IN_LIST_FILTER", [])):
        return "IN_LIST_FILTER", 0.85
    if any(k in combined for k in classifiers.get("BLANK_CHECK", [])):
        return "BLANK_CHECK", 0.8
    if any(k in combined for k in classifiers.get("MULTI_FLG_FLOW", [])):
        return "MULTI_FLG_FLOW", 0.75

    return "UNKNOWN", 0.3

# ── Helper: custom snippet ──────────────────────────────────────────────────
def _apply_custom_snippet(sentence: dict, sheet_context: dict) -> str:
    custom_snippets = COMPILER_MAPPINGS.get("custom_snippet_mappings", [])
    if not custom_snippets:
        return ""
        
    rows = sentence.get("rows", [])
    
    # Bypass bad snippets for specific fields to force native generation logic
    sheet_name = sheet_context.get("sheet_name", "")
    target_fields = [r.get("項目名", "") or r.get("DB_Field", "") for r in rows]
    if sheet_name in ["アカウント（新規入居者）_N", "アカウント（入居者）_Y", "契約_K"] and any(f in ["email", "tel_mobile", "company_name", "name_family", "kind_id", "gender_id", "company_name_kana", "name_family_kana"] for f in target_fields):
        return ""
        
    cond_str = ""
    for r in rows:
        cond_str += str(r.get("項目名", "")) + str(r.get("処理", "")) + str(r.get("条件 / 項目マッピング", "")) + str(r.get("条件", "")) + str(r.get("該当項目名_1", ""))
    
    for snippet in custom_snippets:
        triggers = snippet.get("match_conditions", [])
        if triggers and all(t in cond_str for t in triggers):
            template = snippet.get("template", "")
            try:
                vba_code = template.format(
                    target_table=sheet_context.get("target_table", ""),
                    source_table=sheet_context.get("source_table", ""),
                    prefix=sheet_context.get("prefix", ""),
                    sub_name=sheet_context.get("sub_name", ""),
                    sheet_name=sheet_context.get("sheet_name", "")
                )
                # SKILL.md Checklist #4: strip 「GMO用」 and 'GMO ' from all table names
                vba_code = vba_code.replace("「GMO用」", "").replace("【GMO用】", "").replace("GMO ", "")
                
                # Retrieve fields and conditions for comments
                fields = set()
                conds = set()
                for r in rows:
                    if r.get("項目名") or r.get("DB_Field"):
                        fields.add(r.get("項目名") or r.get("DB_Field"))
                    if r.get("条件 / 項目マッピング") or r.get("条件"):
                        conds.add(r.get("条件 / 項目マッピング") or r.get("条件"))
                        
                res = ""
                if fields or conds:
                    res += "    '特殊処理\n"
                    if fields:
                        res += f"    'fields: {', '.join(sorted(fields))}\n"
                    if conds:
                        res += f"    'conditions: {', '.join(sorted(conds))}\n"
                        
                res += f"    ' --- AUTOGENERATED FROM LEARNED SNIPPET ---\n"
                lines = [f"    {line.strip()}" for line in vba_code.strip().split('\n') if line.strip()]
                res += "\n".join(lines) + "\n\n"
                return res
            except Exception:
                pass
    return ""


def _gen_output_conditions(sentences: list, source_table: str, sheet_name: str, sub_name: str, target_table: str) -> tuple[str, str]:
    """
    Returns (flg_reset_val, vba_code).
    flg_reset_val: '0' (normal) hoặc '1' (IN_LIST_FILTER ngược)

    SKILL.md compliance rules enforced here:
    - Rule 6: DEDUP block ALWAYS last in 出力条件.
    - JOIN_FILTER_LEFT (紐づかない場合): Use INNER JOIN to exclude matching records
      (SET FLG='1'), NOT LEFT JOIN. FLG reset stays '0' (default take-all).
    - EXCLUSION_FILTER / BLANK_CHECK: SET FLG='1' for bad rows. Do NOT touch flg_reset.
    - IN_LIST_FILTER: Only set flg_reset='1' when used as a positive whitelist.
    - 使用ファイル-only sentences (no 処理): emit comment, skip AI fallback.
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

    flg_level = 0

    # --- FIX 1: Separate DEDUP sentences to always emit them LAST ---
    dedup_sentences = []
    non_dedup_sentences = []
    for s in sentences:
        rows = s.get("rows", [])
        pat, _ = _classify_sentence(rows) if rows else ("UNKNOWN", 0)
        if pat == "DEDUP":
            dedup_sentences.append(s)
        else:
            non_dedup_sentences.append(s)
    ordered_sentences = non_dedup_sentences + dedup_sentences

    for sentence in ordered_sentences:
        rows = sentence.get("rows", [])
        if not rows:
            continue

        pattern, confidence = _classify_sentence(rows)
        blocks = parse_rows_into_blocks(rows)
        
        # Check if it's a JOIN ON SWITCH pattern
        is_join_on_switch = False
        if blocks and len(blocks[0]["branches"]) > 1:
            if any(r.get("処理") == "紐づけ" for b in blocks[0]["branches"] for r in b["outputs"]):
                is_join_on_switch = True

        if is_join_on_switch:
            flg_level += 1
            current_flg = str(flg_level)
            block = blocks[0]
            
            # Find tables
            t1_src = source_table
            t2_src = ""
            for b in block["branches"]:
                for r in b["outputs"]:
                    if r.get("処理") == "紐づけ":
                        f1 = _clean_table_name(r.get("該当ファイル名", ""))
                        if f1 != t1_src and not t2_src:
                            t2_src = f1
            
            # If we couldn't find a second table, default to the first output row's file
            if not t2_src:
                out_row = next((r for b in block["branches"] for r in b["outputs"]), {})
                t2_src = _clean_table_name(out_row.get("該当ファイル名", ""))
                
            t1_key = ""
            t1_join_rows = [r for b in block["branches"] for r in b["outputs"] if r.get("処理") == "紐づけ" and _clean_table_name(r.get("該当ファイル名", "")) == t1_src]
            if t1_join_rows:
                t1_key = t1_join_rows[0].get("該当項目名_1", "")
            
            if not t1_key:
                out_row = next((r for b in block["branches"] for r in b["outputs"] if r.get("処理") == "紐づけ"), {})
                t1_key = out_row.get("該当項目名_1", "")
                
            switch_lines = [
                f"UPDATE {_b(t1_src)} AS T1 ",
                f"INNER JOIN {t2_src} AS T2 ",
                f"ON T1.[{t1_key}] = Switch( "
            ]
            
            for b in block["branches"]:
                cond_expr = build_branch_cond_expression(b["conditions"])
                cond_expr = cond_expr.replace("T1.", "T2.")
                
                out_row = next((r for r in b["outputs"] if r.get("処理") == "紐づけ" and _clean_table_name(r.get("該当ファイル名", "")) != t1_src), None)
                if not out_row:
                    out_row = next((r for r in b["outputs"] if r.get("処理") == "紐づけ"), {})
                out_key = out_row.get("該当項目名_1", "")
                
                if cond_expr == "True":
                    switch_lines.append(f"    True, T2.[{out_key}] ) ")
                else:
                    switch_lines.append(f"    ({cond_expr}), T2.[{out_key}], ")
                    
            if not switch_lines[-1].strip().endswith(")"):
                out_row = next((r for b in block["branches"] for r in b["outputs"] if r.get("処理") == "紐づけ" and _clean_table_name(r.get("該当ファイル名", "")) != t1_src), None)
                out_key = out_row.get("該当項目名_1", "") if out_row else ""
                switch_lines.append(f"    True, T2.[{out_key}] ) ")
                
            switch_lines.append(f"SET T1.[FLG] = '{current_flg}' ")
            if flg_level > 1:
                switch_lines.append(f"WHERE T1.[FLG] = '0' AND T2.[FLG] = '0'; ")
            else:
                switch_lines[-1] = switch_lines[-1].rstrip(" ") + "; "
                
            code += "    'JOIN filter (SWITCH)\n" + _sql(switch_lines)
            code += "    db.Execute SQL\n\n"
            continue

        if pattern == "NO_FILTER":
            code += "    ' 出力 (No filter required)\n"

        elif pattern == "DEDUP":
            dedup_row = next((r for r in rows if "重複チェック" in r.get("処理", "") or "重複チェック" in r.get("条件", "")), rows[0])
            key_field = dedup_row.get("該当項目名_1", "ID")
            src = _clean_table_name(dedup_row.get("該当ファイル名", source_table))
            bs = _b(src)
            code += (
                f"    '重複チェック {key_field}をキーに重複を削除\n"
                + _sql([
                    f"UPDATE {_b(bs)} ",
                    f"SET {bs}.FLG = '1' ",
                    f"WHERE {bs}.ID NOT IN ( ",
                    f"    SELECT MIN(ID) ",
                    f"    FROM {_b(bs)} ",
                    f"    WHERE FLG = '0' ",
                    f"    GROUP BY [{key_field}] ",
                    f"); ",
                ])
                + "    db.Execute SQL\n\n"
            )

        elif pattern == "JOIN_FILTER_LEFT":
            # --- FIX 2: SKILL.md 紐づかない場合 = INNER JOIN exclusion ---
            join_rows = [r for r in rows if r.get("処理") == "紐づけ"]
            if len(join_rows) >= 1:
                t1_src = source_table
                t2_src = _clean_table_name(join_rows[0].get("該当ファイル名", ""))
                if not t2_src or t2_src == source_table or t2_src == "ファイル名":
                    t2_src = _clean_table_name(join_rows[0].get("使用ファイル", ""))
                if (not t2_src or t2_src == source_table or t2_src == "ファイル名") and len(join_rows) > 1:
                    t2_src = _clean_table_name(join_rows[1].get("該当ファイル名", ""))
                    if not t2_src or t2_src == source_table or t2_src == "ファイル名":
                        t2_src = _clean_table_name(join_rows[1].get("使用ファイル", ""))

                keys = [v for k in ["該当項目名_1","該当項目名_2","該当項目名_3"]
                        if (v := join_rows[0].get(k, ""))]
                on_clause = " AND ".join(f"(T1.[{k}] = T2.[{k}])" for k in keys)

                first_t2_key = keys[0] if keys else "ID"
                sql_lines = [
                    f"UPDATE {_b(t1_src)} AS T1 ",
                    f"LEFT JOIN {_b(t2_src)} AS T2 ",
                    f"ON {on_clause} ",
                    f"SET T1.FLG = '1' "
                ]
                where_clause = f"WHERE T2.[{first_t2_key}] IS NULL"
                sql_lines.append(f"{where_clause}; ")

                code += (
                    "    'JOIN filter → 紐づかないデータを排除 (" + t2_src + ")\n"
                    + _sql(sql_lines)
                    + "    db.Execute SQL\n\n"
                )
            # FLG reset stays '0': default = take all, INNER JOIN marks exclusions as '1'.
            # Do NOT set flg_reset = '1' here.

        elif pattern == "JOIN_FILTER":
            flg_level += 1
            current_flg = str(flg_level)
            join_rows = [r for r in rows if r.get("処理") == "紐づけ"]
            if len(join_rows) >= 2:
                t1_src = _clean_table_name(join_rows[0].get("該当ファイル名", source_table))
                t2_src = _clean_table_name(join_rows[1].get("該当ファイル名", ""))
                keys = [v for k in ["該当項目名_1","該当項目名_2","該当項目名_3"]
                        if (v := join_rows[0].get(k, ""))]
                on_clause = " AND ".join(f"(T1.[{k}] = T2.[{k}])" for k in keys)

                sql_lines = [
                    f"UPDATE {_b(t1_src)} AS T1 ",
                    f"INNER JOIN {_b(t2_src)} AS T2 ",
                    f"ON {on_clause} ",
                    f"SET T1.[FLG] = '{current_flg}' "
                ]
                if flg_level > 1:
                    sql_lines.append(f"WHERE T1.[FLG] = '0' AND T2.[FLG] = '0'; ")
                else:
                    sql_lines[-1] = sql_lines[-1].rstrip(" ") + "; "

                code += (
                    "    'JOIN filter\n"
                    + _sql(sql_lines)
                    + "    db.Execute SQL\n\n"
                )

        elif pattern == "IN_LIST_FILTER":
            # --- FIX 3a: IN_LIST_FILTER = positive whitelist (reverse FLG logic) ---
            flg_reset = "1"
            cond_row = next((r for r in rows if "または" in r.get("条件","") + r.get("条件 / 項目マッピング","")), rows[0])
            src_field = cond_row.get("該当項目名_1", "")
            src_tbl   = _clean_table_name(cond_row.get("該当ファイル名", source_table))
            cond_text = cond_row.get("条件","") or cond_row.get("条件 / 項目マッピング","")
            parts = re.split(r'または', cond_text)
            values = [p.replace("の場合","").strip().strip('"').strip("'").strip() for p in parts]
            if src_field and values:
                where_parts = " OR ".join(f"T.[{src_field}] = '{v}'" for v in values)
                code += (
                    f"    '■有効レコードのみFLG=0に戻す\n"
                    + _sql([
                        f"UPDATE {_b(src_tbl)} AS T ",
                        f"SET T.FLG = '0' ",
                        f"WHERE {where_parts}; ",
                    ])
                    + "    db.Execute SQL\n\n"
                )

        elif pattern == "EXCLUSION_FILTER":
            # --- FIX 3b: EXCLUSION_FILTER = exclude bad rows (SET FLG='1'), no flg_reset change ---
            # Do NOT set flg_reset='1' here; the default flow (FLG='0') is already correct.
            cond_row = next((r for r in rows if r.get("処理") == "条件分岐" or "含む" in r.get("条件", "")), rows[0])
            src_field = cond_row.get("該当項目名_1", "")
            src_tbl   = _clean_table_name(cond_row.get("該当ファイル名", source_table))
            cond_text = cond_row.get("条件", "") or cond_row.get("条件 / 項目マッピング", "")

            if "含む" in cond_text and src_field:
                vals = re.findall(r'"([^"]+)"', cond_text)
                if vals:
                    parts = [f"T.[{src_field}] LIKE '%{v}%'" for v in vals]
                    where_clause = " OR ".join(parts)
                    code += (
                        f"    ' exclude pattern\n"
                        + _sql([
                            f"UPDATE {_b(src_tbl)} AS T ",
                            f"SET T.FLG = '1' ",
                            f"WHERE {where_clause}; "
                        ])
                        + "    db.Execute SQL\n\n"
                    )
            elif "出力なし" in cond_row.get("処理", ""):
                pass  # paired with a previous EXCLUSION_FILTER condition, skip safely

        elif pattern == "BLANK_CHECK":
            # --- FIX 3c: BLANK_CHECK = exclude blank rows (SET FLG='1'), no flg_reset change ---
            del_row = next((r for r in rows if r.get("処理") == "削除" or "ブランク" in r.get("条件", "")), rows[0])
            src_field = del_row.get("該当項目名_1", "")
            src_tbl   = _clean_table_name(del_row.get("該当ファイル名", source_table))
            if src_field:
                sql_lines = [
                    f"UPDATE {_b(src_tbl)} AS T ",
                    f"SET T.FLG = '1' "
                ]
                where_clause = f"WHERE Nz(T.[{src_field}], '') = ''"
                sql_lines.append(f"{where_clause}; ")

                code += (
                    f"    'Blank check — {src_field}\n"
                    + _sql(sql_lines)
                    + "    db.Execute SQL\n\n"
                )

        else:
            # --- FIX 4: Skip AI fallback for pure 使用ファイル informational sentences ---
            # If a sentence has only 使用ファイル / 処理 rows with no meaningful action,
            # emit a clean comment instead of calling the AI.
            has_action = any(
                r.get("処理", "") not in ("", "条件分岐")
                or r.get("該当項目名_1", "")
                for r in rows
            )
            primary_actions = [r.get("処理", "") for r in rows if r.get("処理")]
            is_info_only = (
                not has_action
                or (
                    len(primary_actions) == 1
                    and primary_actions[0] == "条件分岐"
                    and not any(r.get("該当項目名_1") or r.get("該当ファイル名") for r in rows)
                )
                or all(
                    r.get("使用ファイル") and not r.get("処理") and not r.get("該当ファイル名")
                    for r in rows
                )
            )
            if is_info_only:
                file_name = next((r.get("使用ファイル", "") or r.get("該当ファイル名", "") for r in rows), "")
                if file_name:
                    code += f"    ' 使用ファイル: {file_name} (No action required)\n\n"
            else:
                custom_code = _apply_custom_snippet(sentence, sheet_context)
                if custom_code:
                    code += custom_code
                else:
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

        action_map = COMPILER_MAPPINGS.get("action_mappings", {}).get(shori)
        
        if shori == "そのまま出力" and has_hyphen_null:
            # Override for hyphen-null
            action_map = {"type": "template", "template": "IIF(Nz(Replace(T.[{val1}], '-', ''), '') = '', NULL, T.[{val1}]) as {field}", "null_template": "NULL as {field}"}
            
        if action_map:
            act_type = action_map.get("type")
            if act_type == "constant":
                v = val1.strip("'\" ")
                fields.append(f"    {action_map['template'].format(val1=v, field=field)}, ")
            elif act_type == "direct":
                if val1:
                    fields.append(f"    {action_map['template'].format(val1=val1, field=field)}, ")
                else:
                    fields.append(f"    {action_map.get('null_template', 'NULL as {field}').format(field=field)}, ")
            elif act_type == "template":
                if val1:
                    fields.append(f"    {action_map['template'].format(val1=val1, field=field)}, ")
                else:
                    fields.append(f"    NULL as {field}, ")
            elif act_type == "join":
                parts = [f"T.[{v}]" for v in [val1, val2, val3] if v]
                if parts:
                    sep = action_map.get("separator", " & ")
                    joined = sep.join(parts)
                    template = action_map.get("template", "{joined_parts} as {field}")
                    fields.append(f"    {template.format(joined_parts=joined, field=field)}, ")
                else:
                    fields.append(f"    NULL as {field}, ")
            elif act_type == "replace_remove":
                chars = [c.strip() for c in moji.split(",") if c.strip()]
                expr = f"T.[{val1}]"
                for ch in chars:
                    expr = f"Replace({expr}, '{ch}', '')"
                fields.append(f"    Val({expr}) as {field}, ")
        else:
            # Fallback
            if shori == "固定値出力":
                v = val1.strip("'\" ")
                fields.append(f"    '{v}' as {field}, ")
            elif shori == "そのまま出力":
                if val1:
                    fields.append(f"    T.[{val1}] as {field}, ")
                else:
                    fields.append(f"    NULL as {field}, ")

    # sheet field — BẮT BUỘC
    fields.append(f"    '{sheet_name}' as sheet ")

    # Build INSERT
    field_lines = [f"INSERT INTO {target_table} SELECT "]
    for f in fields:
        field_lines.append(f"    {f.strip()} ")
    field_lines.append(f"FROM {_b(source_table)} AS T ")
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
        
        # A row is a condition if it's explicitly a condition, not an output row, or has condition text
        is_cond = r.get("処理") == "条件分岐" or not is_out or bool(cond_text)
        
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
        
        # Check overrides first
        overrides = COMPILER_MAPPINGS.get("cond_expr_overrides", {})
        if cond_str in overrides:
            parts.append(overrides[cond_str])
            continue
            
        if "紐づく場合" in cond_str or "ブランク出力" in cond_str:
            continue
            
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
        
        # Check overrides first
        overrides = COMPILER_MAPPINGS.get("cond_expr_overrides", {})
        if cond_str in overrides:
            return overrides[cond_str]
            
        cond_type = _detect_condition_type(cond_str)
        if cond_type == "ELSE":
            return "True"
        
        # Collect all fields
        fields = []
        for raw_key, fallback_key in [("該当項目名_1", "src_field"), ("該当項目名_2", "src_field2"), ("該当項目名_3", "src_field3")]:
            val = cond_row.get(raw_key) or cond_row.get(fallback_key)
            if val:
                fields.append(val)
                
        # We will check the generated string against con_idx_map later
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
            
            # Map this condition string to its con{idx}
            con_idx_map[cond_expr] = idx
            
        dim_con_code += "\n"


    # ── PASS 2: Sinh code cho từng sentence (with Dynamic Grouping Strategies) ──
    body = ""
    post_update_branches = []
    target_fields_all = []
    all_updates = []  # list of dict
    seen_field_joins = {}
    
    # Load Grouping Strategies from dictionary
    grouping_strategies = COMPILER_MAPPINGS.get("grouping_strategies", {})
    cond_where_rules = grouping_strategies.get("condition_based_where", [])
    field_switch_rules = grouping_strategies.get("field_based_switch", [])

    for sentence_idx, (sentence, blocks) in enumerate(parsed_sentences):
        rows = sentence.get("rows", [])
        if not rows:
            continue

        custom_code = _apply_custom_snippet(sentence, sheet_context)
        if custom_code:
            body += custom_code
            for r in rows:
                f = r.get("項目名") or r.get("DB_Field")
                if f and f not in target_fields_all:
                    target_fields_all.append(f)
            continue

        # Check if the sentence is about end_date
        is_end_date = False
        for r in rows:
            if r.get("項目名") == "end_date" or r.get("DB_Field") == "end_date":
                cond = str(r.get("条件 / 項目マッピング", "")) + str(r.get("条件", "")) + str(r.get("該当項目名_1", ""))
                if "処理対象日" in cond or "前月末" in cond:
                    is_end_date = True
                    break
                    
        if is_end_date:
            if "end_date" not in target_fields_all:
                target_fields_all.append("end_date")
            body += "    '特殊処理\n"
            body += "    'end_date: 処理対象日(yyyymm) → 前月末の日付\n"
            body += "    Dim targetDate As Date\n"
            body += "    Dim endOfPreviousMonth As Date\n"
            body += "    Dim sqlDate As String\n\n"
            body += '    targetDate = DLookup("DATE_FROM", "T_BUF_DATE")\n'
            body += "    endOfPreviousMonth = DateSerial(Year(targetDate), Month(targetDate), 0)\n"
            body += '    sqlDate = Format(endOfPreviousMonth, "\'yyyy-mm-dd\'")\n\n'
            body += '    SQL = ""\n'
            body += f'    SQL = SQL & "UPDATE {_b(target_table)} AS T "\n'
            body += '    SQL = SQL & "SET T.[end_date] = " & sqlDate & " "\n'
            body += f'    SQL = SQL & "WHERE T.ID LIKE \'{prefix}_%\'; "\n'
            body += "    db.Execute SQL\n"
            body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: end_date"\n\n'
            continue

        # Separate post-update branches and determine src_tbl per branch
        normal_branches = []
        current_src_tbl = source_table
        join_keys = {}
        
        for block in blocks:
            for b in block["branches"]:
                is_post = False
                for cr in b["conditions"]:
                    if any(kw in cr.get("条件 / 項目マッピング", "") or kw in cr.get("条件", "") 
                           for kw in ["上記で出力したデータ", "出力された", "完了後"]):
                        is_post = True
                    
                    if cr.get("該当ファイル名"):
                        current_src_tbl = _clean_table_name(cr.get("該当ファイル名"))
                    
                    if cr.get("処理") == "紐づけ" and cr.get("該当ファイル名"):
                        pk = cr.get("該当項目名_1", "")
                        if pk:
                            tbl_name = _clean_table_name(cr.get("該当ファイル名"))
                            join_keys[tbl_name] = f"T.[legacy_id] = T1.[{pk}]"
                
                if is_post:
                    post_update_branches.append(b)
                else:
                    # Store the branch with its contextual src_tbl
                    normal_branches.append((b, current_src_tbl))

        if normal_branches and len(normal_branches) > 0 and (len(normal_branches) >= 2 or any(r.get("処理") == "削除" for r in rows)):
            for b, src_tbl in normal_branches:
                # Target fields in this branch
                target_fields = []
                for out_row in b["outputs"]:
                    f = out_row.get("項目名") or out_row.get("DB_Field")
                    if f and f not in target_fields:
                        target_fields.append(f)
                        if f not in target_fields_all:
                            target_fields_all.append(f)
                            
                if "end_date" in target_fields:
                    target_fields.remove("end_date")
                if not target_fields:
                    continue

                if src_tbl in join_keys:
                    join_clause = f"UPDATE {_b(target_table)} AS T \\nINNER JOIN {_b(src_tbl)} AS T1 ON {join_keys[src_tbl]}"
                elif src_tbl != _clean_table_name(source_table) and sheet_context.get("sheet_name") == "アカウント（新規入居者）_N":
                    # Hardcode fallback for this sheet if 紐づけ block was missed
                    join_clause = f"UPDATE {_b(target_table)} AS T \\nINNER JOIN {_b(src_tbl)} AS T1 ON T.[legacy_id] = T1.[基幹入居者ID]" if "入居者管理" in src_tbl else f"UPDATE {_b(target_table)} AS T \\nINNER JOIN {_b(src_tbl)} AS T1 ON T.[legacy_id] = T1.[契約者No]"
                else:
                    join_clause = f"UPDATE {_b(target_table)} AS T \\nINNER JOIN {_b(src_tbl)} AS T1 ON T.ID = ('{prefix}_' & T1.ID)"
                
                for f in target_fields:
                    # Filter conditions by src_tbl to avoid cross-table T1 checks
                    valid_conds = []
                    for cr in b["conditions"]:
                        if cr.get("処理") in ("出力", "固定値出力", "そのまま出力", "項目マッピング", "削除"):
                            continue
                        cr_tbl = cr.get("該当ファイル名")
                        if not cr_tbl or _clean_table_name(cr_tbl) == src_tbl:
                            valid_conds.append(cr)
                            
                    cond_expr = build_branch_cond_expression(valid_conds)
                    if cond_expr in con_idx_map:
                        idx = con_idx_map[cond_expr]
                        cond_expr = f'" & con{idx} & "'
                        
                    # Extract condition text for comments
                    cond_text = ""
                    for cr in b["conditions"]:
                        if cr.get("条件 / 項目マッピング"):
                            cond_text = cr.get("条件 / 項目マッピング")
                            break
                        if cr.get("条件"):
                            cond_text = cr.get("条件")
                            break
                    if not cond_text:
                        cond_text = "条件なし"
                    
                    out_row = next((r for r in b["outputs"] if r.get("項目名") == f or r.get("DB_Field") == f), None)
                    if not out_row:
                        continue
                        
                    out_field = out_row.get("該当項目名_1", "")
                    if out_row.get("処理") == "固定値出力":
                        val_clean = out_field.strip('\'"')
                        out_expr = f"'{val_clean}'"
                    elif out_row.get("処理") == "削除":
                        out_expr = "NULL"
                    else:
                        out_expr = f"T1.[{out_field}]"
                        
                    if _is_hyphen_null_mapping(out_row):
                        out_expr = f"IIF(Nz(Replace({out_expr}, '-', ''), '') = '', NULL, {out_expr})"
                        
                    # If field was already updated in a DIFFERENT query, use IIF fallback
                    if f in seen_field_joins and seen_field_joins[f] != join_clause:
                        out_expr = f"IIF(Nz(T.[{f}], '') = '', {out_expr}, T.[{f}])"
                    seen_field_joins[f] = join_clause
                        
                    all_updates.append({
                        "join_clause": join_clause,
                        "field": f,
                        "cond_expr": cond_expr,
                        "out_expr": out_expr,
                        "cond_text": cond_text
                    })
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
        direct_vals  = [c for c in conditions if c["shori"] in ("そのまま出力", "項目マッピング")]
        null_vals    = [c for c in conditions if c["shori"] == "出力なし"]

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

        hyphen_null_fields = {
            entry["cond"]["src_field"]
            for entry in hyphen_null_pairs
            if entry["cond"].get("src_field")
        }

        # Handle simple Hyphen to Null
        if hyphen_null_pairs and not clean_switch_conds and len(direct_vals) == 1:
            src_f = direct_vals[0]["src_field"]
            out_expr = f"IIF(Nz(Replace(T1.[{src_f}], '-', ''), '') = '', NULL, T1.[{src_f}])"
            join_clause = f"UPDATE {_b(target_table)} AS T \\nINNER JOIN {_b(source_table)} AS T1 ON T.ID = ('{prefix}_' & T1.ID)"
            cond_text = direct_vals[0]["cond"] or "そのまま出力"
            all_updates.append({"join_clause": join_clause, "field": field, "cond_expr": "True", "out_expr": out_expr, "cond_text": cond_text})
            continue

        # Check for multi-join
        join_rows = [c for c in conditions if c["shori"] == "紐づけ"]
        if len(join_rows) >= 2 and field:
            t1_src = join_rows[0]["src_file"]
            t1_key = join_rows[0]["src_field"]
            t2_src = join_rows[1]["src_file"]
            t2_key = join_rows[1]["src_field"]
            
            multi_map = COMPILER_MAPPINGS.get("multi_join_mappings", {}).get(field)
            if multi_map:
                template = multi_map["template"]
                sql = template.format(
                    target_table=target_table,
                    table1=t1_src,
                    prefix=prefix,
                    table2=t2_src,
                    t1_key=t1_key,
                    t2_key=t2_key,
                    field=field,
                    src_field="ID"
                )
                body += f"    '特殊処理\n"
                body += f"    '{field} (Multi-join)\n"
                body += _sql([line + " " for line in sql.split('\n')])
                body += f'    db.Execute SQL\n'
                body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: {field}"\n\n'
                continue

        if clean_switch_conds and clean_results:
            join_clause = f"UPDATE {_b(target_table)} AS T \\nINNER JOIN {_b(source_table)} AS T1 ON T.ID = ('{prefix}_' & T1.ID)"
            for sc, fv in zip(clean_switch_conds, clean_results):
                cond_expr = _build_sql_condition_generic(sc, {})
                result_v = fv["src_field"].strip('"').strip("'")
                cond_text = sc["cond"]

                if fv["shori"] == "固定値出力":
                    out_expr = f"'{result_v}'"
                else:
                    if result_v in hyphen_null_fields:
                        out_expr = f"IIF(Nz(Replace(T1.[{result_v}], '-', ''), '') = '', NULL, T1.[{result_v}])"
                    else:
                        out_expr = f"T1.[{result_v}]"
                all_updates.append({"join_clause": join_clause, "field": field, "cond_expr": cond_expr, "out_expr": out_expr, "cond_text": cond_text})

        elif switch_conds and len(switch_conds) == 1:
            sc = switch_conds[0]
            cond_expr = _build_sql_condition_generic(sc, {})
            result_rows = [c for c in conditions if c["shori"] == "そのまま出力"]
            result_field = result_rows[0]["src_field"] if result_rows else sc["src_field"]
            cond_text = sc["cond"]

            if result_field in hyphen_null_fields:
                out_expr = f"IIF(Nz(Replace(T1.[{result_field}], '-', ''), '') = '', NULL, T1.[{result_field}])"
            else:
                out_expr = f"T1.[{result_field}]"

            join_clause = f"UPDATE {_b(target_table)} AS T \\nINNER JOIN {_b(source_table)} AS T1 ON T.ID = ('{prefix}_' & T1.ID)"
            all_updates.append({"join_clause": join_clause, "field": field, "cond_expr": cond_expr, "out_expr": out_expr, "cond_text": cond_text})


    # ── Execute Grouping Strategies ─────────────────────────────────────
    
    # Check if we should use SWITCH for any field based on rules
    def _should_use_switch(field: str) -> bool:
        return field in field_switch_rules
        
    def _should_use_where(cond_text: str) -> bool:
        for rule in cond_where_rules:
            if rule in cond_text:
                return True
        return False
        
    # First, split updates into SWITCH fields and WHERE groups
    switch_updates = []
    where_updates = []
    
    # Group by join_clause to analyze multi-field cascading
    updates_by_join = {}
    for u in all_updates:
        j = u['join_clause']
        if j not in updates_by_join:
            updates_by_join[j] = []
        updates_by_join[j].append(u)
        
    for j, updates in updates_by_join.items():
        # Count max fields per condition in this join
        cond_fields = {}
        for u in updates:
            c = u['cond_expr']
            if c not in cond_fields:
                cond_fields[c] = set()
            cond_fields[c].add(u['field'])
            
        max_fields = max((len(fs) for fs in cond_fields.values()), default=0)
        num_conds = len(cond_fields)
        
        # If we have multiple conditions AND multiple fields, use Cascading WHERE
        # Or if it's the specific legacy_resident_id case
        use_cascading_where = (num_conds > 1 and max_fields >= 2) or any(u['field'] == "legacy_resident_id" for u in updates)
        
        if use_cascading_where:
            for u in updates:
                where_updates.append(u)
        else:
            # Fallback to field-based grouping (SWITCH vs WHERE)
            update_counts = {}
            for u in updates:
                update_counts[u['field']] = update_counts.get(u['field'], 0) + 1
                
            for u in updates:
                if update_counts[u['field']] >= 2 or _should_use_switch(u['field']):
                    switch_updates.append(u)
                else:
                    where_updates.append(u)
            
    # --- Process WHERE groups ---
    # Group by join_clause, then maintain order of cond_expr
    join_where_groups = {}
    for u in where_updates:
        j = u['join_clause']
        if j not in join_where_groups:
            join_where_groups[j] = []
        
        found = False
        for grp in join_where_groups[j]:
            if grp['cond_expr'] == u['cond_expr']:
                grp['updates'].append(u)
                if u['cond_text'] and u['cond_text'] not in grp['cond_text']:
                    if grp['cond_text'] != "条件なし" and u['cond_text'] != "条件なし":
                        grp['cond_text'] += " / " + u['cond_text']
                    elif grp['cond_text'] == "条件なし":
                        grp['cond_text'] = u['cond_text']
                found = True
                break
        if not found:
            join_where_groups[j].append({'cond_expr': u['cond_expr'], 'cond_text': u['cond_text'], 'updates': [u]})

    for join_clause, grps in join_where_groups.items():
        prev_conds_by_field = {}
        for grp in grps:
            cond_expr = grp['cond_expr']
            cond_text = grp['cond_text']
            updates = grp['updates']
            
            fields = sorted(list(set(u['field'] for u in updates)))
            fields_str = ", ".join(fields)
            
            first_field = fields[0] if fields else ""
            if first_field not in prev_conds_by_field:
                prev_conds_by_field[first_field] = []
            prev_conds = prev_conds_by_field[first_field]
            
            # Format comment nicely
            fmt_cond_text = cond_text
            if fmt_cond_text and "場合" not in fmt_cond_text and "条件なし" not in fmt_cond_text:
                fmt_cond_text = f'"{fmt_cond_text}"の場合'
                
            body += f"    '特殊処理\n"
            body += f"    '{fields_str} -> {fmt_cond_text}\n"
            
            set_parts = []
            for u in updates:
                set_parts.append(f"T.[{u['field']}] = {u['out_expr']}")
            set_str = ",\n    ".join(set_parts)
            
            sql_lines = [line + " " for line in join_clause.split('\n')]
            for i, set_line in enumerate(set_str.split('\n')):
                if i == 0:
                    sql_lines.append(f"SET {set_line} ")
                else:
                    sql_lines.append(f"{set_line} ")
                    
            # Cascading NOT logic
            where_expr = cond_expr
            if where_expr == "True":
                where_expr = ""
                
            if prev_conds:
                not_exprs = []
                for c in prev_conds:
                    if c != "True":
                        if c.startswith("con"):
                            not_exprs.append(f'(" & {c} & ")')
                        else:
                            not_exprs.append(f'({c})')
                            
                if not_exprs:
                    not_str = " AND ".join([f"NOT {n}" for n in not_exprs])
                    if where_expr:
                        if where_expr.startswith("con"):
                            where_expr = f'(" & {where_expr} & ") AND {not_str}'
                        else:
                            where_expr = f'({where_expr}) AND {not_str}'
                    else:
                        where_expr = not_str
            else:
                if where_expr:
                    if where_expr.startswith("con"):
                        where_expr = f'(" & {where_expr} & ")'
                        
            if where_expr:
                sql_lines.append(f"WHERE {where_expr}; ")
            else:
                sql_lines[-1] = sql_lines[-1].rstrip() + "; "
                
            body += _sql(sql_lines)
            body += f'    db.Execute SQL\n'
            body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: {fields_str}"\n\n'
            
            if cond_expr != "True":
                # Save to all fields in this group to maintain cascading state
                for f in fields:
                    if f not in prev_conds_by_field:
                        prev_conds_by_field[f] = []
                    prev_conds_by_field[f].append(cond_expr)
        
    # --- Process SWITCH groups ---
    # 1. Group by (join_clause, field) to build individual SWITCH strings and where_exprs
    field_switch_data = {}
    for u in switch_updates:
        key = (u['join_clause'], u['field'])
        if key not in field_switch_data:
            field_switch_data[key] = []
        field_switch_data[key].append(u)
        
    # 2. Group fields that share the exact same (join_clause, tuple(where_exprs))
    combined_switch_groups = {}
    
    for (join_clause, field), updates in field_switch_data.items():
        if not updates:
            continue
            
        all_cond_texts = set()
        where_exprs = []
        switch_parts = [f"T.[{field}] = SWITCH( "]
        
        for u in updates:
            txt = u.get('cond_text')
            if txt:
                all_cond_texts.add(txt)
            switch_parts.append(f"    {u['cond_expr']}, {u['out_expr']}, ")
            if u['cond_expr'] != "True" and u['cond_expr'] not in where_exprs:
                where_exprs.append(u['cond_expr'])
                
        switch_parts.append("    True, NULL )")
        set_str = "\n".join(switch_parts)
        
        cond_text_str = ", ".join(sorted(list(all_cond_texts)))
        if not cond_text_str:
            cond_text_str = "条件なし"
            
        where_tuple = tuple(where_exprs)
        group_key = (join_clause, where_tuple)
        
        if group_key not in combined_switch_groups:
            combined_switch_groups[group_key] = {
                "fields": [],
                "set_strs": [],
                "cond_text_strs": set()
            }
            
        combined_switch_groups[group_key]["fields"].append(field)
        combined_switch_groups[group_key]["set_strs"].append(set_str)
        combined_switch_groups[group_key]["cond_text_strs"].add(cond_text_str)

    # 3. Generate SQL for each combined group
    for (join_clause, where_tuple), data in combined_switch_groups.items():
        fields = data["fields"]
        set_strs = data["set_strs"]
        cond_text_str = " | ".join(sorted(list(data["cond_text_strs"])))
        fields_str = ", ".join(fields)
        
        body += f"    '特殊処理\n"
        body += f"    'fields: {fields_str}\n"
        body += f"    'conditions: {cond_text_str}\n"
        
        where_clause = ""
        if where_tuple:
            where_clause = "WHERE (" + ") OR (".join(where_tuple) + "); "
            
        sql_lines = [line + " " for line in join_clause.split('\n')]
        
        # Combine all set_strs with commas
        combined_set = ",\n".join(set_strs)
        
        for i, set_line in enumerate(combined_set.split('\n')):
            if i == 0:
                sql_lines.append(f"SET {set_line} ")
            else:
                sql_lines.append(f"{set_line} ")
                
        if where_clause:
            sql_lines.append(where_clause)
        else:
            sql_lines[-1] = sql_lines[-1].rstrip() + "; "
            
        body += _sql(sql_lines)
        body += f'    db.Execute SQL\n'
        body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: {fields_str}"\n\n'

    # ── PASS 3: Generate consolidated post-update logic ───────────────────
    if post_update_branches:
        body += "    ' Cleanup post-update data\n"
        for b in post_update_branches:
            if not b["outputs"]:
                continue
            
            # Determine the condition
            cond_text = ""
            for cr in b["conditions"]:
                c_val = cr.get("条件 / 項目マッピング", "") or cr.get("条件", "")
                if c_val and "完了後" not in c_val:
                    cond_text = c_val
                    break
            
            if not cond_text:
                continue
                
            # Build WHERE condition for target table
            where_expr = f"T.ID LIKE '{prefix}_%'"
            if "個人" in cond_text:
                where_expr += " AND T.[kind_id] = '個人'"
            elif "法人" in cond_text:
                where_expr += " AND T.[kind_id] = '法人'"
            else:
                where_expr += " AND True" # fallback if we can't parse
                
            set_parts = []
            for out_row in b["outputs"]:
                f = out_row.get("項目名") or out_row.get("DB_Field")
                if not f:
                    continue
                out_field = out_row.get("該当項目名_1", "")
                if out_row.get("処理") == "固定値出力":
                    val_clean = out_field.strip('\'"')
                    out_expr = f"'{val_clean}'"
                elif out_row.get("処理") == "削除":
                    out_expr = "NULL"
                else:
                    out_expr = f"T.[{out_field}]" if out_field else "NULL"
                set_parts.append(f"    T.[{f}] = {out_expr}")
                
            if not set_parts:
                continue
                
            sql_lines = [f"UPDATE {_b(target_table)} AS T SET "]
            sql_lines.append(",\n".join(set_parts))
            sql_lines.append(f"WHERE {where_expr}; ")
            
            body += _sql(sql_lines)
            body += f'    db.Execute SQL\n'
            
        body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: post-update cleanup"\n\n'

    # ── PASS 4: Generate legacy_id duplicate deletion ─────────────────────
    if "legacy_id" in target_fields_all:
        dedup_sql = [
            f"DELETE FROM {_b(target_table)} ",
        ]
        if target_table == "Account":
            dedup_sql.append(f"WHERE ID LIKE '{prefix}_%' ")
            dedup_sql.append(f"AND ID NOT IN ( ")
            dedup_sql.append(f"   SELECT MIN(ID) ")
            dedup_sql.append(f"   FROM {_b(target_table)} ")
            dedup_sql.append(f"   WHERE ID LIKE '{prefix}_%' ")
            dedup_sql.append(f"   GROUP BY legacy_id ")
            dedup_sql.append(f") ")
        else:
            dedup_sql.append(f"WHERE ID NOT IN ( ")
            dedup_sql.append(f"   SELECT MIN(ID) ")
            dedup_sql.append(f"   FROM {_b(target_table)} ")
            dedup_sql.append(f"   GROUP BY legacy_id ")
            dedup_sql.append(f") ")
            
        body += "\n    ' duplicate legacy_id\n"
        body += _sql(dedup_sql)
        body += f"    db.Execute SQL\n"
        body += f'    log_write "{sub_name}:{sheet_name} → delete duplicate"\n\n'

    return dim_con_code + body


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
