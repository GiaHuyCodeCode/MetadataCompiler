import os

compiler_path = "/home/huyhg/Documents/MetadataCompiler_V1/compiler/vba_compiler.py"

with open(compiler_path, "r", encoding="utf-8") as f:
    content = f.read()

# Locate the beginning of _gen_normal_processing
pos = content.find("def _gen_normal_processing")
if pos == -1:
    print("Error: Could not find _gen_normal_processing")
    exit(1)

header = content[:pos]

new_code = """def _gen_normal_processing(rows: list, source_table: str, target_table: str,
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
            v = val1.strip("'\\" ")
            fields.append(f"    '{v}' as {field}, ")
        elif shori == "そのまま出力" and has_hyphen_null:
            # 項目マッピング says: if value is "-", output NULL
            fields.append(f"    IIF(T.[{val1}] = '-', NULL, T.[{val1}]) as {field}, ")
        elif shori == "そのまま出力":
            fields.append(f"    T.[{val1}] as {field}, ")
        elif shori == "yyyy-mm-dd形式":
            fields.append(f"    IIF(IsDate(T.[{val1}]), Format(T.[{val1}], 'yyyy-mm-dd'), NULL) as {field}, ")
        elif shori == "yyyymm形式":
            fields.append(f"    Format(T.[{val1}], 'yyyymm') as {field}, ")
        elif shori == "ハイフン付出力":
            parts = [f"T.[{v}]" for v in [val1, val2, val3] if v]
            fields.append(f"    {' & \\"-\\" & '.join(parts)} as {field}, ")
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

    code = "    '通常処理\\n" + _sql(field_lines)
    code += f'    db.Execute SQL\\n'
    code += f'    log_write "{sub_name}:{sheet_name} → 通常処理"\\n\\n'
    return code


# ─────────────────────────────────────────────────────────
# 特殊処理 generator
# ─────────────────────────────────────────────────────────

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

    # ── Helper: sắp xếp 契約者N theo thứ tự ưu tiên 1→3→2 ──────────────
    def _con_priority(r: dict) -> int:
        f = r.get("該当項目名_1", "")
        if "契約者1" in f: return 0
        if "契約者3" in f: return 1
        if "契約者2" in f: return 2
        return 9

    def _build_sql_condition_generic(cond_row: dict, con_idx_map: dict[str, str]) -> str:
        cond_str = cond_row.get("条件 / 項目マッピング", "") or cond_row.get("条件", "")
        cond_type = _detect_condition_type(cond_str)
        if cond_type == "ELSE":
            return "True"
        
        # Collect all fields
        fields = []
        for key in ("該当項目名_1", "該当項目名_2", "該当項目名_3"):
            val = cond_row.get(key)
            if val:
                fields.append(val)
                
        # If the first field has a con index, use that
        if fields and fields[0] in con_idx_map:
            idx = con_idx_map[fields[0]]
            return f'" & con{idx} & "'
            
        val_parsed = _parse_condition_value(cond_str)
        # Clean up prefixes
        val_parsed = re.sub(r'^(すべて|いずれも|上記以外で|上記以外かつ|上記以外)', '', val_parsed).strip()
        val_parsed = val_parsed.strip('"\\'')
        
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
    # Chỉ cần 1 bộ sorted_conds (các sentences CONTRACTOR_PRIORITY dùng chung)
    shared_sorted_conds: list = []
    for sentence in sentences:
        rows = sentence.get("rows", [])
        cond_rows = [
            r for r in rows
            if r.get("処理") == "条件分岐" and "入居" in r.get("該当項目名_1", "")
        ]
        if len(cond_rows) >= 2:
            candidate = sorted(cond_rows, key=_con_priority)
            if len(candidate) > len(shared_sorted_conds):
                shared_sorted_conds = candidate

    # ── Build Dim con declarations (1 lần duy nhất) ───────────────────────
    dim_con_code = ""
    con_idx_map: dict[str, str] = {}   # field_name → con_idx ("1\\",\\"3\\",\\"2\\",...)

    if shared_sorted_conds:
        dim_con_code += "    ' Dim con — điều kiện ưu tiên 契約者1→3→2 (khai báo 1 lần)\\n"
        for i, cr in enumerate(shared_sorted_conds):
            vf  = cr.get("該当項目名_1", "")
            idx = ["1", "3", "2"][i] if i < 3 else str(i + 1)
            con_idx_map[vf] = idx
            
            # Extract condition value dynamically
            cond_val = _parse_condition_value(cr.get("条件 / 項目マッピング", "") or cr.get("条件", "")) or "入居有り"
            
            prev = " AND ".join(
                f"T1.[{shared_sorted_conds[j].get('該当項目名_1', '')}] <> '{cond_val}'"
                for j in range(i)
            )
            cur  = f"T1.[{vf}] = '{cond_val}'"
            full = f"({(prev + ' AND ') if prev else ''}{cur})"
            dim_con_code += f"    Dim con{idx} As String\\n"
            dim_con_code += f'    con{idx} = "{full}"\\n'
        dim_con_code += "\\n"

    # ── PASS 2: Sinh code cho từng sentence ───────────────────────────────
    body = ""

    for sentence in sentences:
        rows = sentence.get("rows", [])
        if not rows:
            continue

        contractor_cond_rows = [
            r for r in rows
            if r.get("処理") == "条件分岐" and "入居" in r.get("該当項目名_1", "")
        ]
        output_rows = [
            r for r in rows
            if r.get("処理") == "出力" or (r.get("処理", "") == "" and r.get("該当項目名_1", ""))
        ]
        has_relation_link = any(r.get("処理") == "紐づけ" for r in rows)

        if len(contractor_cond_rows) >= 2 and output_rows and not has_relation_link:
            src_tbl = _clean_table_name(
                next((r.get("該当ファイル名", source_table) for r in rows if r.get("該当ファイル名")), source_table)
            )
            
            # Find all target fields in the sentence
            target_fields = []
            for r in rows:
                f = r.get("項目名") or r.get("DB_Field")
                if f and f not in target_fields:
                    target_fields.append(f)
                    
            for f in target_fields:
                # Group rows into condition-output pairs for this target field
                pairs_for_f = []
                i = 0
                while i < len(rows):
                    r = rows[i]
                    is_cond = r.get("処理") == "条件分岐" or any(k in r.get("条件 / 項目マッピング", "") for k in ("の場合", "すべて", "いずれも", "上記以外", "含まれて"))
                    if is_cond:
                        out_r = None
                        for j in range(i + 1, len(rows)):
                            next_r = rows[j]
                            is_next_cond = next_r.get("処理") == "条件分岐" or any(k in next_r.get("条件 / 項目マッピング", "") for k in ("の場合", "すべて", "いずれも", "上記以外", "含まれて"))
                            if is_next_cond:
                                break
                            if next_r.get("項目名") == f or next_r.get("DB_Field") == f:
                                out_r = next_r
                                break
                        if out_r:
                            pairs_for_f.append((r, out_r))
                    i += 1
                    
                if not pairs_for_f:
                    continue
                    
                switch_lines = [
                    f"UPDATE {target_table} AS T ",
                    f"INNER JOIN {src_tbl} AS T1 ON T.ID = ('{prefix}_' & T1.ID) ",
                    f"SET T.[{f}] = SWITCH( "
                ]
                
                for cond_row, out_row in pairs_for_f:
                    cond_expr = _build_sql_condition_generic(cond_row, con_idx_map)
                    
                    # Build output expression
                    out_field = out_row.get("該当項目名_1", "")
                    if out_row.get("処理") == "固定値出力":
                        out_expr = f"'{out_field.strip(\\'\\"\\')}'"
                    else:
                        out_expr = f"T1.[{out_field}]"
                        
                    # Check for hyphen to null mapping
                    if _is_hyphen_null_mapping(out_row):
                        out_expr = f"IIF({out_expr} = '-', NULL, {out_expr})"
                        
                    switch_lines.append(f'    {cond_expr}, {out_expr}, ')
                    
                # Replace the last comma with closing parenthesis
                if len(switch_lines) > 3:
                    switch_lines[-1] = switch_lines[-1].rstrip(", ") + " ); "
                    
                # Render switch lines
                sql_code = ""
                for line in switch_lines:
                    sql_code += f'    SQL = SQL & "{line}"\\n'
                    
                body += sql_code
                body += f'    db.Execute SQL\\n'
                body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: {f}"\\n\\n'
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

        switch_conds = [c for c in conditions if c["shori"] == "条件分岐"]
        fixed_vals   = [c for c in conditions if c["shori"] == "固定値出力"]
        direct_vals  = [c for c in conditions if c["shori"] == "そのまま出力"]
        null_vals    = [c for c in conditions if c["shori"] == "出力なし"]

        # ── Check for "-"の場合 + 出力なし pair (HYPHEN_TO_NULL in 特殊処理) ──
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
                f"SET T.[{field}] = IIF(T1.[{src_f}] = '-', NULL, T1.[{src_f}]) ",
            ]
            body += _sql(sql_lines)
            body += f'    db.Execute SQL\\n'
            body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: {field}"\\n\\n'
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
                        switch_lines.append(f"    ({cond_expr}), IIF(T1.[{result_v}] = '-', NULL, T1.[{result_v}]), ")
                    else:
                        switch_lines.append(f"    ({cond_expr}), T1.[{result_v}], ")

            switch_lines.append("    True, NULL ); ")
            body += _sql(switch_lines)
            body += f'    db.Execute SQL\\n'
            body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: {field}"\\n\\n'

        elif switch_conds and len(switch_conds) == 1:
            sc = switch_conds[0]
            where = _build_sql_condition_generic(sc, {})

            result_rows = [c for c in conditions if c["shori"] == "そのまま出力"]
            result_field = result_rows[0]["src_field"] if result_rows else sc["src_field"]

            if result_field in hyphen_null_fields:
                set_clause = f"T.[{field}] = IIF(T1.[{result_field}] = '-', NULL, T1.[{result_field}])"
            else:
                set_clause = f"T.[{field}] = T1.[{result_field}]"

            sql_lines = [
                f"UPDATE {target_table} AS T ",
                f"INNER JOIN {source_table} AS T1 ON T.ID = ('{prefix}_' & T1.ID) ",
                f"SET {set_clause} ",
                f"WHERE {where}; ",
            ]
            body += _sql(sql_lines)
            body += f'    db.Execute SQL\\n'
            body += f'    log_write "{sub_name}:{sheet_name} → 特殊処理: {field}"\\n\\n'

        else:
            from compiler.ai_fallback import ai_fallback_generate
            body += ai_fallback_generate(sentence, sheet_context)

    return "    '特殊処理\\n\\n" + dim_con_code + body


# ─────────────────────────────────────────────────────────
# Main compile function
# ─────────────────────────────────────────────────────────

def compile_sheet(sheet_raw: dict) -> str:
    \"\"\"Sinh mã VBA từ một sheet_raw dict.\"\"\"
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
    code += "\\n    End If\\n"
    code += _gen_footer(target_table, sub_name)

    return code



def compile_from_file(json_path: str) -> str:
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)
    return compile_sheet(data)
"""

with open(compiler_path, "w", encoding="utf-8") as f:
    f.write(header + new_code)

print("Rewrote vba_compiler.py successfully!")
