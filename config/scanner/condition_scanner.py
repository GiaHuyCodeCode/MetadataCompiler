#!/usr/bin/env python3
"""
condition_scanner.py — Scan condition columns in sheet_raw.json files,
aggregate their counts, and update sample/semantic_dictionary.json.
"""

import os
import sys
import json
import glob
from datetime import datetime

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

# 30 Missing keywords mappings based on sample/SOURCE.txt and compiler rules
MISSING_MAPPINGS = {
    "紐づく場合": {
        "role": "CONDITION_MARKER",
        "ast_type": "JOIN_FILTER",
        "notes": "Record khớp JOIN trong 出力条件 (thông thường FLG='1' để loại trừ hoặc FLG='0' cho logic ngược). Pattern: INNER JOIN ON T1.key = T2.key.",
        "vba_pattern": "UPDATE T1 INNER JOIN {join_table} AS T2 ON T1.[key]=T2.[key] SET T1.FLG='1'"
    },
    '"会社"を含む場合': {
        "role": "CONDITION_MARKER",
        "ast_type": "LIKE_PATTERN",
        "notes": "Marker điều kiện LIKE '%会社%'. Tìm kiếm chuỗi có chứa '会社' trong cột được chỉ định.",
        "sql_op": "T1.[field] LIKE '%会社%'"
    },
    '"会社"を含まない場合': {
        "role": "CONDITION_MARKER",
        "ast_type": "NOT_LIKE_PATTERN",
        "notes": "Marker điều kiện NOT LIKE '%会社%'. Loại trừ chuỗi chứa '会社' trong cột được chỉ định.",
        "sql_op": "T1.[field] NOT LIKE '%会社%'"
    },
    "データ順で１番最初のデータを出力対象とする": {
        "role": "CONDITION_MARKER",
        "ast_type": "DEDUP",
        "notes": "Chọn record đầu tiên theo thứ tự xuất hiện khi dedup. Thường dùng MIN(ID) và GROUP BY key.",
        "vba_pattern": "UPDATE T1 SET FLG='1' WHERE ID NOT IN (SELECT MIN(ID) FROM T1 WHERE FLG='0' GROUP BY key)"
    },
    '上記以外で"契約中(他社)"の場合': {
        "role": "CONDITION_MARKER",
        "ast_type": "BRANCH_CONDITION",
        "notes": "Marker điều kiện nhánh ELSE IF cho 契約中(他社) trong SWITCH hoặc UPDATE block.",
        "sql_op": "T1.[status] = '契約中(他社)'"
    },
    '"サブリース"': {
        "role": "CONSTANT_VALUE",
        "ast_type": "CONSTANT_VALUE",
        "notes": "Hằng số chuỗi 'サブリース' dùng để so sánh."
    },
    '"管理委託"': {
        "role": "CONSTANT_VALUE",
        "ast_type": "CONSTANT_VALUE",
        "notes": "Hằng số chuỗi '管理委託' dùng để so sánh."
    },
    '"個人"': {
        "role": "CONSTANT_VALUE",
        "ast_type": "CONSTANT_VALUE",
        "notes": "Hằng số chuỗi '個人' dùng để so sánh."
    },
    "データ順で１番最初のデータ1件のみを採用": {
        "role": "CONDITION_MARKER",
        "ast_type": "DEDUP",
        "notes": "Chọn record đầu tiên theo thứ tự xuất hiện khi dedup. Thường dùng MIN(ID) và GROUP BY key.",
        "vba_pattern": "UPDATE T1 SET FLG='1' WHERE ID NOT IN (SELECT MIN(ID) FROM T1 WHERE FLG='0' GROUP BY key)"
    },
    "条件分岐に当てはまる": {
        "role": "CONDITION_MARKER",
        "ast_type": "BRANCH_CONDITION",
        "notes": "Nhánh điều kiện thỏa mãn điều kiện phân nhánh."
    },
    "条件分岐に当てはまらない": {
        "role": "CONDITION_MARKER",
        "ast_type": "BRANCH_CONDITION",
        "notes": "Nhánh điều kiện không thỏa mãn điều kiện phân nhánh."
    },
    '"P"もしくは"駐"が含まれる場合': {
        "role": "CONDITION_MARKER",
        "ast_type": "LIKE_PATTERN",
        "notes": "Marker điều kiện LIKE '%P%' OR LIKE '%駐%'. Dùng để nhận diện hầm/bãi đỗ xe hoặc bến đỗ.",
        "sql_op": "T1.[field] LIKE '%P%' OR T1.[field] LIKE '%駐%'"
    },
    '"契約中"または"契約中(他社)"または"解約予定"の場合': {
        "role": "CONDITION_MARKER",
        "ast_type": "IN_LIST_FILTER",
        "notes": "Lọc các trạng thái 契約中, 契約中(他社), 解約予定.",
        "sql_op": "T1.[field] IN ('契約中', '契約中(他社)', '解約予定')"
    },
    "出力されたkind_idの値が'個人'の場合": {
        "role": "CONDITION_MARKER",
        "ast_type": "BRANCH_CONDITION",
        "notes": "Kiểm tra kind_id sau khi xuất là '個人' để gán hoặc điều chỉnh các cột liên quan.",
        "sql_op": "T.[kind_id] = '個人'"
    },
    "出力されたkind_idの値が'法人'の場合": {
        "role": "CONDITION_MARKER",
        "ast_type": "BRANCH_CONDITION",
        "notes": "Kiểm tra kind_id sau khi xuất là '法人' để gán hoặc điều chỉnh các cột liên quan.",
        "sql_op": "T.[kind_id] = '法人'"
    },
    '"契約中"または"契約中(他社)" または"解約予定"の場合': {
        "role": "CONDITION_MARKER",
        "ast_type": "IN_LIST_FILTER",
        "notes": "Lọc các trạng thái 契約中, 契約中(他社), 解約予定 (có khoảng trắng thừa).",
        "sql_op": "T1.[field] IN ('契約中', '契約中(他社)', '解約予定')"
    },
    '紐づく場合かつ管理形態が"一括借上"の場合': {
        "role": "CONDITION_MARKER",
        "ast_type": "BRANCH_CONDITION",
        "notes": "Điều kiện kết hợp: 紐づく AND 管理形態='一括借上'.",
        "sql_op": "T2.[key] IS NOT NULL AND T1.[管理形態] = '一括借上'"
    },
    '"legacy_contractor_id"がブランクではない場合': {
        "role": "CONDITION_MARKER",
        "ast_type": "BRANCH_CONDITION",
        "notes": "Kiểm tra legacy_contractor_id không rỗng.",
        "sql_op": "Nz(T.[legacy_contractor_id], '') <> ''"
    },
    '"legacy_contractor_id"がブランクの場合': {
        "role": "CONDITION_MARKER",
        "ast_type": "BRANCH_CONDITION",
        "notes": "Kiểm tra legacy_contractor_id rỗng.",
        "sql_op": "Nz(T.[legacy_contractor_id], '') = ''"
    },
    '"電気\u3000家主一括契約"または"水道\u3000家主一括契約"を含まない場合': {
        "role": "CONDITION_MARKER",
        "ast_type": "NOT_LIKE_PATTERN",
        "notes": "Kiểm tra không chứa '電気　家主一括契約' và '水道　家主一括契約'.",
        "sql_op": "T1.[field] NOT LIKE '%電気　家主一括契約%' AND T1.[field] NOT LIKE '%水道　家主一括契約%'"
    },
    '"水道\u3000家主一括契約"を含まない場合': {
        "role": "CONDITION_MARKER",
        "ast_type": "NOT_LIKE_PATTERN",
        "notes": "Kiểm tra không chứa '水道　家主一括契約'.",
        "sql_op": "T1.[field] NOT LIKE '%水道　家主一括契約%'"
    },
    "紐づく場合、結果①を出力": {
        "role": "CONDITION_MARKER",
        "ast_type": "BRANCH_CONDITION",
        "notes": "Khi khớp kết nối, xuất kết quả thứ nhất."
    },
    "紐づかなかった場合": {
        "role": "CONDITION_MARKER",
        "ast_type": "LEFT_JOIN_SET_ZERO",
        "notes": "Khi không khớp kết nối (đồng nghĩa với 紐づかない場合).",
        "vba_pattern": "UPDATE T1 LEFT JOIN {join_table} AS T2 ON T1.[key]=T2.[key] SET T1.FLG='0' WHERE T2.[key] IS NULL"
    },
    "データがある場合": {
        "role": "CONDITION_MARKER",
        "ast_type": "BRANCH_CONDITION",
        "notes": "Khi tồn tại dữ liệu (không rỗng hoặc có bản ghi khớp)."
    },
    '"契約中(他社)"の場合': {
        "role": "CONDITION_MARKER",
        "ast_type": "BRANCH_CONDITION",
        "notes": "Nhánh điều kiện cho 契約中(他社).",
        "sql_op": "T1.[status] = '契約中(他社)'"
    },
    "上記の出力完了後": {
        "role": "CONDITION_MARKER",
        "ast_type": "BRANCH_CONDITION",
        "notes": "Thực hiện sau khi các lệnh trên hoàn tất."
    },
    '"加入"': {
        "role": "CONSTANT_VALUE",
        "ast_type": "CONSTANT_VALUE",
        "notes": "Hằng số chuỗi '加入' dùng để so sánh."
    },
    "連絡不要'が含まれていないまたは\n\"田中総合燃料\"が含まれていない場合": {
        "role": "CONDITION_MARKER",
        "ast_type": "NOT_LIKE_PATTERN",
        "notes": "Kiểm tra không chứa '連絡不要' hoặc '田中総合燃料'.",
        "sql_op": "T1.[field] NOT LIKE '%連絡不要%' OR T1.[field] NOT LIKE '%田中総合燃料%'"
    },
    "連絡不要'が含まれていない場合または\n\"汲み取り業者\"が含まれていない場合": {
        "role": "CONDITION_MARKER",
        "ast_type": "NOT_LIKE_PATTERN",
        "notes": "Kiểm tra không chứa '連絡不要' hoặc '汲み取り業者'.",
        "sql_op": "T1.[field] NOT LIKE '%連絡不要%' OR T1.[field] NOT LIKE '%汲み取り業者%'"
    },
    '"契約者No"≠"legacy_resident_id"': {
        "role": "CONDITION_MARKER",
        "ast_type": "BRANCH_CONDITION",
        "notes": "Điều kiện so sánh khác biệt: 契約者No <> legacy_resident_id.",
        "sql_op": "T1.[契約者No] <> T.[legacy_resident_id]"
    }
}


def scan_and_update():
    # 1. Scan all sheet_raw.json files
    search_path = os.path.join(PROJECT_ROOT, "output", "**", "sheet_raw.json")
    json_files = glob.glob(search_path, recursive=True)
    print(f"🔍 Found {len(json_files)} sheet_raw.json files.")

    cond_stats = {}
    for jf in json_files:
        try:
            with open(jf, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            print(f"⚠️ Error reading {jf}: {e}")
            continue

        sheet_name = data.get("_meta", {}).get("source_sheet", "")
        sections = data.get("sections", {})
        for sec_name, sec_data in sections.items():
            if not isinstance(sec_data, dict):
                continue
            sentences = sec_data.get("sentences", [])
            for sent in sentences:
                for row in sent.get("rows", []):
                    for k, v in row.items():
                        if "条件" in k and isinstance(v, str):
                            # Normalize value: strip spaces, keep internal newlines consistent
                            v_norm = v.strip()
                            # Replace \r\n with \n
                            v_norm = v_norm.replace("\r\n", "\n")
                            if v_norm not in cond_stats:
                                cond_stats[v_norm] = {
                                    "count": 0,
                                    "sheets": set(),
                                    "keys": set()
                                }
                            cond_stats[v_norm]["count"] += 1
                            cond_stats[v_norm]["sheets"].add(sheet_name)
                            cond_stats[v_norm]["keys"].add(k)

    # 2. Load semantic dictionary
    dict_path = os.path.join(PROJECT_ROOT, "sample", "semantic_dictionary.json")
    if not os.path.exists(dict_path):
        print(f"❌ Semantic dictionary not found at {dict_path}")
        sys.exit(1)

    with open(dict_path, "r", encoding="utf-8") as f:
        dict_data = json.load(f)

    entries = dict_data.get("entries", {})
    now_str = datetime.now().isoformat()

    added_count = 0
    updated_count = 0

    # 3. Update existing and insert new entries
    for val, stats in sorted(cond_stats.items(), key=lambda x: x[1]["count"], reverse=True):
        # Determine if we have a match in the dict or if it's new
        # We also check for normalized comparison in case of minor quote/space differences
        matched_key = None
        if val in entries:
            matched_key = val
        else:
            # Check if there is a key with different quotes or newlines
            for k in entries.keys():
                if k.strip().replace("\r\n", "\n") == val:
                    matched_key = k
                    break

        sheets_list = sorted(list(stats["sheets"]))
        keys_list = sorted(list(stats["keys"]))

        if matched_key:
            # Update existing entry
            entry = entries[matched_key]
            # Union of source sheets and keys
            existing_sheets = set(entry.get("source_sheets", []))
            existing_keys = set(entry.get("found_in_keys", []))
            
            entry["count"] = stats["count"]
            entry["source_sheets"] = sorted(list(existing_sheets.union(stats["sheets"])))
            entry["found_in_keys"] = sorted(list(existing_keys.union(stats["keys"])))
            entry["classified_at"] = now_str
            updated_count += 1
        else:
            # Create new entry
            # Check if we have pre-defined metadata for this missing keyword
            mapping = MISSING_MAPPINGS.get(val)
            if not mapping:
                # Fallback mapping if not in pre-defined list
                mapping = {
                    "role": "CONDITION_MARKER",
                    "ast_type": "BRANCH_CONDITION",
                    "notes": "Marker điều kiện phân nhánh"
                }

            entry = {
                "role": mapping.get("role", "CONDITION_MARKER"),
                "ast_type": mapping.get("ast_type", "BRANCH_CONDITION"),
                "count": stats["count"],
                "source_sheets": sheets_list,
                "found_in_keys": keys_list,
                "classified_by": "AUTO",
                "classified_at": now_str,
                "notes": mapping.get("notes", "")
            }
            if "sql_op" in mapping:
                entry["sql_op"] = mapping["sql_op"]
            if "vba_pattern" in mapping:
                entry["vba_pattern"] = mapping["vba_pattern"]

            entries[val] = entry
            added_count += 1

    # Update metadata
    if "_meta" not in dict_data:
        dict_data["_meta"] = {}
    dict_data["_meta"]["last_updated"] = now_str

    # 4. Save updated semantic dictionary
    with open(dict_path, "w", encoding="utf-8") as f:
        json.dump(dict_data, f, ensure_ascii=False, indent=2)

    print("\n═" * 50)
    print("  SCAN & UPDATE COMPLETE")
    print("═" * 50)
    print(f"Total scanned unique conditions: {len(cond_stats)}")
    print(f"Updated existing entries:       {updated_count}")
    print(f"Added new entries:              {added_count}")
    print("═" * 50)


if __name__ == "__main__":
    scan_and_update()
