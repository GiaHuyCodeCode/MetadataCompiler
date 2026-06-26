import os
import json
import re
from pathlib import Path

COMPILER_DIR = Path(__file__).parent
DICT_PATH = COMPILER_DIR.parent.parent / "sample" / "semantic_dictionary.json"

def load_dict():
    with open(DICT_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def save_dict(data):
    with open(DICT_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def extract_snippet_from_bas(bas_content, field_name):
    # Split by log_write
    blocks = re.split(r'log_write\s*"([^"]*)"', bas_content)
    # blocks[0] is code before first log_write
    # blocks[1] is the string inside first log_write
    # blocks[2] is code after first log_write, etc.
    
    for i in range(1, len(blocks), 2):
        log_str = blocks[i]
        if "特殊処理:" in log_str and re.search(r'\b' + re.escape(field_name) + r'\b', log_str):
            # This is our target log_write. The code is in blocks[i-1].
            code_block = blocks[i-1]
            # We want the code starting from the last SQL = "" or Dim
            # Actually, just taking the whole blocks[i-1] after the previous log_write is fine.
            code_block = code_block.strip()
            # If it contains "SQL =", find the first one
            idx = code_block.find('SQL = ""')
            if idx == -1:
                idx = code_block.find("Dim ")
            
            if idx != -1:
                code_block = code_block[idx:].strip()
                
            return code_block + f'\n    log_write "{log_str}"'
            
    return None

def generalize_snippet(vba_code, context):
    target_table = context.get("target_table", "")
    source_table = context.get("source_table", "")
    prefix = context.get("prefix", "")
    sub_name = context.get("sub_name", "")
    sheet_name = context.get("sheet_name", "")

    if target_table:
        vba_code = vba_code.replace(target_table, "{target_table}")
    if source_table:
        vba_code = vba_code.replace(source_table, "{source_table}")
    if prefix:
        vba_code = vba_code.replace(f"'{prefix}_'", "'{prefix}_'")
        vba_code = vba_code.replace(f"'{prefix}_%'", "'{prefix}_%'")
        vba_code = vba_code.replace(f"('{prefix}_'", "('{prefix}_'")

    if sub_name and sheet_name:
        vba_code = vba_code.replace(f'{sub_name}:{sheet_name}', '{sub_name}:{sheet_name}')
        
    return vba_code

def extract_keywords_from_rows(rows):
    keywords = set()
    for r in rows:
        keywords.add(r.get("項目名", ""))
        cond = r.get("条件 / 項目マッピング", "") or r.get("条件", "") or r.get("該当項目名_1", "")
        if cond:
            keywords.add(cond)
        if r.get("処理"):
            keywords.add(r.get("処理"))
    
    keywords = [k for k in keywords if k and len(k) > 1]
    return list(keywords)

def run_learner(test_dir):
    base_path = Path(test_dir)
    if not base_path.exists():
        print(f"Directory {test_dir} not found.")
        return
        
    dict_data = load_dict()
    # Reset for clean testing
    dict_data.get("compiler_mappings", {})["custom_snippet_mappings"] = []
    custom_snippets = dict_data.get("compiler_mappings", {})["custom_snippet_mappings"]
    
    existing_triggers = []
    for snippet in custom_snippets:
        existing_triggers.append(snippet.get("match_conditions", []))
        
    import sys
    sys.path.insert(0, str(COMPILER_DIR.parent.parent))
    from config.compiler.vba_compiler import _primary_source_table, _infer_target, _infer_prefix
    
    count_learned = 0
    
    for root, dirs, files in os.walk(base_path):
        if "sheet_raw.json" in files:
            json_path = Path(root) / "sheet_raw.json"
            bas_files = [f for f in files if f.endswith(".bas")]
            if not bas_files:
                continue
                
            bas_path = Path(root) / bas_files[0]
            
            with open(bas_path, "r", encoding="utf-8") as f:
                bas_content = f.read()
                
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    sheet_data = json.load(f)
            except Exception as e:
                print(f"Error parsing {json_path}: {e}")
                continue
                
            meta = sheet_data.get("_meta", {})
            sheet_name = meta.get("source_sheet", "Sheet")
            sections   = sheet_data.get("sections", {})
            filenames    = sections.get("ファイル名", [])
            special_proc = sections.get("特殊処理", {})
            
            source_table          = _primary_source_table(filenames)
            target_table, sub_name = _infer_target(sheet_name)
            prefix                = _infer_prefix(sheet_name)
            
            sheet_context = {
                "target_table": target_table,
                "source_table": source_table,
                "prefix": prefix,
                "sheet_name": sheet_name,
                "sub_name": sub_name
            }
            
            sp_sentences = special_proc.get("sentences", [])
                
            for sentence in sp_sentences:
                rows = sentence.get("rows", [])
                if not rows:
                    continue
                    
                field = rows[0].get("項目名", "")
                if not field:
                    continue
                    
                vba_code = extract_snippet_from_bas(bas_content, field)
                if not vba_code:
                    continue
                    
                keywords = extract_keywords_from_rows(rows)
                
                already_known = False
                for t in existing_triggers:
                    if set(t) == set(keywords):
                        already_known = True
                        break
                        
                if already_known:
                    continue
                    
                gen_code = generalize_snippet(vba_code, sheet_context)
                
                custom_snippets.append({
                    "description": f"Learned from {sheet_name} -> {field}",
                    "match_conditions": keywords,
                    "template": gen_code
                })
                count_learned += 1
                existing_triggers.append(keywords)
                print(f"Learned snippet for field '{field}' with keywords {keywords}")
                
    if count_learned > 0:
        save_dict(dict_data)
        print(f"Successfully learned {count_learned} new snippets and saved to semantic_dictionary.json")
    else:
        print("No new snippets learned.")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        run_learner(sys.argv[1])
