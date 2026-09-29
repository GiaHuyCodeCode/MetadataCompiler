"""
main.py — MetadataCompiler V1  CLI Entry Point
=================================================
Usage:
    python3 main.py scan    --test2          # Scan xlsx → sheet_raw.json
    python3 main.py compile --test2          # Compile sheet_raw.json → .bas
    python3 main.py run     --test2          # scan + compile (all-in-one)
    python3 main.py learn   --test1          # Học snippet tự động từ file .bas

Options:
    --folder <name>   hoặc  --<name>   ví dụ: --test2
    --file   <name>   chỉ xử lý một file xlsx cụ thể (optional)
"""

import sys
import os
import json
import argparse
from datetime import datetime

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT_ROOT)
sys.path.insert(0, os.path.join(PROJECT_ROOT, "config"))  # resolve compiler/ → config/compiler/

from scanner.xlsx_scanner import scan_xlsx
from compiler.vba_compiler import compile_sheet as compile_sheet_static
from compiler.agent_compiler import compile_sheet_agent


# ─────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────

def _sanitize(name: str) -> str:
    invalid = r'/\:*?"<>|'
    result = name.strip()
    for ch in invalid:
        result = result.replace(ch, "_")
    return result


def _banner(text: str):
    print(f"\n{'═'*60}")
    print(f"  {text}")
    print(f"{'═'*60}")


def _sep(w=60):
    return "─" * w


def _parse_folder_arg(args: list[str]) -> str:
    for a in args:
        if a.startswith("--"):
            return a[2:]
    return ""


# ─────────────────────────────────────────────────────────
# SCAN command
# ─────────────────────────────────────────────────────────

def cmd_scan(folder_name: str, only_file: str = "") -> list[str]:
    """
    Scan tất cả xlsx trong input/<folder_name>/ → output/<folder_name>/.
    Returns list of exported json paths.
    """
    input_dir  = os.path.join(PROJECT_ROOT, "input",  folder_name)
    output_dir = os.path.join(PROJECT_ROOT, "output", folder_name)

    if not os.path.isdir(input_dir):
        print(f"❌  Không tìm thấy: {input_dir}")
        sys.exit(1)

    xlsx_files = sorted([
        f for f in os.listdir(input_dir)
        if f.lower().endswith(".xlsx") and not f.startswith("~")
        and (not only_file or f == only_file)
    ])

    if not xlsx_files:
        print(f"⚠️  Không có file xlsx trong: {input_dir}")
        return []

    _banner(f"SCAN  →  input/{folder_name}")
    print(f"  📊  {len(xlsx_files)} file xlsx\n")

    exported_json: list[str] = []

    for xlsx_filename in xlsx_files:
        file_path = os.path.join(input_dir, xlsx_filename)
        file_stem = os.path.splitext(xlsx_filename)[0]

        print(f"{_sep()}")
        print(f"📄  {xlsx_filename}")

        try:
            result = scan_xlsx(file_path)
        except Exception as e:
            print(f"  ❌  Lỗi: {e}")
            continue

        for sheet_name in result["valid_sheets"]:
            safe  = _sanitize(sheet_name)
            out_dir = os.path.join(output_dir, file_stem, safe)
            os.makedirs(out_dir, exist_ok=True)
            out_path = os.path.join(out_dir, "sheet_raw.json")

            payload = {
                "_meta": {
                    "generated_at": datetime.now().isoformat(),
                    "source_file":  xlsx_filename,
                    "source_sheet": sheet_name,
                    "folder":       folder_name,
                },
                "sections": result["data"][sheet_name],
            }
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)

            rel = os.path.relpath(out_path, PROJECT_ROOT)
            print(f"  ✔  [{sheet_name}]  →  {rel}")
            exported_json.append(out_path)

        for s in result["skipped_sheets"]:
            print(f"  ⏭   skip: {s}")

    _banner(f"SCAN DONE — {len(exported_json)} sheet")
    return exported_json


# ─────────────────────────────────────────────────────────
# COMPILE command
# ─────────────────────────────────────────────────────────

def cmd_compile(folder_name: str, json_paths: list[str] = None, only_file: str = "", use_agent: bool = False) -> list[str]:
    """
    Compile sheet_raw.json → .bas.
    Nếu json_paths=None thì quét toàn bộ output/<folder_name>/.
    Returns list of written .bas paths.
    """
    output_dir = os.path.join(PROJECT_ROOT, "output", folder_name)
    
    compiler_func = compile_sheet_agent if use_agent else compile_sheet_static

    if json_paths is None:
        # Tìm tất cả sheet_raw.json trong folder
        json_paths = []
        for root, _, files in os.walk(output_dir):
            for fn in files:
                if fn == "sheet_raw.json":
                    json_paths.append(os.path.join(root, fn))
                    
    if only_file:
        # Lọc danh sách json_paths nếu có cờ --file
        # (có thể là tên thư mục chứa sheet, vd: 契約_K)
        json_paths = [p for p in json_paths if only_file in p]
        
    json_paths.sort()

    if not json_paths:
        print(f"⚠️  Không tìm thấy sheet_raw.json trong: {output_dir}")
        return []

    _banner(f"COMPILE  →  output/{folder_name}")
    print(f"  📋  {len(json_paths)} sheet_raw.json\n")

    written: list[str] = []
    errors: list[str]  = []

    for json_path in json_paths:
        try:
            with open(json_path, encoding="utf-8") as f:
                data = json.load(f)

            sheet_name = data.get("_meta", {}).get("source_sheet", "")
            vba_code   = compiler_func(data)

            # Ghi ra file .bas cùng thư mục với sheet_raw.json
            out_dir  = os.path.dirname(json_path)
            bas_name = _sanitize(sheet_name) + ".bas"
            bas_path = os.path.join(out_dir, bas_name)

            with open(bas_path, "w", encoding="utf-8") as f:
                f.write(vba_code)

            rel = os.path.relpath(bas_path, PROJECT_ROOT)
            print(f"  ✔  [{sheet_name}]  →  {rel}")
            written.append(bas_path)

        except Exception as e:
            print(f"  ❌  Lỗi [{json_path}]: {e}")
            errors.append(json_path)

    _banner(f"COMPILE DONE — {len(written)} file .bas  |  {len(errors)} lỗi")
    
    if written:
        print("\n💡 [NEXT STEP] Copy đoạn prompt dưới đây để giao việc cho AI Agent review/sửa code:")
        print("─" * 60)
        print(f"Hãy rà soát và chỉnh sửa code VBA hàng loạt cho các sheet trong thư mục `output/{folder_name}` so với spec `sheet_raw.json`. Yêu cầu: BỎ QUA việc lập Implementation Plan và chờ Approve. Agent hãy tự động phân tích, sửa code thẳng vào các file `.bas` nếu có sai lệch so với spec và `.agents/skills/vba-access-architect/SKILL.md`. Bắt buộc vẫn phải tự rà soát Checklist (Bước 4.3) trước khi hoàn tất mỗi file.")
        print("─" * 60)
        print("")
        
    return written


# ─────────────────────────────────────────────────────────
# CLI entry
# ─────────────────────────────────────────────────────────

def main():
    raw_args = sys.argv[1:]
    if not raw_args:
        print(__doc__)
        sys.exit(0)

    command = "run"
    folder = ""
    only_file = ""
    use_agent = False
    
    i = 0
    while i < len(raw_args):
        arg = raw_args[i]
        lower_arg = arg.lower()
        
        # Hỗ trợ cả trường hợp người dùng gõ dư dấu '--' (ví dụ: --scan, --compile)
        cmd_stripped = lower_arg.lstrip("-")
        if cmd_stripped in ("scan", "compile", "run", "all", "learn"):
            command = cmd_stripped
        elif arg == "--file":
            if i + 1 < len(raw_args):
                only_file = raw_args[i+1]
                i += 1
        elif arg == "--agent":
            use_agent = True
        elif arg.startswith("--"):
            folder = arg[2:]
        i += 1

    if not folder:
        print("❌  Vui lòng chỉ định folder: python main.py <command> --<folder>")
        print("    Ví dụ: python main.py run --test2")
        sys.exit(1)

    if command == "scan":
        cmd_scan(folder, only_file)
    elif command == "compile":
        cmd_compile(folder, None, only_file, use_agent)
    elif command == "learn":
        from config.compiler.learner import run_learner
        # Chạy learn trên thư mục output
        test_dir = os.path.join(PROJECT_ROOT, "output", folder)
        run_learner(test_dir)
    elif command in ("run", "all"):
        if only_file and not only_file.lower().endswith(".xlsx"):
            # Nếu truyền --file là tên sheet (vd: 契約_K)
            json_paths = cmd_scan(folder)
            cmd_compile(folder, json_paths, only_file, use_agent)
        else:
            # Nếu truyền --file là tên file xlsx (hoặc ko truyền)
            json_paths = cmd_scan(folder, only_file)
            cmd_compile(folder, json_paths, "", use_agent)
    else:
        print(f"❌  Lệnh không hợp lệ: {command}")
        print("    Dùng: scan | compile | run | learn")
        sys.exit(1)


if __name__ == "__main__":
    main()
