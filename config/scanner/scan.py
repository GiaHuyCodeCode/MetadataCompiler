"""
scan.py — MetadataCompiler V1 CLI Entry Point

Usage:
    python scan.py --<folder_name>

Example:
    python scan.py --test1

This scans all .xlsx files inside  input/<folder_name>/
and writes JSON output to        output/<folder_name>/<sheet_name>/sheet_raw.json

Only sheets containing all 4 keywords are exported:
    ファイル名, 出力条件, 通常処理, 特殊処理
"""

import sys
import os
import json
import argparse
from datetime import datetime

# Allow running from project root
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT_ROOT)

from scanner.xlsx_scanner import scan_xlsx


# ─────────────────────────────────────────────────────────
# Sanitise sheet name → safe directory name
# ─────────────────────────────────────────────────────────

def _sanitize(name: str) -> str:
    invalid = r'/\:*?"<>|'
    result = name.strip()
    for ch in invalid:
        result = result.replace(ch, "_")
    return result


# ─────────────────────────────────────────────────────────
# Pretty separator helpers
# ─────────────────────────────────────────────────────────

def _sep(char="─", width=60):
    return char * width


def _banner(text: str):
    print(f"\n{'═'*60}")
    print(f"  {text}")
    print(f"{'═'*60}")


# ─────────────────────────────────────────────────────────
# Main scan logic
# ─────────────────────────────────────────────────────────

def run_scan(folder_name: str):
    input_dir  = os.path.join(PROJECT_ROOT, "input",  folder_name)
    output_dir = os.path.join(PROJECT_ROOT, "output", folder_name)

    if not os.path.isdir(input_dir):
        print(f"❌  Thư mục input không tồn tại: {input_dir}")
        sys.exit(1)

    # Collect xlsx files (skip Excel lock files starting with ~$)
    xlsx_files = [
        f for f in os.listdir(input_dir)
        if f.lower().endswith(".xlsx") and not f.startswith("~")
    ]

    if not xlsx_files:
        print(f"⚠️  Không tìm thấy file .xlsx trong: {input_dir}")
        sys.exit(0)

    _banner(f"SCAN  →  input/{folder_name}")
    print(f"  📂  Input  : {input_dir}")
    print(f"  📁  Output : {output_dir}")
    print(f"  📊  Tìm thấy {len(xlsx_files)} file xlsx\n")

    total_exported = 0
    total_skipped  = 0
    exported_paths = []   # list of (sheet_name, path)

    for xlsx_filename in sorted(xlsx_files):
        file_path = os.path.join(input_dir, xlsx_filename)
        file_stem = os.path.splitext(xlsx_filename)[0]

        print(f"{_sep()}")
        print(f"📄  File: {xlsx_filename}")

        try:
            scan_result = scan_xlsx(file_path)
        except Exception as exc:
            print(f"  ❌  Lỗi khi đọc file: {exc}")
            continue

        valid_sheets   = scan_result["valid_sheets"]
        skipped_sheets = scan_result["skipped_sheets"]
        data           = scan_result["data"]

        if skipped_sheets:
            print(f"  ⏭️   Bỏ qua {len(skipped_sheets)} sheet (thiếu keyword):")
            for s in skipped_sheets:
                print(f"        - {s}")

        if not valid_sheets:
            print(f"  ⚠️   Không có sheet nào hợp lệ trong file này.")
            total_skipped += len(skipped_sheets)
            continue

        print(f"  ✅  {len(valid_sheets)} sheet hợp lệ sẽ được xuất:")
        for sheet_name in valid_sheets:
            safe_sheet = _sanitize(sheet_name)
            sheet_out_dir = os.path.join(output_dir, file_stem, safe_sheet)
            os.makedirs(sheet_out_dir, exist_ok=True)

            out_path = os.path.join(sheet_out_dir, "sheet_raw.json")

            sheet_data = data[sheet_name]

            # Attach metadata
            output_payload = {
                "_meta": {
                    "generated_at": datetime.now().isoformat(),
                    "source_file": xlsx_filename,
                    "source_sheet": sheet_name,
                    "folder": folder_name,
                },
                "sections": {
                    "ファイル名": sheet_data["ファイル名"],
                    "出力条件":   sheet_data["出力条件"],
                    "通常処理":   sheet_data["通常処理"],
                    "特殊処理":   sheet_data["特殊処理"],
                },
            }

            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(output_payload, f, ensure_ascii=False, indent=2)

            rel_path = os.path.relpath(out_path, PROJECT_ROOT)
            print(f"        ✔  '{sheet_name}'  →  {rel_path}")
            exported_paths.append((sheet_name, out_path))
            total_exported += 1

        total_skipped += len(skipped_sheets)

    # ── Summary ──
    _banner("KẾT QUẢ")
    print(f"  ✅  Đã xuất  : {total_exported} sheet")
    print(f"  ⏭️   Bỏ qua  : {total_skipped} sheet")
    print(f"\n  📁  Đường dẫn output:")
    for sheet_name, path in exported_paths:
        rel = os.path.relpath(path, PROJECT_ROOT)
        print(f"      • [{sheet_name}]  →  {rel}")
    print()


# ─────────────────────────────────────────────────────────
# CLI entry
# ─────────────────────────────────────────────────────────

def main():
    # Accept --<folder_name> style arguments
    # Example: python scan.py --test1
    #          python scan.py --my_folder

    # Filter out standard flags; treat first --xxx as folder name
    args = sys.argv[1:]

    folder_name = None
    for arg in args:
        if arg.startswith("--"):
            folder_name = arg[2:]   # strip leading --
            break

    if not folder_name:
        print("❌  Vui lòng cung cấp tên folder: python scan.py --<tên_folder>")
        print("    Ví dụ: python scan.py --test1")
        sys.exit(1)

    run_scan(folder_name)


if __name__ == "__main__":
    main()
