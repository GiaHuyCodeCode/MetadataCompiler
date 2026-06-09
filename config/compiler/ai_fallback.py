"""
ai_fallback.py — MetadataCompiler V1
AI Agent fallback: sinh VBA snippet cho các sentence không nhận diện được,
dựa theo rules trong sk-architect SKILL.md, không tự suy đoán business logic.
"""

import os
import json
from pathlib import Path

# Path mặc định đến SKILL
_DEFAULT_SKILL = Path(__file__).parent.parent.parent / "agents" / "sk-architect" / "SKILL.md"


def _load_skill(skill_path: str | Path = _DEFAULT_SKILL) -> str:
    """Đọc nội dung SKILL.md."""
    try:
        return Path(skill_path).read_text(encoding="utf-8")
    except Exception:
        return ""


def _build_prompt(skill_content: str, sentence_json: dict, sheet_context: dict) -> str:
    """Tạo prompt gửi cho LLM."""
    ctx = {
        "sheet_name":    sheet_context.get("sheet_name", ""),
        "source_table":  sheet_context.get("source_table", ""),
        "target_table":  sheet_context.get("target_table", ""),
        "prefix":        sheet_context.get("prefix", ""),
        "sub_name":      sheet_context.get("sub_name", ""),
    }
    return f"""
# SKILL Reference
{skill_content[:6000]}

---

# Task
Sinh VBA code snippet cho một sentence từ sheet_raw.json mà compiler Python không nhận diện được.

## Sheet context
{json.dumps(ctx, ensure_ascii=False, indent=2)}

## Sentence JSON (rows)
{json.dumps(sentence_json.get("rows", []), ensure_ascii=False, indent=2)}

## Yêu cầu
1. Chỉ sinh phần VBA cho sentence này (KHÔNG sinh toàn bộ Sub).
2. Tuân thủ TUYỆT ĐỐI các rule trong SKILL (SQL = SQL & "...", alias T/T1/T2, log_write sau mỗi Execute, v.v.).
3. Không tự thêm business rule ngoài spec.
4. Trả về chỉ VBA code, không giải thích.
""".strip()


def ai_fallback_generate(
    sentence_json: dict,
    sheet_context: dict,
    skill_path: str | Path = _DEFAULT_SKILL,
    model: str = "gemini-2.5-flash",
) -> str:
    """
    Gọi LLM với SKILL.md + sentence JSON → sinh VBA snippet.
    Wrap output bằng comment '--- AI GENERATED ---'.

    Thực hiện kiểm tra cache cục bộ trong `ai_cache.json` trước để tăng tốc độ và độ tin cậy.
    Returns VBA string (có thể là stub nếu LLM không available).
    """
    # 1. Kiểm tra cache cục bộ trước
    try:
        cache_path = Path(__file__).parent / "ai_cache.json"
        if cache_path.exists():
            with open(cache_path, "r", encoding="utf-8") as f:
                cache = json.load(f)
            
            # Canonical key của sentence rows
            key = json.dumps(sentence_json.get("rows", []), sort_keys=True)
            if key in cache:
                cached_vba = cache[key]
                return (
                    "    ' --- AI GENERATED (sk-architect) ---\n"
                    + "\n".join("    " + l for l in cached_vba.splitlines())
                    + "\n    ' --- END AI GENERATED ---\n\n"
                )
    except Exception as cache_err:
        # Bỏ qua lỗi đọc cache và tiếp tục chạy AI fallback thông thường
        pass

    skill_content = _load_skill(skill_path)
    if not skill_content:
        return _stub(sentence_json, "SKILL.md không tìm thấy")

    # Thử dùng google.generativeai nếu available
    try:
        import google.generativeai as genai  # type: ignore

        api_key = os.environ.get("GOOGLE_API_KEY", "AQ.Ab8RN6IZyKYef3royZ4rrmOELvAkWD4Um0WWBP71ByNpPVIHtA")
        if not api_key:
            return _stub(sentence_json, "GOOGLE_API_KEY chưa set")

        env_model = os.environ.get("GEMINI_MODEL", "gemini-2.5-pro")
        if env_model:
            model = env_model.strip().lower().replace(" ", "-")

        genai.configure(api_key=api_key)
        m = genai.GenerativeModel(model)
        prompt = _build_prompt(skill_content, sentence_json, sheet_context)
        response = m.generate_content(prompt)
        vba_raw = response.text.strip()

        # Strip markdown fences nếu có
        if vba_raw.startswith("```"):
            lines = vba_raw.split("\n")
            vba_raw = "\n".join(
                l for l in lines
                if not l.startswith("```")
            ).strip()

        return (
            "    ' --- AI GENERATED (sk-architect) ---\n"
            + "\n".join("    " + l for l in vba_raw.splitlines())
            + "\n    ' --- END AI GENERATED ---\n\n"
        )

    except ImportError:
        return _stub(sentence_json, "google-generativeai không install")
    except Exception as e:
        return _stub(sentence_json, str(e)[:120])


def _stub(sentence_json: dict, reason: str) -> str:
    """Emit stub comment khi AI không available."""
    ctx = json.dumps(sentence_json.get("rows", [])[:2], ensure_ascii=False)
    return (
        f"    ' TODO:AI_REVIEW — AI fallback failed: {reason}\n"
        f"    ' Context: {ctx[:300]}\n\n"
    )
