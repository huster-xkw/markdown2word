#!/usr/bin/env python3
import argparse
import os
import shutil
import sys
from pathlib import Path


def find_project_root(start: Path) -> Path:
    env_root = os.getenv("DEEPEXTRACT_ROOT", "").strip()
    if env_root:
        p = Path(env_root).expanduser().resolve()
        if (p / "md2word_final.py").exists() and (p / "mineru_extract.py").exists():
            return p
        raise FileNotFoundError(
            f"DEEPEXTRACT_ROOT is set but invalid: {p}. Missing md2word_final.py or mineru_extract.py"
        )

    cur = start.resolve()
    candidates = [cur] + list(cur.parents)
    for p in candidates:
        if (p / "md2word_final.py").exists() and (p / "mineru_extract.py").exists():
            return p
    raise FileNotFoundError("Cannot locate DeepExtract project root.")


def normalize_target(target: str) -> str:
    t = target.strip().lower()
    if t in {"md", "markdown"}:
        return "md"
    if t in {"doc", "docx", "word"}:
        return "docx"
    raise ValueError(f"Unsupported target format: {target}")


def ensure_key_notice(project_root: Path) -> None:
    has_env = bool(os.getenv("MINERU_API_KEY", "").strip())
    key_file = project_root / "apikey.md"
    has_file_key = False
    if key_file.exists():
        for line in key_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith("MINERU_API_KEY=") and line.split("=", 1)[1].strip():
                has_file_key = True
                break
    if not has_env and not has_file_key:
        print(
            "[WARN] MINERU_API_KEY not found. Conversions requiring MinerU may fail.",
            file=sys.stderr,
        )


def convert_markdown_to_docx(
    project_root: Path,
    input_file: Path,
    output_file: Path,
    docx_options: dict,
    formula_numbering_mode: str,
) -> Path:
    sys.path.insert(0, str(project_root))
    import md2word_final

    output_file.parent.mkdir(parents=True, exist_ok=True)
    md2word_final.convert_with_python_docx(
        str(input_file),
        str(output_file),
        formula_numbering_mode=formula_numbering_mode or None,
        doc_style_options=(docx_options or None),
    )
    return output_file


def convert_with_mineru(
    project_root: Path,
    input_file: Path,
    target: str,
    output_file: Path,
    docx_options: dict,
    formula_numbering_mode: str,
) -> Path:
    sys.path.insert(0, str(project_root))
    import mineru_extract
    import md2word_final

    result = mineru_extract.upload_and_extract(str(input_file))
    output_file.parent.mkdir(parents=True, exist_ok=True)

    if target == "md":
        src_zip = Path(result["zip_path"])
        if output_file.suffix.lower() != ".zip":
            output_file = output_file.with_suffix(".zip")
        shutil.copy2(src_zip, output_file)
        return output_file

    md2word_final.convert_with_python_docx(
        result["md_path"],
        str(output_file),
        formula_numbering_mode=formula_numbering_mode or None,
        doc_style_options=(docx_options or None),
    )
    return output_file


def build_default_output(input_file: Path, target: str) -> Path:
    if target == "md":
        return input_file.with_suffix(".zip")
    return input_file.with_suffix(".docx")


def build_docx_options(args: argparse.Namespace) -> tuple[dict, str]:
    options = {}
    numeric_keys = {
        "font_size_body",
        "font_size_h1",
        "font_size_h2",
        "font_size_h3",
        "font_size_h4",
        "line_spacing",
        "paragraph_spacing",
    }
    all_keys = numeric_keys | {"font_zh", "font_en"}

    for key in all_keys:
        value = getattr(args, key)
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        options[key] = value

    return options, (args.formula_numbering_mode or "").strip().lower()


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert files with local DeepExtract")
    parser.add_argument("--input", required=True, help="Input file path")
    parser.add_argument("--target", required=True, help="Target format: markdown/docx")
    parser.add_argument("--output", default="", help="Optional output path")
    parser.add_argument("--font-zh", dest="font_zh", default=None)
    parser.add_argument("--font-en", dest="font_en", default=None)
    parser.add_argument("--font-size-body", dest="font_size_body", type=int, default=None)
    parser.add_argument("--font-size-h1", dest="font_size_h1", type=int, default=None)
    parser.add_argument("--font-size-h2", dest="font_size_h2", type=int, default=None)
    parser.add_argument("--font-size-h3", dest="font_size_h3", type=int, default=None)
    parser.add_argument("--font-size-h4", dest="font_size_h4", type=int, default=None)
    parser.add_argument("--line-spacing", dest="line_spacing", type=float, default=None)
    parser.add_argument("--paragraph-spacing", dest="paragraph_spacing", type=float, default=None)
    parser.add_argument("--formula-numbering-mode", dest="formula_numbering_mode", default=None)
    args = parser.parse_args()

    input_file = Path(args.input).expanduser().resolve()
    if not input_file.exists() or not input_file.is_file():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    target = normalize_target(args.target)
    output_file = (
        Path(args.output).expanduser().resolve()
        if args.output.strip()
        else build_default_output(input_file, target)
    )

    project_root = find_project_root(Path.cwd())
    docx_options, formula_numbering_mode = build_docx_options(args)

    ext = input_file.suffix.lower()
    markdown_input = {".md", ".markdown"}

    if ext in markdown_input and target == "docx":
        final_path = convert_markdown_to_docx(
            project_root,
            input_file,
            output_file,
            docx_options,
            formula_numbering_mode,
        )
    elif ext in markdown_input and target == "md":
        if output_file.suffix.lower() not in {".md", ".markdown"}:
            output_file = output_file.with_suffix(".md")
        output_file.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(input_file, output_file)
        final_path = output_file
    else:
        ensure_key_notice(project_root)
        final_path = convert_with_mineru(
            project_root,
            input_file,
            target,
            output_file,
            docx_options,
            formula_numbering_mode,
        )

    print(str(final_path))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
