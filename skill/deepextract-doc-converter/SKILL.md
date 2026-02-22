---
name: deepextract-doc-converter
description: Convert documents with local DeepExtract when users ask things like "帮我把这个文档转为xx格式", "转成 Word", "转成 Markdown", or "convert this file to docx/markdown". Use for PDF, images, Word, PPT, HTML, and Markdown inputs, with outputs in Markdown or DOCX.
---

# DeepExtract Document Converter

Use this skill to execute real local file conversion through the DeepExtract codebase.

## Workflow

1. Parse user intent:
   - Input file path
   - Target format (`markdown` or `docx`)
2. Run the bundled script:

```bash
python "$HOME/.config/opencode/skills/deepextract-doc-converter/scripts/convert_with_deepextract.py" --input "<input_path>" --target "<target>"
```

For Word output, pass style options when user specifies them, for example:

```bash
python "$HOME/.config/opencode/skills/deepextract-doc-converter/scripts/convert_with_deepextract.py" \
  --input "<input_path>" \
  --target docx \
  --line-spacing 1.75 \
  --font-zh "宋体" \
  --font-size-body 12
```

3. If user gave an explicit output file path, add `--output`.
4. Return the generated result path to the user.

## Target format mapping

- `markdown`, `md` -> Markdown output
- `word`, `docx`, `doc` -> DOCX output

## Word style options

- `--font-zh` Chinese font, e.g. `宋体`, `微软雅黑`
- `--font-en` English font, e.g. `Calibri`, `Times New Roman`
- `--font-size-body` body font size
- `--font-size-h1` / `--font-size-h2` / `--font-size-h3` / `--font-size-h4`
- `--line-spacing` line spacing
- `--paragraph-spacing` paragraph spacing (pt)
- `--formula-numbering-mode` one of: `none`, `chapter_index`, `global`, `chapter`

## Notes

- The script auto-detects project root from current directory or parent directories.
- For global usage in any directory, set `DEEPEXTRACT_ROOT` to your DeepExtract project path.
- For MinerU-based conversions, API key is required:
  - `MINERU_API_KEY` environment variable, or
  - `apikey.md` with `MINERU_API_KEY=...`
- If both input path and target are clear, run directly without asking extra questions.
