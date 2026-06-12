---
name: zh-docx-safe
description: Stabilize Chinese Word/document generation on Windows. Use when creating, editing, or validating .docx/Word outputs that contain Chinese text, Chinese file paths, mixed Chinese-English tables, links, or long academic titles; also use when prior attempts hit mojibake, GBK/UTF-8 errors, PowerShell pipe issues, or Chinese path failures. Pair with the docx skill for advanced Word features.
---

# Zh Docx Safe

## Purpose

Use this as a safety layer for Chinese `.docx` work on Windows. It prevents the recurring failures caused by PowerShell pipe encoding, GBK defaults, Chinese paths, Chinese fonts, and fragile DOCX zip/XML output.

## Non-Negotiables

- Do not pass large Chinese text through PowerShell pipes, here-strings, `echo`, `cat`, or inline `python -c`.
- Write reusable generation code to a UTF-8 `.py` file, then run it with:

```powershell
$env:PYTHONUTF8='1'; python -X utf8 path\to\script.py
```

- Use raw Python paths, e.g. `Path(r"D:\中文目录\文件.docx")`.
- Set Chinese fonts explicitly: usually `微软雅黑` for reports, `宋体` for formal manuscripts, and `Arial` for English fallback.
- Validate generated `.docx` as a zip archive with `ZipFile.testzip()` or the docx skill validator before reporting success.
- Keep temporary scripts/files prefixed with `CX_tmp_` or delete them after successful generation.

## Workflow

1. If the task needs complex Word features, also use the `docx` skill. This skill handles Chinese/Windows reliability, not every OOXML detail.
2. Draft content in memory or a UTF-8 JSON/Markdown source file.
3. Generate Word from a UTF-8 Python script file, not from shell-piped Chinese.
4. Use `python-docx` for ordinary reports and tables. Use the existing `docx` skill's XML workflow only when tracked changes, comments, or advanced DOCX internals are required.
5. Validate:
   - output path exists
   - file size is nonzero
   - `zipfile.ZipFile(output).testzip()` returns `None`
6. In the final response, link the absolute output path and mention any approximate or time-sensitive metadata that needs later verification.

## Helper Script

For routine Chinese reports, use `scripts/zh_docx_safe.py`. It reads a UTF-8 JSON spec and writes a `.docx` with Chinese fonts, headings, paragraphs, bullets, tables, source links, and zip validation.

Example JSON:

```json
{
  "title": "中文报告标题",
  "subtitle": "CX 生成：2026-06-12",
  "output": "D:/path/CX_中文报告.docx",
  "sections": [
    {"heading": "核心判断", "level": 1},
    {"paragraph": "这里是中文正文。"},
    {"bullets": ["第一点", "第二点"]},
    {"table": {"headers": ["标题", "说明"], "rows": [["A", "中文说明"]]}}
  ]
}
```

Run:

```powershell
$env:PYTHONUTF8='1'; python -X utf8 "C:\Users\柴鱼\.codex\skills\zh-docx-safe\scripts\zh_docx_safe.py" --spec "D:\path\CX_tmp_spec.json"
```

## Debugging

If Chinese text becomes `???`, `锟斤拷`, or path creation fails, read `references/windows-unicode.md` and rerun from a UTF-8 script file. The common cause is not DOCX itself; it is usually shell or process encoding before the document is even created.
