# zh-docx-safe

A Codex/Claude skill for stable Chinese DOCX/Word generation on Windows.

It prevents the usual failure modes around Chinese text and paths: UTF-8/GBK mojibake, PowerShell pipe corruption, Chinese path failures, East Asian font issues, mixed Chinese-English tables, and invalid DOCX zip/XML output.

中文简介：这是一个用于稳定生成中文 Word/DOCX 的 Codex/Claude skill，专门规避 Windows 中文路径、UTF-8/GBK 编码乱码、PowerShell 管道传中文、中文字体和中英混排表格等常见问题。

## What It Solves

- Chinese text becomes `???`, `锟斤拷`, or other mojibake.
- PowerShell pipes or here-strings corrupt Chinese-heavy scripts.
- Python on Windows falls back to GBK and raises `UnicodeEncodeError`.
- Chinese file paths fail when generating `.docx` reports.
- Word files render Chinese with wrong fonts.
- DOCX files are created but are not valid zip/XML packages.
- Mixed Chinese-English academic reports need stable headings, tables, bullets, and links.

## When To Use

Use this skill when an agent needs to create, edit, or validate `.docx` files that contain:

- Chinese text
- Chinese file paths
- mixed Chinese-English text
- academic paper titles or literature tables
- source links and report tables
- repeated Windows encoding problems

Pair it with a full DOCX skill for advanced Word features such as tracked changes, comments, or direct OOXML editing.

## Installation

With the Skills CLI:

```powershell
npx skills add baihe26/zh-docx-safe -g
```

Manual installation:

```powershell
git clone https://github.com/baihe26/zh-docx-safe.git "$env:USERPROFILE\.codex\skills\zh-docx-safe"
```

For Claude/Agents-style skill directories, clone or copy it to:

```powershell
$env:USERPROFILE\.agents\skills\zh-docx-safe
```

## Usage

Tell Codex/Claude something like:

```text
Use zh-docx-safe to generate a Chinese Word report from this content.
```

The skill instructs the agent to:

1. avoid passing Chinese-heavy code through PowerShell pipes;
2. write generation code or content to UTF-8 files;
3. run Python with UTF-8 mode;
4. set East Asian fonts explicitly;
5. validate the generated DOCX as a zip archive before reporting success.

## Helper Script

The bundled helper script can generate a Chinese-friendly `.docx` report from a UTF-8 JSON spec:

```powershell
$env:PYTHONUTF8='1'
python -X utf8 "$env:USERPROFILE\.codex\skills\zh-docx-safe\scripts\zh_docx_safe.py" --spec "D:\path\report_spec.json"
```

Generate a starter spec:

```powershell
$env:PYTHONUTF8='1'
python -X utf8 "$env:USERPROFILE\.codex\skills\zh-docx-safe\scripts\zh_docx_safe.py" --write-example "D:\path\basic-report.json"
```

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
    {"note_box": {"title": "注意", "items": ["中文路径要保持 UTF-8。"]}},
    {"quote": "Chinese text should survive the JSON -> Python -> DOCX pipeline."},
    {
      "table": {
        "headers": ["标题", "说明"],
        "rows": [["A", "中文说明"]]
      }
    }
  ],
  "sources": [
    ["GitHub", "https://github.com/baihe26/zh-docx-safe"]
  ]
}
```

## Repository Layout

```text
zh-docx-safe/
├── LICENSE
├── README.md
├── requirements.txt
├── examples/
│   └── basic-report.json
├── SKILL.md
├── agents/
│   └── openai.yaml
├── references/
│   └── windows-unicode.md
└── scripts/
    └── zh_docx_safe.py
```

## Notes

- The helper script requires `python-docx`.
- For formal Chinese manuscripts, prefer `宋体`; for internal reports, `微软雅黑` is usually more readable.
- If a terminal still corrupts Chinese paths, write content to UTF-8 files first and run scripts with `python -X utf8`.
