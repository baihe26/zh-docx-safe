# Windows Unicode Notes For Chinese DOCX

## Common Failure Modes

- `???` in a path means the string was already corrupted before Python saw it. Re-running the same command with a different Python writer will not recover the original characters.
- `UnicodeEncodeError: 'gbk' codec can't encode ...` means Python stdout/file defaults are using the Windows ANSI code page.
- Mojibake like `锟斤拷` often means UTF-8 bytes were decoded as GBK or vice versa.
- PowerShell may preserve Unicode interactively but still corrupt text when long Chinese code is embedded in command arguments or piped into another process.

## Stable Pattern

1. Put Chinese-heavy code or content into a UTF-8 file.
2. Run Python with both:
   - `$env:PYTHONUTF8='1'`
   - `python -X utf8 script.py`
3. Use `Path(r"...")` for Windows paths.
4. Avoid building shell commands that contain long Chinese strings.
5. For generated Word files, always test the output zip.

## Good Python Defaults

```python
from pathlib import Path
import json

spec = json.loads(Path(spec_path).read_text(encoding="utf-8"))
out = Path(spec["output"])
out.parent.mkdir(parents=True, exist_ok=True)
```

## DOCX Font Notes

`python-docx` must set East Asian fonts separately:

```python
run.font.name = "微软雅黑"
run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
```

For academic reports, use `微软雅黑` for readable internal reports. Use `宋体` only when a formal Chinese manuscript style is requested.
