# markitdown-converter

A small, robust CLI wrapper around [Microsoft MarkItDown](https://github.com/microsoft/markitdown) that turns almost any document into clean, token-efficient **Markdown** — built so an LLM can read the content accurately and cheaply, without the binary/markup noise of the source file.

Packaged as a [Claude Code](https://claude.com/claude-code) skill (`SKILL.md`), but the `scripts/convert.py` wrapper is a standalone tool you can run with plain Python.

## Supported inputs

PDF · Word (`.docx`) · Excel (`.xlsx`/`.xls`) · PowerPoint (`.pptx`) · images (metadata/EXIF + OCR) · audio (metadata + speech transcription) · HTML · CSV/TSV · JSON · XML/RSS · `.ipynb` · EPub · ZIP (iterates contents, incl. nested) · YouTube URLs.

## Install

```bash
pip install -r requirements.txt        # markitdown[all] + offline-fallback libs (+ test deps)
```

## Usage

```bash
python scripts/convert.py INPUT [INPUT ...] [-o OUTPUT] [options]
```

```bash
# Single file -> report.md (also echoed to stdout)
python scripts/convert.py report.pdf -o report.md

# A whole folder of mixed docs -> one .md per file in markdown/
python scripts/convert.py ./docs/ -o markdown/

# A YouTube URL -> transcript
python scripts/convert.py "https://www.youtube.com/watch?v=XXXX" -o transcript.md
```

Output rules:

| Invocation | Result |
|---|---|
| one input, no `-o` | writes `<basename>.md` in CWD **and** prints to stdout |
| one input, `-o file.md` | writes that file |
| one input, `-o dir/` | writes `<basename>.md` into the dir |
| many inputs / a dir, `-o dir/` | one `.md` per input (keeps original extension in the name; never overwrites on collision) |

Optional flags: `--use-plugins`, `--llm-model MODEL` (+`--llm-api-key`/`--llm-base-url`), `--docintel-endpoint URL`, `--recursive`.

## Graceful degradation

With `markitdown[all]` installed, every format uses MarkItDown's native backend at full fidelity. When an optional backend is missing — or it succeeds but returns nothing because a system tool (`exiftool`) or model (transcription/OCR) isn't present — the script falls back to a best-effort extractor and prints a `NOTE:` instead of failing or writing an empty file:

- `.docx` → `python-docx` when `mammoth` is absent
- `.xlsx` → `openpyxl` (**primary** path, so literal cell text like `NA`/`N/A`/`NULL` survives instead of being coerced to `NaN`)
- images → Pillow metadata/EXIF
- `.wav` → stdlib `wave` metadata; other audio → an informative stub

The run exits non-zero only if an input genuinely could not be read, and it never emits a raw Python traceback.

## Tests

```bash
python tests/run_all.py            # runs all suites
```

- `selftest.py` — quick per-format smoke test
- `test_suite.py` — 38 format / encoding / output / error checks
- `test_adversarial.py` — 18 edge cases (bad `-o`, collisions, big files, symlinks, encodings, stdout)
- `verify_native.py` — confirms network/system-dependent backends on your machine (`SKIP` when a backend isn't installed; only `FAIL` if an installed one errors)

The suites assert zero Python tracebacks and cover Cyrillic/cp1251, BOM, HTML entities, the literal-`NA` xlsx case, Markdown-table escaping, and non-ASCII stdout.

## Notes

- MarkItDown optimizes for structure (headings, lists, tables, links) over pixel-perfect visual fidelity — exactly what you want for feeding a model.
- Audio transcription additionally needs `ffmpeg`; image OCR/EXIF needs the `exiftool` binary (or pass `--llm-model` for LLM image captions).
- Security: MarkItDown reads with the current process's privileges — don't point it at untrusted URLs/paths without sanitizing first.
