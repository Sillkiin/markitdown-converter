---
name: markitdown-converter
description: Convert documents and files into clean Markdown for use with LLMs using Microsoft's MarkItDown. Use this skill whenever the user wants to turn a PDF, Word (.docx), Excel (.xlsx), PowerPoint (.pptx), image, audio file, HTML, CSV, JSON, XML, EPub, ZIP, or YouTube URL into Markdown - or whenever they ask to run MarkItDown on a file, say "convert to markdown", "extract clean text from this document for the model", "prep this file for Claude/an LLM", or complain that a raw PDF is messy / wastes tokens. Trigger even if the user doesn't say the word "markdown" but clearly wants the readable text content pulled out of a document so a model can work with it accurately and cheaply. Do NOT trigger when the user only wants a summary, translation, or analysis of a file's content, is authoring Markdown from scratch with no source file, or is merely asking what MarkItDown is or how to install it - trigger only when the deliverable is the file's own content rendered as Markdown.
---

# MarkItDown Converter

Convert almost any document into clean, token-efficient Markdown using [Microsoft MarkItDown](https://github.com/microsoft/markitdown). The point is to strip the messy binary/markup of source files (PDF, Office docs, images, etc.) down to structured plain text that an LLM reads accurately and cheaply.

## When to use this

Reach for this skill whenever the real goal is "get the readable content out of this file so a model can use it." Common triggers: a user hands over a PDF/Word/Excel/PowerPoint/image/audio file and wants its text; a user says a raw PDF is bloated or burns tokens; a user explicitly names MarkItDown; a user wants a YouTube transcript as Markdown.

Supported inputs: PDF, Word (.docx), PowerPoint (.pptx), Excel (.xlsx/.xls), images (EXIF + OCR), audio (EXIF + speech transcription), HTML, text formats (CSV, JSON, XML), ZIP (iterates contents), EPub, and YouTube URLs.

## How to convert

Use the bundled wrapper script `scripts/convert.py`. It handles installation-on-demand and works on single files, multiple files, directories, and URLs.

```bash
python scripts/convert.py INPUT [INPUT ...] [-o OUTPUT] [options]
```

Output behavior:
- One input, no `-o`: writes `<basename>.md` in the current directory and also prints the Markdown to stdout.
- One input with `-o file.md`: writes to that file.
- One input with `-o dir/`: writes `<basename>.md` into that directory.
- Multiple inputs or a directory with `-o dir/`: writes one `.md` per input into that directory, keeping the original extension in the name (`a.csv` -> `a.csv.md`) so files with the same stem don't collide.

Examples:

```bash
# Single PDF -> report.md, also echoed to stdout
python scripts/convert.py /path/to/report.pdf -o report.md

# Whole folder of mixed docs -> markdown/ directory
python scripts/convert.py /path/to/docs/ -o markdown/

# A YouTube URL -> transcript.md
python scripts/convert.py "https://www.youtube.com/watch?v=XXXX" -o transcript.md
```

## Full MarkItDown functionality

The wrapper exposes MarkItDown's complete feature set through optional flags:

- `--use-plugins` - enable installed third-party MarkItDown plugins (search GitHub `#markitdown-plugin`).
- `--llm-model MODEL` - use an LLM (e.g. `gpt-4o`) to describe images inside documents. Requires the `openai` package. Pair with `--llm-api-key` (or the `OPENAI_API_KEY` env var) and optionally `--llm-base-url` for an OpenAI-compatible endpoint.
- `--docintel-endpoint URL` - route extraction through Azure Document Intelligence for higher-quality OCR on scanned PDFs and complex layouts.

```bash
# Describe images in a PDF using an LLM
python scripts/convert.py scan.pdf -o scan.md --llm-model gpt-4o --llm-api-key sk-...

# High-quality scanned-PDF extraction via Azure Document Intelligence
python scripts/convert.py contract.pdf -o contract.md --docintel-endpoint "https://<resource>.cognitiveservices.azure.com/"
```

To get every backend, install the full package once:

```bash
pip install 'markitdown[all]'            # add --break-system-packages on a managed Python
```

## Format support and graceful degradation

With `markitdown[all]` installed, every format is handled by MarkItDown's native backend at full fidelity. When an optional backend is missing — or when it succeeds but returns nothing because a system tool (exiftool) or model (transcription/OCR) isn't present — the script falls back to a best-effort extractor and prints a `NOTE:` instead of failing or writing an empty file. Bundled fallbacks:

- `.docx` → `python-docx` (headings, bullet lists, tables) when `mammoth` is absent.
- `.xlsx` → `openpyxl` (one Markdown table per sheet) is the **primary** path so literal cell text like `"NA"`/`"N/A"`/`"NULL"` is preserved; MarkItDown's native pandas reader coerces those to `NaN`, which would corrupt e.g. a region code. Native MarkItDown is the fallback.
- images (`.png/.jpg/.jpeg/.gif/.bmp/.tiff/.webp`) → Pillow metadata: format, dimensions, mode, embedded text chunks, EXIF tags.
- `.wav` → stdlib `wave` metadata: channels, sample rate, duration.
- other audio (`.mp3/.m4a/.flac/.ogg`) → an informative stub (type, size) noting that transcription needs `markitdown[all]`.

These fallbacks give partial-but-useful output offline (e.g. image metadata instead of an OCR/LLM caption); installing `markitdown[all]` upgrades each to the full native path automatically. When several inputs are given and one fails, the others still convert; the run exits non-zero only if at least one input genuinely could not be read. The script never emits a raw Python traceback.

## Verifying it works

Install the full stack first (pinned in `requirements.txt`):

```bash
pip install -r requirements.txt          # markitdown[all] + fallback libs
```

> **Running the converter needs no manual setup.** `convert.py` makes a compatible
> MarkItDown available itself: it uses an importable `markitdown>=0.1.6`, else installs
> the pinned `markitdown[all]>=0.1.6` (the pin avoids pip silently downgrading to an
> incompatible old release), and as a last resort — e.g. a Python like 3.14 with no
> compatible wheels yet — builds a dedicated managed venv under
> `~/.markitdown-converter/venv` with a compatible Python (3.10–3.13) and re-execs there.
> The `pip install` above is for running the **test suite** and pre-warms the runtime.

Then run the bundled tests — `tests/run_all.py` runs all of them:

```bash
python tests/run_all.py            # all suites below, in sequence
python tests/test_suite.py         # 48 format/encoding/output/error checks
python tests/test_adversarial.py   # 18 edge cases (bad -o, collisions, big files, symlinks…)
python tests/selftest.py           # quick smoke test
python tests/verify_native.py      # confirm network/system-dependent backends (see below)
```

The first three suites cover every supported format (PDF, DOCX, XLSX, PPTX, CSV, TSV, HTML, JSON, XML, RSS, ipynb, EPub, ZIP/nested ZIP, images, WAV/audio), encodings (Cyrillic, BOM, HTML entities), the literal-`NA` xlsx regression, output modes, and error handling — and assert zero Python tracebacks.

### Verifying the network/system-dependent backends

Some backends can't be exercised in a locked-down/offline build (no PyPI, no `exiftool`): **native `.docx` (mammoth), audio transcription, YouTube transcripts, image OCR/EXIF, legacy `.xls`, Outlook `.msg`**. `tests/verify_native.py` checks these on *your* machine. It reports `SKIP` (with the exact install command) for any backend that isn't installed and only `FAIL`s if an installed backend errors — so a clean run is meaningful. Confirm the three that need real input or network by pointing it at real files/URLs:

```bash
python tests/verify_native.py                              # backend inventory + native docx/image checks
python tests/verify_native.py --audio path/to/speech.mp3   # confirm a real transcription
python tests/verify_native.py --image path/to/photo.jpg    # confirm native image metadata/OCR
python tests/verify_native.py --youtube https://youtu.be/XXXX   # confirm a real transcript
```

Audio transcription additionally needs `ffmpeg` and `pip install 'markitdown[audio-transcription]'`; image OCR/EXIF needs the `exiftool` binary (or use `--llm-model` for LLM captions).

## Notes and good practice

- MarkItDown is built for machine consumption, not pixel-perfect human reproduction - it prioritizes preserving structure (headings, lists, tables, links) over exact visual fidelity. That's exactly what you want for feeding a model.
- Image and audio conversion does OCR / speech-to-text; quality depends on the source. For scanned PDFs, consider `--docintel-endpoint` or `--llm-model`.
- Security: MarkItDown reads with the current process's privileges. Don't point it at untrusted URLs or paths in a sensitive environment without sanitizing first.
- After converting, sanity-check the `.md` (tables and multi-column PDFs are the usual trouble spots) before handing it to the model.

See `references/markitdown.md` for the full capability/option reference.
