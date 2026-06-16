# MarkItDown reference

Source: https://github.com/microsoft/markitdown (Microsoft, MIT license). Python 3.10+.

## Install

```bash
pip install 'markitdown[all]'          # everything
pip install 'markitdown[pdf, docx, pptx]'   # selected formats only
```

Optional dependency groups: `[all]`, `[pptx]`, `[docx]`, `[xlsx]`, `[xls]`, `[pdf]`,
`[outlook]`, `[az-doc-intel]`, `[az-content-understanding]`, `[audio-transcription]`,
`[youtube-transcription]`.

## CLI

```bash
markitdown path-to-file.pdf > document.md
markitdown path-to-file.pdf -o document.md
cat path-to-file.pdf | markitdown
```

## Python API

```python
from markitdown import MarkItDown

md = MarkItDown(enable_plugins=False)
result = md.convert("test.xlsx")
print(result.text_content)   # newer versions also expose result.markdown
```

### LLM image descriptions (images / pptx)

```python
from markitdown import MarkItDown
from openai import OpenAI

client = OpenAI()
md = MarkItDown(llm_client=client, llm_model="gpt-4o")
result = md.convert("example.jpg")
```

### Azure Document Intelligence (higher-quality PDF/scan extraction)

```bash
markitdown path-to-file.pdf -o out.md -d -e "<doc_intelligence_endpoint>"
```

## Supported inputs

PDF, PowerPoint, Word, Excel, images (EXIF + OCR), audio (EXIF + speech transcription),
HTML, text-based formats (CSV, JSON, XML), ZIP (iterates contents), YouTube URLs, EPub, and more.

## Plugins

Disabled by default. `markitdown --list-plugins`, enable with `--use-plugins`.
Find third-party plugins via the GitHub hashtag `#markitdown-plugin`.

## Security

MarkItDown performs I/O with the privileges of the current process (like `open()` /
`requests.get()`). Sanitize untrusted input; prefer the narrowest method
(`convert_local()`, `convert_stream()`, `convert_response()`) over the permissive `convert()`
in hosted/server contexts.
