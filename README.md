# nb2pdf_agent — Jupyter Notebook → Professional PDF Report

Convert any `.ipynb` Jupyter Notebook into a polished, ready-to-submit PDF lab
report — with syntax-highlighted code, formatted markdown, styled output blocks,
a table of contents, running header/footer, and page numbers.

---

## Features

| Feature | Detail |
|---------|--------|
| **Markdown rendering** | Headings H1–H4, bold/italic, inline code, blockquotes, ordered & unordered lists, tables, horizontal rules |
| **Syntax highlighting** | Catppuccin-dark theme — keywords, strings, numbers, comments all distinctly coloured |
| **Output blocks** | `stdout`/`stderr` in pale-blue blocks; error tracebacks in red blocks with `ValueError:` labels |
| **Embedded images** | `image/png` and `image/jpeg` cell outputs are decoded and inserted inline |
| **Table of Contents** | Auto-generated from H1/H2/H3 headings with page numbers (two-pass build) |
| **Running header** | Report title + generation date on every page |
| **Footer** | Script name + page number on every page |
| **Cover page** | Title, subtitle, source filename, kernel, and timestamp |

---

## Requirements

```
Python >= 3.10
reportlab >= 4.0
pygments >= 2.15
Pillow >= 10.0
```

Install all dependencies with:

```bash
pip install reportlab pygments Pillow
```

---

## Usage

### Basic — output named after the notebook

```bash
python nb2pdf_agent.py my_notebook.ipynb
# → my_notebook.pdf
```

### Custom output path

```bash
python nb2pdf_agent.py my_notebook.ipynb --output report.pdf
```

### Override report title

```bash
python nb2pdf_agent.py my_notebook.ipynb \
  --output lab3_report.pdf \
  --title "Lab 3 — Linear Regression Analysis"
```

### Full help

```bash
python nb2pdf_agent.py --help
```

---

## Output Structure

```
Generated PDF
├── Cover page       — title, kernel, generation timestamp
├── Table of Contents — H1/H2/H3 headings + page numbers
└── Notebook body
    ├── Markdown cells   — formatted prose with heading hierarchy
    ├── Code cells       — dark-themed syntax-highlighted blocks (In [N]:)
    └── Output cells     — styled output blocks (Out [N]:) + tracebacks
```

---

## Project Files

```
nb2pdf_agent.py      Main agent script
README.md            This file
sample_notebook.ipynb  Demo notebook (sales EDA)
sample_output.pdf    Demo PDF generated from sample_notebook.ipynb
```

---

## How It Works

1. **Parse** — `NotebookParser` reads the `.ipynb` JSON and extracts `cell_type`,
   `source`, `outputs`, and `execution_count` for every cell.

2. **Render markdown** — `MarkdownRenderer` walks each line and converts
   headings, lists, blockquotes, tables, and inline styles to ReportLab
   `Paragraph` / `Table` flowables.

3. **Highlight code** — `make_code_block()` uses Pygments to tokenise Python
   source and emits colour-tagged `<font>` XML in a dark-background `Table`.

4. **Style outputs** — `make_output_block()` dispatches to stream/error
   renderers, strips ANSI escapes, and wraps text in pale-blue or red tables.
   Base64-encoded images are decoded and inserted as `Image` flowables.

5. **Assemble** — cover page → TOC placeholder → all cells → `multiBuild()`
   for two-pass TOC page-number resolution.

6. **Header/footer** — `NBDocTemplate` (subclasses `BaseDocTemplate`) draws
   a brand-coloured header bar and grey footer on every page via `onPage`.

---

## Customisation

All colours are defined as constants at the top of `nb2pdf_agent.py`:

```python
C_BRAND   = "#2D3A8C"   # heading / header colour
C_ACCENT  = "#4F6FD8"   # subheading / accent
C_CODE_BG = "#1E1E2E"   # code block background (Catppuccin Mocha)
```

Change these to match your institution's brand colours.

---

## Known Limitations

- Rich HTML cell outputs (DataFrames rendered as HTML tables) are shown as
  plain text (`[HTML output — see notebook]`). Full HTML rendering would
  require a headless browser.
- LaTeX / MathJax equations in markdown are passed through as raw text.
- Very wide code lines may overflow the column — add `\n` line breaks in the
  source notebook for best results.
