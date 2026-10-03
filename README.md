# doc-to-md

Converts Word documents (.docx) to clean, structured Markdown. Runs locally — no API calls.

Pipeline: mammoth (.docx → HTML) → turndown (HTML → Markdown) → post-processing:

- Headings come from Word heading styles and are numbered in code (`1.`, `1.1`, `1.1.1`).
- Tables are converted to Markdown tables (first row becomes the header).
- Cover page + TOC are dropped when the document has a Word TOC.
- Sections listed in `OMITTED_SECTIONS` (`src/index.ts`) are dropped: Έλεγχος Εγγράφου, Ιστορικό Αλλαγών, Ανασκόπηση, Διανομή.
- Images are replaced with an `*[image]*` marker.

> **Note:** Headings are only detected when the document uses Word heading styles (Heading 1/2/3). Bold text styled as a heading by hand stays bold text.

## Setup

```bash
npm install
```

## Usage

**Single file:**
```bash
ts-node src/index.ts report.docx
```

**Multiple files:**
```bash
ts-node src/index.ts report.docx spec.docx --output-dir ./output
```

**Entire directory:**
```bash
ts-node src/index.ts --dir ./docs --output-dir ./output
```

Each .docx produces a `.md` file with the same base name in the output directory (default: next to the source file).

> In PowerShell, call `npx ts-node src/index.ts ...` directly — `npm run dev -- ... --output-dir X` loses the flag to npm.

## Options

| Flag | Default | Description |
|------|---------|-------------|
| `--dir <path>` | — | Process all .docx files in a directory |
| `--output-dir <path>` | source file's folder | Directory to write `.md` files |
