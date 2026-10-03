import * as fs from "fs";
import * as path from "path";
import mammoth from "mammoth";
import TurndownService from "turndown";

// ─── Word → Markdown ──────────────────────────────────────────────────────────
// mammoth → HTML (keeps Word heading styles and tables) → turndown → Markdown.

async function readDocx(filePath: string, dropToc: boolean, omitted: string[]): Promise<string> {
    const result = await mammoth.convertToHtml({ path: filePath });
    const markdown = buildTurndown().turndown(result.value);
    return numberHeadings(omitSections(markdown, dropToc, omitted));
}

// Minimal shape of the DOM nodes turndown hands to rules (tsconfig has no DOM lib)
interface DomNode {
    nodeName: string;
    childNodes: ArrayLike<DomNode>;
    innerHTML: string;
    getAttribute(name: string): string | null;
}

function buildTurndown(): TurndownService {
    const td = new TurndownService({ headingStyle: "atx", bulletListMarker: "-", codeBlockStyle: "fenced" });
    // Images arrive as inline base64 (hundreds of KB each) — leave a marker instead.
    // addRule, not remove(): turndown's built-in image rule takes precedence over remove().
    td.addRule("image", { filter: "img", replacement: () => "*[image]*" });
    td.addRule("table", {
        filter: "table",
        replacement: (_content, node) => tableToMarkdown(td, node as unknown as DomNode),
    });
    return td;
}

// Mammoth tables have no header rows, which turndown's GFM plugin requires,
// so tables are built here: first row becomes the header, cells keep inline formatting.
function tableToMarkdown(td: TurndownService, table: DomNode): string {
    const rows: string[][] = [];
    const collectRows = (parent: DomNode) => {
        for (const child of Array.from(parent.childNodes)) {
            if (child.nodeName === "TR") rows.push(rowCells(td, child));
            else if (["THEAD", "TBODY", "TFOOT"].includes(child.nodeName)) collectRows(child);
        }
    };
    collectRows(table);

    const nonEmpty = rows.filter((r) => r.some((c) => c !== ""));
    if (nonEmpty.length === 0) return "";

    const width = Math.max(...nonEmpty.map((r) => r.length));
    const line = (r: string[]) =>
        `| ${[...r, ...Array(width - r.length).fill("")].join(" | ")} |`;

    return [
        "",
        line(nonEmpty[0]),
        `|${" --- |".repeat(width)}`,
        ...nonEmpty.slice(1).map(line),
        "",
    ].join("\n") + "\n";
}

function rowCells(td: TurndownService, row: DomNode): string[] {
    const cells: string[] = [];
    for (const cell of Array.from(row.childNodes)) {
        if (cell.nodeName !== "TD" && cell.nodeName !== "TH") continue;
        const text = td
            .turndown(cell.innerHTML)
            .trim()
            .replace(/\n+/g, "<br>") // Markdown table cells must be single-line
            .replace(/\|/g, "\\|");
        // Pad merged columns so the following cells stay aligned
        const span = Number(cell.getAttribute("colspan") ?? 1);
        cells.push(text, ...Array(span - 1).fill(""));
    }
    return cells;
}

// Drops every section whose heading contains one of `omitted`, down to the next heading
// of equal or higher level. Content before the first heading is dropped only when it
// contains a Word TOC (cover page + TOC) and dropToc is set; documents without headings are kept whole.
function omitSections(markdown: string, dropToc: boolean, omitted: string[]): string {
    const preamble: string[] = [];
    const out: string[] = [];
    let skipUntilLevel = 0; // 0 = not skipping
    let seenHeading = false;

    for (const line of markdown.split("\n")) {
        const heading = /^(#{1,6})\s+(.*)$/.exec(line);
        if (heading) {
            const level = heading[1].length;
            seenHeading = true;
            if (skipUntilLevel && level > skipUntilLevel) continue;
            skipUntilLevel = omitted.some((s) => heading[2].includes(s)) ? level : 0;
            if (skipUntilLevel) continue;
        }
        if (!seenHeading) preamble.push(line);
        else if (!skipUntilLevel) out.push(line);
    }

    const hasToc = dropToc && preamble.some((l) => l.includes("](#_Toc"));
    return [...(hasToc ? [] : preamble), ...out].join("\n").trim() + "\n";
}

// Numbers headings (1., 1.1, 1.1.1) from their level — Word's auto-numbering is not in mammoth's output.
function numberHeadings(markdown: string): string {
    const counters: number[] = [];
    let inFence = false;

    return markdown
        .split("\n")
        .map((line) => {
            if (line.startsWith("```")) inFence = !inFence;
            const match = !inFence && /^(#{1,6})\s+(.*)$/.exec(line);
            if (!match) return line;

            const level = match[1].length;
            counters.length = level;
            for (let i = 0; i < level; i++) counters[i] ??= 0;
            counters[level - 1]++;

            // Drop any numbering the model or the source already put there
            const text = match[2].replace(/^\d{1,2}(\.\d{1,2})*\.?\s+/, "");
            const number = level === 1 ? `${counters[0]}.` : counters.join(".");
            return `${match[1]} ${number} ${text}`;
        })
        .join("\n");
}

// ─── Input files ──────────────────────────────────────────────────────────────

function isDocx(filename: string): boolean {
    // "~$" files are Word's lock files for open documents
    return path.extname(filename).toLowerCase() === ".docx" && !path.basename(filename).startsWith("~$");
}

// Expands directories into the .docx files they contain
function collectFiles(inputs: string[]): string[] {
    const files: string[] = [];
    for (const input of inputs) {
        if (fs.statSync(input).isDirectory()) {
            const entries = fs.readdirSync(input, { withFileTypes: true });
            files.push(...entries.filter((e) => e.isFile() && isDocx(e.name)).map((e) => path.join(input, e.name)));
        } else if (isDocx(input)) {
            files.push(input);
        } else {
            throw new Error(`Unsupported file type (expected .docx): ${input}`);
        }
    }
    if (files.length === 0) throw new Error("No Word (.docx) files found in the provided paths.");
    return files;
}

function outputPath(filePath: string, outputDir: string): string {
    return path.join(outputDir, `${path.basename(filePath, path.extname(filePath))}.md`);
}

// ─── Main ─────────────────────────────────────────────────────────────────────

function parseFlag(args: string[], flag: string): string | undefined {
    const idx = args.indexOf(flag);
    if (idx !== -1 && args[idx + 1] !== undefined) {
        const value = args[idx + 1];
        args.splice(idx, 2);
        return value;
    }
    return undefined;
}

async function main(): Promise<void> {
    const args = process.argv.slice(2);

    const outputDir = parseFlag(args, "--output-dir");
    const dropTocArg = parseFlag(args, "--drop-toc") ?? "true";
    if (!["true", "false"].includes(dropTocArg)) {
        console.error(`--drop-toc expects true or false, got: ${dropTocArg}`);
        process.exit(1);
    }
    const dropToc = dropTocArg === "true";
    const omitArg = parseFlag(args, "--omit-sections");
    // Sections dropped from the output, matched against heading text; none by default
    const omitted = (omitArg ?? "").split(",").map((s) => s.trim()).filter((s) => s !== "");
    const dir = parseFlag(args, "--dir");
    if (dir) args.push(dir);

    if (args.length === 0) {
        console.error(
            "Usage:\n" +
            "  ts-node src/index.ts <file.docx> [file2.docx ...] [--output-dir ./output] [--drop-toc true|false] [--omit-sections \"A,B\"]\n" +
            "  ts-node src/index.ts --dir <directory> [--output-dir ./output] [--drop-toc true|false] [--omit-sections \"A,B\"]"
        );
        process.exit(1);
    }

    if (outputDir && !fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    const files = collectFiles(args);
    console.log(`Converting ${files.length} document(s)...`);

    let successCount = 0;
    for (const file of files) {
        process.stdout.write(`  → ${path.basename(file)} ... `);
        try {
            const markdown = await readDocx(file, dropToc, omitted);
            const outFile = outputPath(file, outputDir ?? path.dirname(file));
            fs.writeFileSync(outFile, markdown, "utf-8");
            console.log(`saved → ${outFile}`);
            successCount++;
        } catch (err) {
            console.error(`FAILED — ${(err as Error).message}`);
        }
    }

    console.log(`\nDone. ${successCount}/${files.length} file(s) converted.`);
}

main().catch((err) => {
    console.error(err.message);
    process.exit(1);
});
