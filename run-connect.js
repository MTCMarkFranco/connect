/**
 * run-connect.js
 *
 * Single orchestration script that runs all steps end-to-end:
 *   1. Scrape Power BI report (Playwright + Edge)
 *   2. Summarise metrics via Azure OpenAI (multi-modal)
 *   3. Merge fleet instructions + metrics into one self-contained /fleet prompt
 *   4. Copy the merged prompt to clipboard
 *   5. Ensure Copilot CLI is authenticated (auto-login if needed)
 *   6. Launch GitHub Copilot CLI interactively
 *
 * Usage:
 *   node run-connect.js --quarter FY26Q3
 *   node run-connect.js --quarter FY26Q3 --headless
 *   node run-connect.js --quarter FY26Q3 --headless --date-range "Jan 1, 2026 - Mar 31, 2026"
 *   node run-connect.js --skip-scrape --quarter FY26Q3   # reuse existing final-metrics.md
 *   node run-connect.js --skip-to-copilot --quarter FY26Q3 # jump straight to Copilot CLI
 *   node run-connect.js --word-only --quarter FY26Q3       # generate final.docx from existing temp/ files
 */

const { execFileSync, execSync } = require("child_process");
const fs = require("fs");
const path = require("path");
const docx = require("docx");

// ── Parse CLI args ─────────────────────────────────────────────────────────
const args = process.argv.slice(2);
function getArg(name) {
  const idx = args.indexOf(name);
  return idx !== -1 && idx + 1 < args.length ? args[idx + 1] : null;
}
const quarter = getArg("--quarter");
const dateRange = getArg("--date-range");
const headless = args.includes("--headless");
const skipScrape = args.includes("--skip-scrape");
const skipToCopilot = args.includes("--skip-to-copilot");
const wordOnly = args.includes("--word-only");

if (!quarter) {
  console.error("Error: --quarter is required (e.g. --quarter Y26Q3)");
  process.exit(1);
}

const ROOT = __dirname;
const TEMP_DIR = path.join(ROOT, "temp");
const FINAL_METRICS = path.join(TEMP_DIR, "final-metrics.md");
const FLEET_INSTRUCTIONS = path.join(ROOT, "gh-cli-prompts", "quarterly-connect-fleet-instructions.txt");
const FLEET_PROMPT_FILE = path.join(TEMP_DIR, "fleet-prompt.txt");

if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });

// ── Generate Word doc from markdown ────────────────────────────────────────
function generateWordDoc(mdContent, outputPath) {
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
          Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType } = docx;

  const children = [];
  const lines = mdContent.split(/\r?\n/);

  // Parse inline markdown: **bold**, *italic*, `code`
  function parseInline(text) {
    const runs = [];
    const regex = /(\*\*(.+?)\*\*|\*(.+?)\*|`(.+?)`)/g;
    let lastIdx = 0;
    let match;
    while ((match = regex.exec(text)) !== null) {
      if (match.index > lastIdx) {
        runs.push(new TextRun(text.slice(lastIdx, match.index)));
      }
      if (match[2]) {
        runs.push(new TextRun({ text: match[2], bold: true }));
      } else if (match[3]) {
        runs.push(new TextRun({ text: match[3], italics: true }));
      } else if (match[4]) {
        runs.push(new TextRun({ text: match[4], font: "Consolas", size: 20 }));
      }
      lastIdx = match.index + match[0].length;
    }
    if (lastIdx < text.length) {
      runs.push(new TextRun(text.slice(lastIdx)));
    }
    return runs;
  }

  // Detect if a line is a Markdown table row: | col1 | col2 | ...
  function isTableRow(line) {
    return /^\|(.+\|)+\s*$/.test(line.trim());
  }

  // Detect separator row: |---|---|  or | :---: | --- |
  function isSeparatorRow(line) {
    return /^\|(\s*:?-+:?\s*\|)+\s*$/.test(line.trim());
  }

  // Parse a table row into cell text values
  function parseCells(line) {
    return line.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map(c => c.trim());
  }

  // Build a native Word table from collected rows
  function buildTable(headerCells, dataRows) {
    const borderStyle = {
      style: BorderStyle.SINGLE,
      size: 1,
      color: "999999",
    };
    const borders = {
      top: borderStyle,
      bottom: borderStyle,
      left: borderStyle,
      right: borderStyle,
    };

    // Header row
    const headerRow = new TableRow({
      tableHeader: true,
      children: headerCells.map(cell => new TableCell({
        borders,
        shading: { type: ShadingType.SOLID, color: "D9E2F3" },
        children: [new Paragraph({ children: [new TextRun({ text: cell, bold: true, size: 20 })] })],
      })),
    });

    // Data rows
    const rows = dataRows.map(cells => new TableRow({
      children: cells.map(cell => new TableCell({
        borders,
        children: [new Paragraph({ children: parseInline(cell) })],
      })),
    }));

    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [headerRow, ...rows],
    });
  }

  let i = 0;
  while (i < lines.length) {
    const line = lines[i];

    // ── Table detection: look ahead for header | separator | data rows ──
    if (isTableRow(line) && i + 1 < lines.length && isSeparatorRow(lines[i + 1])) {
      const headerCells = parseCells(line);
      i += 2; // skip header + separator
      const dataRows = [];
      while (i < lines.length && isTableRow(lines[i]) && !isSeparatorRow(lines[i])) {
        const cells = parseCells(lines[i]);
        // Pad or trim to match header column count
        while (cells.length < headerCells.length) cells.push("");
        dataRows.push(cells.slice(0, headerCells.length));
        i++;
      }
      children.push(buildTable(headerCells, dataRows));
      children.push(new Paragraph({ children: [] })); // spacing after table
      continue;
    }

    // ── Headings ──
    if (line.startsWith("#### ")) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_4, children: parseInline(line.slice(5)) }));
    } else if (line.startsWith("### ")) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_3, children: parseInline(line.slice(4)) }));
    } else if (line.startsWith("## ")) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: parseInline(line.slice(3)) }));
    } else if (line.startsWith("# ")) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: parseInline(line.slice(2)) }));
    } else if (line.startsWith("---")) {
      children.push(new Paragraph({ children: [] }));
    } else if (/^>\s/.test(line)) {
      children.push(new Paragraph({
        indent: { left: 720 },
        children: parseInline(line.replace(/^>\s*/, "")),
      }));
    } else if (/^\s*[-*]\s/.test(line) && !isTableRow(line)) {
      children.push(new Paragraph({
        bullet: { level: 0 },
        children: parseInline(line.replace(/^\s*[-*]\s+/, "")),
      }));
    } else if (/^\s*\d+\.\s/.test(line)) {
      children.push(new Paragraph({
        numbering: { reference: "default-numbering", level: 0 },
        children: parseInline(line.replace(/^\s*\d+\.\s+/, "")),
      }));
    } else if (line.trim() === "") {
      children.push(new Paragraph({ children: [] }));
    } else {
      children.push(new Paragraph({ children: parseInline(line) }));
    }
    i++;
  }

  const doc = new Document({
    numbering: {
      config: [{
        reference: "default-numbering",
        levels: [{ level: 0, format: "decimal", text: "%1.", alignment: AlignmentType.START }],
      }],
    },
    sections: [{ children }],
  });

  return Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(outputPath, buffer);
  });
}

// ── Word-only mode: generate final.docx from existing temp/ files and exit ─
if (wordOnly) {
  const sourcePath = path.join(TEMP_DIR, "Connect-Draft.md");

  if (!fs.existsSync(sourcePath)) {
    console.error(`Error: ${sourcePath} not found. Run the full pipeline first to generate the Connect Draft.`);
    process.exit(1);
  }

  console.log(`Reading Connect Draft from → ${sourcePath}`);
  const mdContent = fs.readFileSync(sourcePath, "utf-8");
  const wordPath = path.join(TEMP_DIR, "final.docx");

  generateWordDoc(mdContent, wordPath).then(() => {
    console.log(`✓ Word document saved → ${wordPath}`);
    process.exit(0);
  }).catch((err) => {
    console.error("Failed to generate Word document:", err.message);
    process.exit(1);
  });
} else {

if (!skipToCopilot) {
// ── Step 1 & 2: Scrape Power BI + Azure OpenAI summarisation ──────────────
if (!skipScrape) {
  console.log("═".repeat(60));
  console.log("STEP 1 — Scraping Power BI report & summarising with AI");
  console.log("═".repeat(60));

  const scrapeArgs = ["scrape-powerbi.js", "--quarter", quarter];
  if (headless) scrapeArgs.push("--headless");

  try {
    execFileSync("node", scrapeArgs, { cwd: ROOT, stdio: "inherit" });
  } catch (err) {
    console.error("\nScraper failed. Fix the issue above and retry, or use --skip-scrape to reuse an existing final-metrics.md.");
    process.exit(1);
  }
} else {
  console.log("Skipping scrape (--skip-scrape). Reusing existing final-metrics.md.");
}

// Verify outputs
if (!fs.existsSync(FINAL_METRICS)) {
  console.error(`\nError: ${FINAL_METRICS} not found. Run without --skip-scrape first.`);
  process.exit(1);
}

// ── Step 3: Merge fleet instructions + metrics into one prompt file ────────
console.log("\n" + "═".repeat(60));
console.log("STEP 2 — Merging fleet instructions + core metrics");
console.log("═".repeat(60));

const instructionsContent = fs.readFileSync(FLEET_INSTRUCTIONS, "utf-8");
const metricsContent = fs.readFileSync(FINAL_METRICS, "utf-8");

// Build the merged file: quarter context + full instruction pack + full metrics.
// This file will be referenced via @fleet-prompt.txt in the Copilot CLI command.
let merged = `Create my quarterly Connect draft using the full instruction pack and core metrics provided below.\n\n`;
merged += `Quarter: ${quarter}\n`;
if (dateRange) {
  merged += `Date range: ${dateRange}\n`;
}
merged += `\n`;
merged += `=== INSTRUCTION PACK ===\n\n`;
merged += instructionsContent.trimEnd() + `\n\n`;
merged += `=== END INSTRUCTION PACK ===\n\n`;
merged += `=== CORE METRICS (${quarter}) ===\n\n`;
merged += metricsContent.trimEnd() + `\n\n`;
merged += `=== END CORE METRICS ===\n\n`;
merged += `=== OUTPUT INSTRUCTIONS ===\n\n`;
merged += `When complete, save the final Connect draft as the file: temp/Connect-Draft.md\n\n`;
merged += `=== END OUTPUT INSTRUCTIONS ===\n`;

fs.writeFileSync(FLEET_PROMPT_FILE, merged, "utf-8");
console.log(`Merged prompt saved → ${FLEET_PROMPT_FILE}`);
console.log(`  Instructions: ${instructionsContent.length} chars`);
console.log(`  Metrics:      ${metricsContent.length} chars`);
console.log(`  Total file:   ${merged.length} chars`);

// ── Step 4: Copy merged prompt to clipboard (fallback) ─────────────────────
console.log("\n" + "═".repeat(60));
console.log("STEP 3 — Copying prompt to clipboard");
console.log("═".repeat(60));

try {
  execFileSync("clip", [], { input: merged, cwd: ROOT });
  console.log("✓ Full merged prompt copied to clipboard.");
} catch {
  console.log("Could not copy to clipboard automatically. Copy the prompt from fleet-prompt.txt.");
}

} // end skipToCopilot

// ── Step 4: Copy setup commands to clipboard and launch Copilot CLI ────────
console.log("\n" + "═".repeat(60));
console.log("STEP 4 — Launching Copilot CLI");
console.log("═".repeat(60));

const setupCommands = `Run the following as individual commands:\n\n/allow-all\nworkiq accepteula\nExecute Prompt: @'${FLEET_PROMPT_FILE}'\n`;

try {
  execFileSync("clip", [], { input: setupCommands, cwd: ROOT });
  console.log("\n✓ Setup commands copied to clipboard.");
  console.log("  Once Copilot opens, paste from clipboard (Ctrl+V) into the prompt.\n");
  console.log("  The clipboard contains:");
  console.log("Run the following as individual commands:\n\n")
  console.log("    /allow-all");
  console.log("    workiq accepteula");
  console.log(`    Execute Prompt: @${FLEET_PROMPT_FILE}'\n`);
} catch {
  console.log("\nCould not copy to clipboard. Run these commands manually in Copilot:");
  console.log("  /allow-all");
  console.log("  workiq accepteula");
  console.log(`  copilot -p "$(Get-Content '${FLEET_PROMPT_FILE}' -Raw)"\n`);
}

console.log("Launching Copilot CLI...\n");

try {
  execFileSync("copilot", [], { cwd: ROOT, stdio: "inherit", shell: true });
} catch {
  // Copilot exited — not necessarily an error
}

console.log("\n" + "═".repeat(60));
console.log("Copilot session ended.");
console.log(`Generate the Word doc with: node run-connect.js --word-only --quarter ${quarter}`);
console.log("═".repeat(60));
} // end else (full pipeline)
