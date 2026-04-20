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

const { execFileSync, execSync, spawn } = require("child_process");
const fs = require("fs");
const os = require("os");
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
// ── Full pipeline continues below ────────────────────────────────────────── ──────────────────────────────────────────────
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
merged += `=== END CORE METRICS ===\n`;

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

// Verify fleet prompt exists before launching Copilot
if (skipToCopilot && !fs.existsSync(FLEET_PROMPT_FILE)) {
  console.error(`\nError: ${FLEET_PROMPT_FILE} not found. Run without --skip-to-copilot first.`);
  process.exit(1);
}

// ── Step 5: Ensure Copilot CLI is authenticated ───────────────────────────
console.log("\n" + "═".repeat(60));
console.log("STEP 4 — Checking Copilot CLI authentication");
console.log("═".repeat(60));

try {
  // A lightweight probe: if not authenticated, copilot -p exits with code 1.
  execSync('copilot -p "ping" --no-auto-update --no-alt-screen -s', {
    cwd: ROOT,
    stdio: ["ignore", "ignore", "ignore"],
    timeout: 60000,
  });
  console.log("✓ Already authenticated with GitHub Copilot.");
} catch {
  console.log("Not logged in — starting Copilot CLI login flow...\n");
  try {
    execFileSync("copilot", ["login"], { cwd: ROOT, stdio: "inherit" });
    console.log("\n✓ Login successful.");
  } catch (loginErr) {
    console.error("Login failed. Please run 'copilot login' manually and retry.");
    process.exit(1);
  }
}

// ── Step 5b: Accept WorkIQ EULA before launching Copilot CLI ───────────────
console.log("\n" + "═".repeat(60));
console.log("STEP 4b — Accepting WorkIQ EULA");
console.log("═".repeat(60));

try {
  execSync("workiq accept-eula", { cwd: ROOT, stdio: "inherit", timeout: 30000 });
  console.log("✓ WorkIQ EULA accepted.");
} catch {
  console.log("⚠ Could not accept WorkIQ EULA (may already be accepted or workiq not installed).");
}

// ── Step 6: Launch Copilot CLI with the merged prompt ──────────────────────
console.log("\n" + "═".repeat(60));
console.log("STEP 5 — Launching GitHub Copilot CLI with fleet prompt");
console.log("═".repeat(60));

const CONNECT_DRAFT_FILE = path.join(TEMP_DIR, "Connect-Draft.md");

const copilot = spawn(
  "powershell",
  ["-NoProfile", "-Command", `copilot -i (Get-Content '${FLEET_PROMPT_FILE}' -Raw)`],
  { cwd: ROOT, stdio: ["inherit", "pipe", "inherit"] }
);

// Capture stdout: tee to console and accumulate for saving
let copilotOutput = "";
copilot.stdout.on("data", (chunk) => {
  process.stdout.write(chunk);
  copilotOutput += chunk.toString();
});

copilot.on("error", (err) => {
  console.error("Failed to launch Copilot CLI. Is it installed? Run: winget install GitHub.Copilot");
  console.error(err.message);
  process.exit(1);
});

copilot.on("close", (code) => {
  console.log(`\nCopilot CLI exited (code ${code}).`);

  // ── Persist the Connect Draft from captured output ──────────────────
  if (copilotOutput.trim().length > 0) {
    fs.writeFileSync(CONNECT_DRAFT_FILE, copilotOutput.trim(), "utf-8");
    console.log(`\n✓ Connect Draft saved → ${CONNECT_DRAFT_FILE}`);
  } else {
    // Fallback: try to find it in Copilot's session workspace
    const sessionStateDir = path.join(os.homedir(), ".copilot", "session-state");
    let draftSrc = null;

    if (fs.existsSync(sessionStateDir)) {
      const sessions = fs.readdirSync(sessionStateDir);
      let latestTime = 0;
      for (const session of sessions) {
        const filesDir = path.join(sessionStateDir, session, "files");
        if (!fs.existsSync(filesDir)) continue;
        for (const file of fs.readdirSync(filesDir)) {
          if (file.endsWith("-Connect-Draft.md")) {
            const fullPath = path.join(filesDir, file);
            const mtime = fs.statSync(fullPath).mtimeMs;
            if (mtime > latestTime) {
              latestTime = mtime;
              draftSrc = fullPath;
            }
          }
        }
      }
    }

    if (draftSrc) {
      fs.copyFileSync(draftSrc, CONNECT_DRAFT_FILE);
      console.log(`\n✓ Connect Draft copied → ${CONNECT_DRAFT_FILE}`);
    } else {
      console.log("\n⚠ Could not capture or find a Connect Draft.");
    }
  }

  // ── Generate Word document from the Connect Draft ──────────────────
  if (fs.existsSync(CONNECT_DRAFT_FILE)) {
    const mdContent = fs.readFileSync(CONNECT_DRAFT_FILE, "utf-8");
    const wordPath = path.join(TEMP_DIR, "final.docx");
    generateWordDoc(mdContent, wordPath).then(() => {
      console.log(`✓ Word document saved → ${wordPath}`);
    }).catch((docErr) => {
      console.error("⚠ Failed to generate Word document:", docErr.message);
    });
  }

  // ── ASCII art finish ─────────────────────────────────────────────────
  console.log(`\n╔══════════════════════════════════════════════════════════════════╗`);
  console.log(`  ║  ★  C O M P L E T E  ★  Find your final output in temp/       ║`);
  console.log(`  ╚══════════════════════════════════════════════════════════════════╝\n`);
});
} // end else (full pipeline)
