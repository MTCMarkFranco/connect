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
 */

const { execFileSync, execSync, spawn } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");

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

// ── Step 6: Launch Copilot CLI with the merged prompt ──────────────────────
console.log("\n" + "═".repeat(60));
console.log("STEP 5 — Launching GitHub Copilot CLI with fleet prompt");
console.log("═".repeat(60));

const copilot = spawn(
  "powershell",
  ["-NoProfile", "-Command", `copilot -i (Get-Content '${FLEET_PROMPT_FILE}' -Raw)`],
  { cwd: ROOT, stdio: "inherit" }
);

copilot.on("error", (err) => {
  console.error("Failed to launch Copilot CLI. Is it installed? Run: winget install GitHub.Copilot");
  console.error(err.message);
  process.exit(1);
});

copilot.on("close", (code) => {
  console.log(`\nCopilot CLI exited (code ${code}).`);

  // ── Copy the generated Connect Draft from Copilot's session workspace ──
  const sessionStateDir = path.join(os.homedir(), ".copilot", "session-state");
  let draftSrc = null;

  if (fs.existsSync(sessionStateDir)) {
    // Find the most recently modified *Connect-Draft.md across all sessions
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
    const draftDest = path.join(TEMP_DIR, path.basename(draftSrc));
    fs.copyFileSync(draftSrc, draftDest);
    console.log(`\n✓ Connect Draft copied → ${draftDest}`);
  } else {
    console.log("\n⚠ Could not find a Connect Draft in Copilot session workspace.");
  }

  // ── ASCII art finish ─────────────────────────────────────────────────
  console.log(`
  ╔══════════════════════════════════════════════════════════╗
  ║                                                          ║
  ║     ██████╗ ██████╗ ███╗   ███╗██████╗ ██╗     ███████╗ ║
  ║    ██╔════╝██╔═══██╗████╗ ████║██╔══██╗██║     ██╔════╝ ║
  ║    ██║     ██║   ██║██╔████╔██║██████╔╝██║     █████╗   ║
  ║    ██║     ██║   ██║██║╚██╔╝██║██╔═══╝ ██║     ██╔══╝   ║
  ║    ╚██████╗╚██████╔╝██║ ╚═╝ ██║██║     ███████╗███████╗ ║
  ║     ╚═════╝ ╚═════╝ ╚═╝     ╚═╝╚═╝     ╚══════╝╚══════╝ ║
  ║                                                          ║
  ║    ████████╗███████╗                                     ║
  ║    ╚══██╔══╝██╔════╝                                     ║
  ║       ██║   █████╗                                       ║
  ║       ██║   ██╔══╝                                       ║
  ║       ██║   ███████╗                                     ║
  ║       ╚═╝   ╚══════╝                                     ║
  ║                                                          ║
  ║    Find your final output in the temp/ directory         ║
  ║                                                          ║
  ╚══════════════════════════════════════════════════════════╝
`);
});
