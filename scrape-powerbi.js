/**
 * scrape-powerbi.js
 *
 * Uses Playwright with Microsoft Edge to scrape a Power BI report page.
 * Authentication is handled interactively on the first run; the browser
 * profile is persisted so subsequent runs reuse the session.
 *
 * Usage:
 *   node scrape-powerbi.js                  # headed (visible browser)
 *   node scrape-powerbi.js --headless       # headless after auth is cached
 *   node scrape-powerbi.js --quarter FY26Q3 # tag the output with a quarter
 */

const { chromium } = require("playwright");
const { AzureOpenAI } = require("openai");
const { DefaultAzureCredential, getBearerTokenProvider } = require("@azure/identity");
const fs = require("fs");
const path = require("path");

require("dotenv").config({ path: path.join(__dirname, ".env") });

// ── Configuration ──────────────────────────────────────────────────────────
const REPORT_URL =
  "https://msit.powerbi.com/groups/me/apps/bc16e81a-2071-4f5f-8131-c2e9b7211346/reports/99396eba-3a62-499c-bd28-2165ac0a0737/ReportSectionc481b05a185bb8cce548?experience=power-bi";

const AUTH_STATE_DIR = path.join(__dirname, ".auth");
const TEMP_DIR = path.join(__dirname, "temp");
const RAW_TEXT_FILE = path.join(TEMP_DIR, "core-metrics.txt");
const SCREENSHOT_FILE = path.join(TEMP_DIR, "core-metrics.png");
const FINAL_OUTPUT_FILE = path.join(TEMP_DIR, "final-metrics.md");

// How long to wait for the report visuals to finish rendering (ms)
const RENDER_WAIT_MS = 15_000;
// Max time to wait for initial page load (ms)
const PAGE_LOAD_TIMEOUT_MS = 120_000;

// ── CLI args ───────────────────────────────────────────────────────────────
const args = process.argv.slice(2);
const headless = args.includes("--headless");
const quarterIdx = args.indexOf("--quarter");
const quarter = quarterIdx !== -1 ? args[quarterIdx + 1] : null;

// ── Helpers ────────────────────────────────────────────────────────────────
function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

// ── Main ───────────────────────────────────────────────────────────────────
(async () => {
  ensureDir(AUTH_STATE_DIR);
  ensureDir(TEMP_DIR);

  console.log("Launching Microsoft Edge…");
  console.log(headless ? "  Mode: headless" : "  Mode: headed (visible browser)");

  // Persistent context keeps cookies/sessions across runs so you only
  // authenticate interactively once.
  const context = await chromium.launchPersistentContext(AUTH_STATE_DIR, {
    channel: "msedge",
    headless,
    viewport: { width: 1920, height: 1080 },
    args: ["--start-maximized"],
    // Generous timeout for AAD auth redirects
    timeout: PAGE_LOAD_TIMEOUT_MS,
  });

  const page = context.pages()[0] || (await context.newPage());

  // ── Navigate & authenticate ──────────────────────────────────────────
  console.log("Navigating to Power BI report…");
  console.log("  If a login prompt appears, complete authentication in the browser window.");

  try {
    await page.goto(REPORT_URL, {
      waitUntil: "domcontentloaded",
      timeout: PAGE_LOAD_TIMEOUT_MS,
    });
  } catch (err) {
    // On the very first run AAD may redirect multiple times; retry once.
    console.log("  Initial navigation timed out (likely auth redirects). Retrying…");
    await page.goto(REPORT_URL, {
      waitUntil: "domcontentloaded",
      timeout: PAGE_LOAD_TIMEOUT_MS,
    });
  }

  // Wait for the user to finish any interactive sign-in.
  // We detect a successful load by waiting for a Power BI–specific element.
  console.log("Waiting for Power BI report to render…");
  try {
    await page.waitForSelector(
      'div[class*="visual"], exploration-container, .reportCanvas, div[aria-label*="report"], .visualContainerHost',
      { timeout: PAGE_LOAD_TIMEOUT_MS }
    );
  } catch {
    console.log("  Could not detect report visuals via selector. Waiting extra time…");
  }

  // Extra settling time for all visuals to finish loading data.
  console.log(`Waiting ${RENDER_WAIT_MS / 1000}s for visuals to finish rendering…`);
  await page.waitForTimeout(RENDER_WAIT_MS);

  // ── Screenshot by scrolling through the page ─────────────────────────
  const screenshotFiles = [];

  // Power BI renders inside an internal scrollable container, not the document body.
  // Find the scrollable element that holds the report content.
  const viewportHeight = page.viewportSize().height;

  const scrollInfo = await page.evaluate(() => {
    // Common Power BI scrollable containers (ordered by likelihood)
    const selectors = [
      '.canvasFlexBox',
      '.visualContainerHost',
      'div[class*="scroll"]',
      'exploration-container',
      '.reportCanvas',
      '.canvas',
    ];

    // Also try: find the tallest scrollable div on the page
    function findScrollableContainer() {
      for (const sel of selectors) {
        const el = document.querySelector(sel);
        if (el && el.scrollHeight > el.clientHeight + 50) {
          return el;
        }
      }
      // Fallback: find the element with the largest scrollHeight
      const allDivs = document.querySelectorAll('div');
      let best = null;
      let bestScroll = 0;
      for (const div of allDivs) {
        if (div.scrollHeight > div.clientHeight + 50 && div.scrollHeight > bestScroll) {
          bestScroll = div.scrollHeight;
          best = div;
        }
      }
      return best;
    }

    const container = findScrollableContainer();
    if (container) {
      // Tag it so we can find it later
      container.setAttribute('data-scrape-scroll', 'true');
      return {
        found: true,
        scrollHeight: container.scrollHeight,
        clientHeight: container.clientHeight,
      };
    }

    // No internal scroller — fall back to document scroll
    return {
      found: false,
      scrollHeight: document.documentElement.scrollHeight,
      clientHeight: document.documentElement.clientHeight,
    };
  });

  const scrollAmount = Math.max(scrollInfo.clientHeight, viewportHeight) - 100; // overlap 100px
  const totalScreenshots = Math.max(1, Math.ceil(scrollInfo.scrollHeight / scrollAmount));

  console.log(`Scroll container: ${scrollInfo.found ? "internal Power BI element" : "document body"}`);
  console.log(`Content height: ${scrollInfo.scrollHeight}px, step: ${scrollAmount}px → ${totalScreenshots} screenshot(s)`);

  for (let i = 0; i < totalScreenshots; i++) {
    const scrollY = i * scrollAmount;
    await page.evaluate(({ y, hasContainer }) => {
      if (hasContainer) {
        const el = document.querySelector('[data-scrape-scroll="true"]');
        if (el) { el.scrollTop = y; return; }
      }
      window.scrollTo(0, y);
    }, { y: scrollY, hasContainer: scrollInfo.found });

    // Pause for lazy-loaded visuals to render after scroll
    await page.waitForTimeout(2000);

    const screenshotPath = path.join(TEMP_DIR, `core-metrics-${i + 1}.png`);
    await page.screenshot({ path: screenshotPath });
    screenshotFiles.push(screenshotPath);
    console.log(`  Screenshot ${i + 1}/${totalScreenshots} saved → ${screenshotPath}`);
  }

  // Also save the first screenshot with the legacy filename for backwards compat
  if (screenshotFiles.length > 0) {
    fs.copyFileSync(screenshotFiles[0], SCREENSHOT_FILE);
  }
  console.log(`${screenshotFiles.length} screenshot(s) captured.`);

  // ── Extract all visible text from the page ───────────────────────────
  console.log("Extracting page text…");
  const bodyText = await page.evaluate(() => {
    return document.body.innerText || "";
  });

  // ── Write output ─────────────────────────────────────────────────────
  let output = "";
  if (quarter) {
    output += `QUARTERLY CORE METRICS — ${quarter}\n`;
  } else {
    output += "QUARTERLY CORE METRICS\n";
  }
  output += `Extracted: ${new Date().toISOString()}\n`;
  output += `Source: Power BI report (automated Playwright scrape)\n`;
  output += "=".repeat(60) + "\n\n";
  output += bodyText + "\n";

  fs.writeFileSync(RAW_TEXT_FILE, output, "utf-8");
  console.log(`\nRaw metrics saved → ${RAW_TEXT_FILE}`);

  await context.close();

  // ── Send to Azure OpenAI for summarisation ───────────────────────────
  const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
  const deployment = process.env.AZURE_OPENAI_DEPLOYMENT || "gpt-4o-mini";

  if (!endpoint ) {
    console.log("\nAZURE_OPENAI_ENDPOINT not set in .env — skipping AI summarisation.");
    console.log("Done.");
    return;
  }

  console.log(`\nCalling Azure OpenAI (${deployment}) to summarise metrics…`);

  // Encode all screenshots as base64 for the multi-modal prompt
  const imageContentParts = screenshotFiles.map((filePath, idx) => {
    const base64 = fs.readFileSync(filePath).toString("base64");
    return {
      type: "image_url",
      image_url: { url: `data:image/png;base64,${base64}`, detail: "high" },
    };
  });
  console.log(`  Attaching ${imageContentParts.length} screenshot(s) to the prompt.`);

  const credential = new DefaultAzureCredential();
  const azureADTokenProvider = getBearerTokenProvider(
    credential,
    "https://cognitiveservices.azure.com/.default"
  );

  const client = new AzureOpenAI({
    endpoint,
    azureADTokenProvider,
    apiVersion: "2024-10-21",
    deployment,
  });

  const completion = await client.chat.completions.create({
    model: deployment,
    messages: [
      {
        role: "system",
        content:
          "You are a business analyst. The user will provide raw text and one or more screenshot images from a Power BI quarterly performance report. " +
          "Use both the text and all the images to extract every visible metric, KPI, and data point. " +
          "Organise them into a clean, well-structured **Markdown** summary grouped by theme " +
          "(e.g. Customer Satisfaction, Engagements, Outcomes, Solution Areas, Journeys, Opportunities). " +
          "Use Markdown headings (##, ###), bullet lists, bold for metric names, and tables where appropriate. " +
          "Keep every number exact — do not round or invent data. Use concise business-ready prose.",
      },
      {
        role: "user",
        content: [
          {
            type: "text",
            text: `Here is the raw text extracted from the Power BI report${quarter ? ` for ${quarter}` : ""}:\n\n${bodyText}`,
          },
          ...imageContentParts,
        ],
      },
    ],
    temperature: 0.2,
    max_completion_tokens: 4096,
  });

  const summary = completion.choices[0].message.content;

  // Write the final Markdown summary to final-metrics.md
  let mdOutput = "";
  if (quarter) {
    mdOutput += `# Quarterly Core Metrics — ${quarter}\n\n`;
  } else {
    mdOutput += "# Quarterly Core Metrics\n\n";
  }
  mdOutput += `> **Extracted:** ${new Date().toISOString()}  \n`;
  mdOutput += `> **Source:** Power BI report → summarised by Azure OpenAI (${deployment})\n\n`;
  mdOutput += "---\n\n";
  mdOutput += summary + "\n";

  fs.writeFileSync(FINAL_OUTPUT_FILE, mdOutput, "utf-8");
  console.log(`\nSummarised metrics saved → ${FINAL_OUTPUT_FILE}`);
  console.log("Done.");
})();
