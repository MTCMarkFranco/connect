/**
 * baby-carlos.js
 *
 * Single orchestration script that runs all steps end-to-end:
 *   1. Scrape Power BI report (Playwright + Edge)
 *   2. Summarise metrics via Azure OpenAI (multi-modal)
 *   3. Merge fleet instructions + metrics into one self-contained /fleet prompt
 *   4. Copy the merged prompt to clipboard
 *   5. Ensure Copilot CLI is authenticated (auto-login if needed)
 *   6. Launch GitHub Copilot CLI interactively
 *   7. Iteratively refine the Connect draft against the measuring-stick rubric
 *      until every dimension reaches "Exceptional impact"
 *
 * Usage:
 *   node baby-carlos.js --quarter FY26Q3
 *   node baby-carlos.js --quarter FY26Q3 --headless
 *   node baby-carlos.js --quarter FY26Q3 --headless --date-range "Jan 1, 2026 - Mar 31, 2026"
 *   node baby-carlos.js --skip-scrape --quarter FY26Q3   # reuse existing final-metrics.md
 *   node baby-carlos.js --skip-to-copilot --quarter FY26Q3 # jump straight to Copilot CLI
 *   node baby-carlos.js --word-only --quarter FY26Q3       # generate final.docx from existing temp/ files
 *   node baby-carlos.js --refine-only --quarter FY26Q3     # run only the measuring-stick refinement loop
 *   node baby-carlos.js --skip-refine --quarter FY26Q3     # skip the refinement loop
 *   node baby-carlos.js --max-refine-passes 5 --quarter FY26Q3  # set max refinement iterations (default 3)
 *   node baby-carlos.js --target-score 10 --quarter FY26Q3      # set target Exceptional cells (default 10 of 12)
 */

const { execFileSync, execSync } = require("child_process");
const fs = require("fs");
const path = require("path");
const readline = require("readline");
const crypto = require("crypto");
const docx = require("docx");
const { AzureOpenAI } = require("openai");
const { DefaultAzureCredential, getBearerTokenProvider } = require("@azure/identity");

require("dotenv").config({ path: path.join(__dirname, ".env") });

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
const refineOnly = args.includes("--refine-only");
const skipRefine = args.includes("--skip-refine");
const maxRefinePasses = parseInt(getArg("--max-refine-passes") || "3", 10);
const noClarify = args.includes("--no-clarify");
const mergeOnly = args.includes("--merge-only");
const mergePass = parseInt(getArg("--merge-pass") || "1", 10);
const workiqMaxConcurrency = Math.max(1, parseInt(getArg("--workiq-max-concurrency") || "4", 10));
const workiqBatchSize = Math.max(1, parseInt(getArg("--workiq-batch-size") || "2", 10));
const workiqJitterMinMs = Math.max(0, parseInt(getArg("--workiq-jitter-min-ms") || "1200", 10));
const workiqJitterMaxMs = Math.max(workiqJitterMinMs, parseInt(getArg("--workiq-jitter-max-ms") || "6000", 10));
const workiqRetries = Math.max(0, parseInt(getArg("--workiq-retries") || "2", 10));
const workiqRetryBackoffMs = Math.max(0, parseInt(getArg("--workiq-retry-backoff-ms") || "4000", 10));

if (!quarter) {
  console.error("Error: --quarter is required (e.g. --quarter Y26Q3)");
  process.exit(1);
}

const ROOT = __dirname;
const TEMP_DIR = path.join(ROOT, "temp");
const FINAL_METRICS = path.join(TEMP_DIR, "final-metrics.md");
const FLEET_INSTRUCTIONS = path.join(ROOT, "gh-cli-prompts", "quarterly-connect-fleet-instructions.txt");
const FLEET_PROMPT_FILE = path.join(TEMP_DIR, "fleet-prompt.txt");
const MEASURING_STICK = path.join(ROOT, "guidance", "measuring-stick.md");
const MORE_EVIDENCE_DIR = path.join(ROOT, "more-evidence");
const USER_CONTEXT_FILE = path.join(TEMP_DIR, "user-context.txt");

if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });

// ── Measuring-stick refinement loop (WorkIQ-powered) ───────────────────────
//
// HOW IT WORKS:
//
// 1. EVALUATE: Send the current Connect draft + measuring stick rubric to
//    Azure OpenAI. The model scores every cell in the 3-circles × 4-dimensions
//    grid as "Exceptional", "Successful", or "Lower". It returns structured
//    JSON with the rating and, for any non-Exceptional cell, a concrete
//    improvement instruction explaining WHAT is missing and HOW to fix it.
//
// 2. CHECK: If the target score (default 10/12) is met, stop.
//
// 3. GENERATE WORKIQ PROMPTS: For each non-Exceptional cell, generate a
//    targeted WorkIQ search prompt. Each prompt focuses on finding evidence
//    of HOW the author made impact — not just what was delivered, but the
//    specific actions, decisions, and behaviours that drove the outcome.
//    Prompts are derived from the evaluation's reasoning + improvement
//    guidance + the measuring stick's Exceptional-tier language.
//
// 4. LAUNCH COPILOT CLI WITH /FLEET: Build a single fleet prompt that runs
//    all WorkIQ searches in parallel across emails, Teams, documents, and
//    Loop. The prompt instructs Copilot to save consolidated evidence to
//    temp/workiq-evidence-pass-{N}.md.
//
// 5. MERGE EVIDENCE: After Copilot CLI exits, read the gathered evidence
//    and use Azure OpenAI to weave it into the Connect draft. The merge
//    focuses on HOW — the specific actions, decisions, and behaviors that
//    demonstrate Exceptional-tier impact. New evidence is integrated
//    naturally; no existing Exceptional content is weakened.
//
// 6. RE-EVALUATE: Save the updated draft as Connect-Draft-v{N}.md and
//    repeat from step 1. The loop runs for at most --max-refine-passes
//    iterations (default 3).
//
// 7. PROSE CONVERSION: After all refinement passes, convert the best-scoring
//    evidence-ledger draft into polished first-person narrative prose. The
//    prose version is re-evaluated to ensure the score holds; if it drops,
//    the conversion is retried with feedback about what was lost.
//
// 8. FINALIZE: The ledger version is saved as Connect-Draft-ledger.md (for
//    reference) and the prose version overwrites Connect-Draft.md. The Word
//    doc is generated from the prose version.

const TARGET_EXCEPTIONAL_CELLS = parseInt(getArg("--target-score") || "10", 10);

async function createAzureOpenAIClient() {
  const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
  const deployment = process.env.AZURE_OPENAI_DEPLOYMENT || "gpt-4o-mini";
  if (!endpoint) {
    throw new Error("AZURE_OPENAI_ENDPOINT is not set in .env");
  }
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
  return { client, deployment };
}

const EVALUATION_SYSTEM_PROMPT = `You are a performance-review evaluator. You will receive:
1. A "measuring stick" rubric that defines three impact tiers — Exceptional, Successful, and Lower — across four dimensions (Business Acumen, Citizenship, Technical Excellence, Role Excellence) and three circles (Individual Accomplishments, Building on Others, Contributing to Others' Success).
2. A quarterly Connect draft written by the employee.

Your job is to evaluate the Connect draft against every cell in the 3×4 grid and return a JSON object.

EVALUATION RULES:
- Rate each cell as "Exceptional", "Successful", or "Lower" based on whether the draft content demonstrates the specific behaviours described in that tier.
- A cell is "Exceptional" ONLY if the draft contains concrete, evidence-backed content that clearly maps to the Exceptional tier description for that dimension and circle. Pay special attention to whether the draft explains HOW the impact was achieved — the specific actions, decisions, methods, and behaviours — not just WHAT was delivered.
- If a cell is not Exceptional, you MUST provide a specific, actionable improvement instruction that explains:
  (a) WHAT is missing or insufficient — reference the exact rubric language the draft fails to satisfy.
  (b) HOW to fix it — describe what kind of evidence, framing, or "how" narrative would elevate it.
  (c) SEARCH GUIDANCE — suggest specific WorkIQ search terms (for emails, Teams, documents, Loop) that could surface the missing evidence about HOW impact was made.

Return ONLY valid JSON in this exact schema (no markdown fencing, no commentary):
{
  "allExceptional": true/false,
  "exceptionalCount": <number>,
  "cells": [
    {
      "circle": "Individual Accomplishments",
      "dimension": "Business Acumen",
      "rating": "Exceptional",
      "reasoning": "Brief explanation of why this rating was given",
      "improvement": null,
      "searchTerms": null
    },
    {
      "circle": "Individual Accomplishments",
      "dimension": "Citizenship",
      "rating": "Successful",
      "reasoning": "...",
      "improvement": "Specific instruction on what to change and how",
      "searchTerms": ["term1", "term2", "term3"]
    }
  ]
}

You must include all 12 cells (3 circles × 4 dimensions). Do not omit any.`;

async function evaluateDraft(client, deployment, draftContent, rubricContent) {
  const response = await client.chat.completions.create({
    model: deployment,
    messages: [
      { role: "system", content: EVALUATION_SYSTEM_PROMPT },
      {
        role: "user",
        content: `=== MEASURING STICK RUBRIC ===\n\n${rubricContent}\n\n=== END RUBRIC ===\n\n=== CONNECT DRAFT ===\n\n${draftContent}\n\n=== END DRAFT ===`,
      },
    ],
    temperature: 0.1,
    max_completion_tokens: 4096,
    response_format: { type: "json_object" },
  });

  const raw = response.choices[0].message.content;
  return JSON.parse(raw);
}

function printEvaluationSummary(evaluation, passLabel) {
  console.log(`\n  Pass ${passLabel} evaluation results:`);
  console.log("  " + "─".repeat(56));

  const circles = [
    "Individual Accomplishments",
    "Building on Others",
    "Contributing to Others' Success",
  ];
  const dimensions = ["Business Acumen", "Citizenship", "Technical Excellence", "Role Excellence"];

  for (const circle of circles) {
    console.log(`\n  ${circle}:`);
    for (const dim of dimensions) {
      const cell = evaluation.cells.find(
        (c) => c.circle === circle && c.dimension === dim
      );
      if (cell) {
        const icon = cell.rating === "Exceptional" ? "★" : cell.rating === "Successful" ? "●" : "○";
        console.log(`    ${icon} ${dim}: ${cell.rating}`);
      }
    }
  }

  const exceptional = evaluation.cells.filter((c) => c.rating === "Exceptional").length;
  console.log(`\n  Score: ${exceptional}/12 cells at Exceptional impact`);
  return exceptional;
}

// ── Interactive search-term clarification ──────────────────────────────────

function askUser(question) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => {
    rl.question(question, (answer) => {
      rl.close();
      resolve(answer.trim());
    });
  });
}

// Generic terms that are too vague for WorkIQ searches
const GENERIC_TERMS = new Set([
  "community engagement", "team recognition", "leadership initiatives",
  "cross-hub solutions", "organizational improvements", "collaboration outcomes",
  "hub experiences", "customer interactions", "collaborative projects",
  "team performance", "collaborative success", "sales growth",
  "innovative solutions", "stakeholder needs", "business efficiency",
  "mentoring impact", "team support", "challenge handling",
  "organizational change", "technical guidance", "product improvement",
  "reusable work", "peer coaching", "effectiveness improvement",
  "team motivation", "cultural shifts", "collaboration improvements",
]);

function isAmbiguous(terms) {
  if (!terms || terms.length === 0) return true;
  return terms.every((t) => GENERIC_TERMS.has(t.toLowerCase().trim()));
}

async function clarifySearchTerms(gaps) {
  console.log("\n  ── Clarifying ambiguous search terms ──");
  console.log("  For each gap with vague search terms, provide specific details.");
  console.log("  Examples: project names, customer names, colleague names, technologies,");
  console.log("  email subjects, Teams channel names, specific initiatives, etc.");
  console.log("  Press Enter to skip if you have nothing to add.\n");

  let clarified = 0;
  for (const gap of gaps) {
    const terms = gap.searchTerms || [];
    if (!isAmbiguous(terms)) continue;

    console.log(`  [${ gap.circle } → ${ gap.dimension }] (${ gap.rating })`);
    console.log(`    Current search terms: ${ terms.join(", ") || "(none)" }`);
    console.log(`    What the rubric wants: ${ gap.improvement.substring(0, 200) }`);
    const answer = await askUser("    → Your specific examples/names/projects: ");
    if (answer) {
      gap.userContext = answer;
      clarified++;
    }
    console.log();
  }

  if (clarified > 0) {
    console.log(`  ✓ Added user context for ${clarified} gap(s).\n`);
  } else {
    console.log("  No additional context provided.\n");
  }
  return gaps;
}

// ── WorkIQ prompt generation ───────────────────────────────────────────────

/**
 * Build short, specific WorkIQ queries from a gap cell.
 * WorkIQ times out on long narrative queries, so keep each query to ONE
 * concise topic (3-5 words) plus the date hint.
 *
 * Slot allocation:
 *   - Up to 3 cell-specific searchTerms from the evaluation
 *   - Up to 2 user-context hints from user-context.txt (always included)
 *   - Dimension-based fallback if nothing else exists
 *
 * Returns up to 5 short query strings per cell.
 */
function buildShortQueries(gap, quarter) {
  const CELL_SLOTS = 3;
  const USER_SLOTS = 2;
  const queries = [];
  const cellTerms = (gap.searchTerms || []).filter(Boolean);
  const userTerms = gap.userContext ? gap.userContext.split(/,\s*/).filter(Boolean) : [];

  // Quarter date range for context
  const dateHint = quarter.replace("FY26Q3", "January–March 2026")
    .replace("FY26Q4", "April–June 2026")
    .replace("FY26Q2", "October–December 2025")
    .replace("FY26Q1", "July–September 2025");

  // 1. Cell-specific terms (capped at CELL_SLOTS)
  for (const term of cellTerms) {
    if (queries.length >= CELL_SLOTS) break;
    queries.push(`${term} ${dateHint}`);
  }

  // 2. User-context hints (guaranteed USER_SLOTS, skip duplicates)
  const seen = new Set(queries.map((q) => q.toLowerCase()));
  let userAdded = 0;
  for (const hint of userTerms) {
    if (userAdded >= USER_SLOTS) break;
    const candidate = `${hint} ${dateHint}`;
    if (!seen.has(candidate.toLowerCase())) {
      queries.push(candidate);
      seen.add(candidate.toLowerCase());
      userAdded++;
    }
  }

  // 3. Fallback: dimension-based short query
  if (queries.length === 0) {
    const dimQueries = {
      "Business Acumen": [`customer impact and Azure pipeline ${dateHint}`],
      "Citizenship": [`community contributions and recognition ${dateHint}`],
      "Technical Excellence": [`technical architecture and solution design ${dateHint}`],
      "Role Excellence": [`coaching mentoring and team leadership ${dateHint}`],
    };
    queries.push(...(dimQueries[gap.dimension] || [`work accomplishments ${dateHint}`]));
  }

  return queries;
}

function generateWorkIQFleetPrompt(gaps, passNumber, quarter, options = {}) {
  const parallelLimit = options.parallelLimit || workiqMaxConcurrency;
  const batchLabel = options.batchLabel ? ` (${options.batchLabel})` : "";

  // Build workstreams with pre-built short queries
  const workstreams = gaps.map((gap, idx) => {
    const queries = buildShortQueries(gap, quarter);
    const queryBlock = queries.map((q, qi) => `  QUERY ${qi + 1}: "${q}"`).join("\n");

    return `
WORKSTREAM ${idx + 1}: [${gap.circle} → ${gap.dimension}]
${queryBlock}
GOAL: Find evidence of HOW impact was made for this dimension.`;
  });

  const evidenceFile = path.join(TEMP_DIR, `workiq-evidence-pass-${passNumber}.md`).replace(/\\/g, "/");

  return `Search WorkIQ to find evidence for a quarterly Connect (${quarter})${batchLabel}.

IMPORTANT — QUERY RULES:
- Use ONLY the exact queries listed below. Do NOT rewrite, expand, or combine them.
- Each query must be sent to WorkIQ as-is — short queries work, long ones time out.
- Do NOT include the person's name in queries (WorkIQ already scopes to the current user).
- Run workstreams in parallel, but cap active workstreams at ${parallelLimit}.
- Exclude personal/private life content completely. Ignore items such as banking, mortgage renewal, personal finance, family, health, travel, social plans, or non-work admin.
- Include only work-relevant evidence tied to professional impact, delivery, collaboration, leadership, coaching, or community contributions.

IMPORTANT — OUTPUT RULES:
- Write evidence to the output file INCREMENTALLY — after completing each workstream, immediately append that workstream's evidence to the file. Do NOT wait until all workstreams are done.
- Keep each evidence item CONCISE — 2-3 sentences max per field. Do NOT reproduce full WorkIQ response text.
- If a WorkIQ call times out, retry up to 3 times with a shorter query:
  Retry 1: rephrase under 10 words, drop qualifiers.
  Retry 2: use only the single most important keyword or phrase.
  Retry 3: use the broadest single word (e.g. "events", "kudos").
  If all retries fail, log the gap and move on.

${workstreams.join("\n")}

For each result, note: Source type, reference, date, people involved, and HOW impact was made.
If a result is personal/non-work, discard it and do not include it in the evidence file.

Save all evidence to: ${evidenceFile}

Format:

# WorkIQ Evidence Gathered — Pass ${passNumber}

## [Circle → Dimension] (Workstream N)
### Evidence Item 1
- **Source:** [type] — [reference]
- **Date:** [date]
- **People/Orgs:** [names]
- **HOW (the action/decision/behaviour):** [description]
- **Why it matters:** [rubric alignment]
`;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function randomIntInclusive(min, max) {
  if (max <= min) return min;
  return crypto.randomInt(min, max + 1);
}

function getWorkIQJitterMs() {
  return randomIntInclusive(workiqJitterMinMs, workiqJitterMaxMs);
}

function chunkArray(items, size) {
  const chunks = [];
  for (let i = 0; i < items.length; i += size) {
    chunks.push(items.slice(i, i + size));
  }
  return chunks;
}

function looksLikeMcpTimeout(text) {
  if (!text) return false;
  return /MCP error -32001|Request timed out|timed out/i.test(text);
}

function runCopilotEvidenceCapture(fleetPromptPath, evidenceFilePath) {
  try {
    const promptContent = fs.readFileSync(fleetPromptPath, "utf-8");
    const copilotLoader = path.join(process.env.APPDATA, "npm", "node_modules", "@github", "copilot", "npm-loader.js");
    const copilotOutput = execFileSync("node", [copilotLoader, "--allow-all", "-i", promptContent], {
      cwd: ROOT, encoding: "utf-8", timeout: 1200000, maxBuffer: 10 * 1024 * 1024,
    });

    if (fs.existsSync(evidenceFilePath) && fs.statSync(evidenceFilePath).size > 500) {
      return {
        ok: true,
        fromFile: true,
        timedOut: false,
        output: copilotOutput || "",
        bytes: fs.statSync(evidenceFilePath).size,
      };
    }

    fs.writeFileSync(evidenceFilePath, copilotOutput, "utf-8");
    return {
      ok: true,
      fromFile: false,
      timedOut: looksLikeMcpTimeout(copilotOutput),
      output: copilotOutput || "",
      bytes: (copilotOutput || "").length,
    };
  } catch (err) {
    if (fs.existsSync(evidenceFilePath) && fs.statSync(evidenceFilePath).size > 500) {
      const content = fs.readFileSync(evidenceFilePath, "utf-8");
      return {
        ok: true,
        fromFile: true,
        timedOut: looksLikeMcpTimeout(content) || looksLikeMcpTimeout(err.stdout || "") || looksLikeMcpTimeout(err.message || ""),
        output: content,
        bytes: fs.statSync(evidenceFilePath).size,
      };
    }

    if (err.stdout) {
      fs.writeFileSync(evidenceFilePath, err.stdout, "utf-8");
      return {
        ok: false,
        fromFile: false,
        timedOut: looksLikeMcpTimeout(err.stdout) || looksLikeMcpTimeout(err.message || ""),
        output: err.stdout,
        bytes: err.stdout.length,
      };
    }

    return {
      ok: false,
      fromFile: false,
      timedOut: looksLikeMcpTimeout(err.message || ""),
      output: "",
      bytes: 0,
      message: err.message,
    };
  }
}

// ── Evidence merge ─────────────────────────────────────────────────────────

// ── Per-cell merge helpers ──────────────────────────────────────────────────

/**
 * Parse the evidence markdown into per-cell blocks.
 * Returns a Map: "Circle → Dimension" => evidence text
 */
function parseEvidenceByCells(evidenceContent) {
  const cellMap = new Map();
  const cellHeaderRe = /^## \[(.+?)\]\s*\(Workstream \d+\)/gm;
  const headers = [];
  let m;
  while ((m = cellHeaderRe.exec(evidenceContent)) !== null) {
    headers.push({ key: m[1].trim(), index: m.index });
  }
  for (let i = 0; i < headers.length; i++) {
    const start = headers[i].index;
    const end = i + 1 < headers.length ? headers[i + 1].index : evidenceContent.length;
    cellMap.set(headers[i].key, evidenceContent.slice(start, end).trim());
  }
  return cellMap;
}

/**
 * Determine which draft section a cell maps to.
 * Returns the section header string to insert after.
 */
function cellToDraftSection(circle, dimension) {
  // Map cells to the draft's Part 4 sections
  const sectionMap = {
    "Individual Accomplishments": {
      "Business Acumen": "### A. Reflect on the Past — Results Delivered and How",
      "Citizenship": "**Community Impact and Technical Leadership**",
      "Technical Excellence": "### A. Reflect on the Past — Results Delivered and How",
      "Role Excellence": "**Coaching and Mentoring**",
    },
    "Building on Others": {
      "Business Acumen": "### A. Reflect on the Past — Results Delivered and How",
      "Citizenship": "**Collaboration and Recognition Culture**",
      "Technical Excellence": "### A. Reflect on the Past — Results Delivered and How",
      "Role Excellence": "**Coaching and Mentoring**",
    },
    "Contributing to Others' Success": {
      "Business Acumen": "### A. Reflect on the Past — Results Delivered and How",
      "Citizenship": "**Community Impact and Technical Leadership**",
      "Technical Excellence": "### A. Reflect on the Past — Results Delivered and How",
      "Role Excellence": "**Coaching and Mentoring**",
    },
  };
  return sectionMap[circle]?.[dimension] || "### A. Reflect on the Past — Results Delivered and How";
}

const PROSE_CONVERSION_SYSTEM_PROMPT = `You are a senior performance-review writer. You will receive:
1. A structured Connect draft in evidence-ledger format (with Theme, Claim Summary, Period, People/Orgs, What Did I Do, How I Did It, Business Value, Related Metric fields).
2. The measuring-stick rubric for reference.

Your job is to convert the structured draft into a polished, natural first-person narrative that reads as a professionally authored Connect document — ready for submission to a manager.

CONVERSION RULES:
- Write in first person throughout, as if the employee wrote this directly to their manager.
- Use a confident, strategic, professional tone that showcases impact clearly.
- Convert each evidence-ledger block into flowing narrative paragraphs. Do NOT use the labelled field format (Theme/Claim Summary/Period/etc.) in the output.
- Group related accomplishments naturally under section headings (e.g., "### Customer Impact at Scale", "### Azure Platform Breadth and Technical Leadership", "### Revenue and Consumption Impact", "### Thought Leadership and Community Contribution", "### Recognition and Coaching").
- Preserve the existing document structure: "## Reflect on the Past: Results Delivered", "## Reflect on Recent Setbacks: Lessons Learned and Growth", "## Plan for the Future: Goals for the Upcoming Period".
- Bold customer/org names on first mention (e.g., **Hydro One**).
- Keep every number, metric, date, and name exact — do not round, approximate, or reinterpret.
- Do NOT invent or embellish. Only use content from the input draft.
- Do NOT drop any evidence, claims, or metrics. Every item in the ledger must appear in the prose.
- Do NOT include meta commentary about search quality, evidence coverage, rubric grading, model reasoning, or tool behavior.
- Do NOT mention LLMs, prompts, evaluation logic, or process language.
- Exclude personal-life content entirely.
- The setbacks section should feel reflective and honest — lessons learned, not excuses.
- The future goals section should feel forward-looking and action-oriented.
- Use markdown formatting: # for title, ## for major sections, ### for subsections, **bold** for emphasis and names.

OUTPUT:
- Return the complete Connect draft as polished first-person narrative in markdown. Nothing else — no preamble, no commentary.`;

const MERGE_SYSTEM_PROMPT = `You are a senior performance-review writer. You will receive:
1. An existing quarterly Connect draft.
2. An evidence file containing WorkIQ search results organised by rubric dimension.

Your job is to merge the evidence into the Connect draft, producing a complete updated draft.

STRUCTURE RULES:
- Do NOT use tables anywhere in the output. All content must be natural language paragraphs.
- For each piece of evidence or accomplishment, structure it as a labelled paragraph block using EXACTLY these fields:

**Theme:** <Sentence(s) describing the overarching theme or category>
**Claim Summary:** <Sentence(s) summarising the accomplishment or claim>
**Period:** <Sentence(s) noting the time frame>
**People / Orgs:** <Sentence(s) naming the people, teams, or organisations involved>
**What Did I Do:** <Sentence(s) describing the result, outcome, or deliverable>
**How I Did It:** <Sentence(s) describing the specific actions, decisions, methods, behaviours, and collaborations>
**Business Value:** <Sentence(s) explaining the impact or value to the business>
**Related Metric:** <Sentence(s) citing any relevant metrics, numbers, or KPIs>

- Group related items under the appropriate existing section headings in the draft.
- Where evidence from the evidence file corresponds to content already in the draft, enrich the existing item with the new detail. Do NOT duplicate items.
- Where evidence introduces entirely new accomplishments not in the draft, add them as new items in the most appropriate section.
- If a field has no information available, write "N/A" for that field. Do not omit the field.

WRITING RULES:
- First-person, strategic, professional tone throughout.
- Write as if the employee is speaking directly to their manager in their own voice.
- Showcase impact clearly and confidently while staying evidence-backed.
- Keep every number, metric, and date exact — do not round or reinterpret.
- Do NOT invent evidence. Only use content from the draft and the evidence file.
- Paraphrase sensitive content — do not copy verbatim.
- Preserve all existing draft content. Restructure it into the field format above but do not remove information.
- NEVER use markdown tables (| col | col |). Use only the labelled paragraph format above.
- Do NOT include meta commentary about search quality, evidence coverage, what was/was not found, rubric grading, model reasoning, or tool behavior.
- Do NOT mention LLMs, prompts, internal evaluation logic, or any "I found/I could not find" process language.
- Exclude personal-life content entirely. Do not include personal finance/banking/mortgage, family, health, travel, or other non-work personal matters.
- If personal/non-work evidence appears in inputs, ignore it and do not mention it.
- METRIC ATTRIBUTION: Many metrics from the Power BI report are team or territory-level aggregates, NOT the individual's personal output. Do not attribute aggregate numbers as personal accomplishments. An individual typically contributes 1–3 engagements per customer. If a metric is labelled as team/territory/org-level, present it as context ("My territory achieved…") rather than a personal claim ("I delivered…"). When scope is ambiguous, default to team/territory attribution.

OUTPUT:
- Return the complete updated Connect draft in markdown. Nothing else — no preamble, no commentary.`;

async function mergeEvidenceIntoDraft(client, deployment, draftContent, rubricContent, evidenceContent) {
  console.log(`  Sending draft + evidence to LLM for merge...`);

  const response = await client.chat.completions.create({
    model: deployment,
    messages: [
      { role: "system", content: MERGE_SYSTEM_PROMPT },
      {
        role: "user",
        content:
          `=== CURRENT CONNECT DRAFT ===\n\n${draftContent}\n\n=== END DRAFT ===\n\n` +
          `=== EVIDENCE FILE ===\n\n${evidenceContent}\n\n=== END EVIDENCE ===\n\n` +
          `Merge the evidence into the draft. Structure every accomplishment using the labelled paragraph format (Theme, Claim Summary, Period, People / Orgs, What Did I Do, How I Did It, Business Value, Related Metric). No tables. Write in first person directly to the manager, as if personally authored by the employee. Exclude any meta/process language (including what was/was not found, search diagnostics, or model reasoning). Exclude all personal/non-work content (for example banking, mortgage renewal, personal finance, family, health, travel). Return the complete updated draft.`,
      },
    ],
    temperature: 0.3,
    max_completion_tokens: 16384,
  });

  const merged = response.choices[0].message.content?.trim();
  if (!merged) {
    console.log(`  ⚠ LLM returned empty response. Draft unchanged.`);
    return draftContent;
  }

  console.log(`  ✓ Merged draft received (${merged.length} chars)`);
  return merged;
}

// ── Prose conversion with validation ───────────────────────────────────────

/**
 * Convert a structured evidence-ledger draft into polished first-person
 * narrative prose, then re-evaluate to ensure the score holds.
 * If the prose version scores lower, retry up to maxRetries times with
 * feedback about what was lost.
 * Returns { prose, score } with the best prose version.
 */
async function convertToProseAndValidate(client, deployment, ledgerContent, rubricContent, maxRetries = 2) {
  console.log(`\n${"═".repeat(60)}`);
  console.log("PROSE CONVERSION — Converting evidence ledger to narrative");
  console.log("═".repeat(60));

  // Evaluate the ledger baseline score
  console.log("  Evaluating ledger draft baseline...");
  const ledgerEval = await evaluateDraft(client, deployment, ledgerContent, rubricContent);
  const ledgerScore = ledgerEval.cells.filter((c) => c.rating === "Exceptional").length;
  console.log(`  Ledger baseline: ${ledgerScore}/12 Exceptional`);

  let bestProse = null;
  let bestProseScore = 0;
  let feedback = "";

  for (let attempt = 1; attempt <= maxRetries + 1; attempt++) {
    console.log(`\n  Prose conversion attempt ${attempt}/${maxRetries + 1}...`);

    const userContent = feedback
      ? `=== STRUCTURED CONNECT DRAFT ===\n\n${ledgerContent}\n\n=== END DRAFT ===\n\n=== RUBRIC ===\n\n${rubricContent}\n\n=== END RUBRIC ===\n\nCONVERSION FEEDBACK FROM PREVIOUS ATTEMPT:\n${feedback}\n\nConvert the structured draft to polished first-person narrative prose. Address the feedback above — ensure no evidence or "how" detail is lost. Return the complete Connect draft.`
      : `=== STRUCTURED CONNECT DRAFT ===\n\n${ledgerContent}\n\n=== END DRAFT ===\n\n=== RUBRIC ===\n\n${rubricContent}\n\n=== END RUBRIC ===\n\nConvert the structured draft to polished first-person narrative prose. Return the complete Connect draft.`;

    const response = await client.chat.completions.create({
      model: deployment,
      messages: [
        { role: "system", content: PROSE_CONVERSION_SYSTEM_PROMPT },
        { role: "user", content: userContent },
      ],
      temperature: 0.3,
      max_completion_tokens: 16384,
    });

    const prose = response.choices[0].message.content?.trim();
    if (!prose) {
      console.log(`  ⚠ LLM returned empty prose. Skipping attempt.`);
      continue;
    }
    console.log(`  ✓ Prose draft received (${prose.length} chars)`);

    // Evaluate the prose version
    console.log(`  Evaluating prose draft...`);
    const proseEval = await evaluateDraft(client, deployment, prose, rubricContent);
    const proseScore = proseEval.cells.filter((c) => c.rating === "Exceptional").length;
    printEvaluationSummary(proseEval, `prose-attempt-${attempt}`);

    if (proseScore > bestProseScore) {
      bestProse = prose;
      bestProseScore = proseScore;
    }

    if (proseScore >= ledgerScore) {
      console.log(`\n  ✓ Prose conversion maintained score (${proseScore}/${ledgerScore}). Accepting.`);
      return { prose, score: proseScore };
    }

    // Build feedback for next attempt: which cells dropped?
    const droppedCells = proseEval.cells.filter((pc) => {
      const lc = ledgerEval.cells.find(
        (l) => l.circle === pc.circle && l.dimension === pc.dimension
      );
      return lc && lc.rating === "Exceptional" && pc.rating !== "Exceptional";
    });

    if (droppedCells.length > 0) {
      feedback = `The prose version dropped these cells from Exceptional:\n` +
        droppedCells.map((c) =>
          `- [${c.circle} → ${c.dimension}]: ${c.reasoning}\n  Missing: ${c.improvement}`
        ).join("\n") +
        `\n\nEnsure the prose retains all the specific evidence, "how" detail, metrics, and actions that support Exceptional ratings for these cells.`;
      console.log(`\n  ⚠ Prose scored ${proseScore} vs ledger ${ledgerScore}. ${droppedCells.length} cell(s) dropped. Retrying with feedback...`);
    } else {
      console.log(`\n  ⚠ Prose scored ${proseScore} vs ledger ${ledgerScore}. Retrying...`);
      feedback = `The prose version scored lower than the ledger (${proseScore} vs ${ledgerScore}). Ensure all evidence, metrics, "how" narratives, and specific actions are preserved in the prose conversion.`;
    }
  }

  // Return best prose even if it didn't match ledger score
  if (bestProse) {
    console.log(`\n  ⚠ Best prose scored ${bestProseScore} vs ledger ${ledgerScore}. Using best prose version.`);
    return { prose: bestProse, score: bestProseScore };
  }

  // Fallback: return ledger content if all prose attempts failed
  console.log(`\n  ⚠ All prose conversions failed. Keeping ledger format.`);
  return { prose: ledgerContent, score: ledgerScore };
}

// ── Clean up previous refinement artifacts ─────────────────────────────────

function cleanRefinementArtifacts() {
  const patterns = [
    /^Connect-Draft-v\d+\.md$/,
    /^Connect-Draft-ledger\.md$/,
    /^evaluation-pass-\d+\.json$/,
    /^evaluation-final\.json$/,
    /^evaluation-prose\.json$/,
    /^workiq-evidence-pass-\d+(-batch-\d+)?\.md$/,
    /^workiq-fleet-prompt-pass-\d+(-batch-\d+)?\.txt$/,
  ];

  if (!fs.existsSync(TEMP_DIR)) return;

  const files = fs.readdirSync(TEMP_DIR);
  let removed = 0;
  for (const file of files) {
    if (patterns.some((p) => p.test(file))) {
      fs.unlinkSync(path.join(TEMP_DIR, file));
      removed++;
    }
  }

  if (removed > 0) {
    console.log(`  🧹 Cleaned ${removed} artifact(s) from previous refinement runs.`);
  }
}

// ── Main refinement loop ───────────────────────────────────────────────────

async function runRefinementLoop(draftPath, maxPasses) {
  if (!fs.existsSync(MEASURING_STICK)) {
    console.error(`Error: Measuring stick not found at ${MEASURING_STICK}`);
    process.exit(1);
  }
  if (!fs.existsSync(draftPath)) {
    console.error(`Error: Connect draft not found at ${draftPath}`);
    process.exit(1);
  }

  const rubricContent = fs.readFileSync(MEASURING_STICK, "utf-8");
  let draftContent = fs.readFileSync(draftPath, "utf-8");
  let bestDraftContent = draftContent;
  let bestExceptionalCount = 0;

  // Clean up artifacts from previous refinement runs
  cleanRefinementArtifacts();

  console.log("Connecting to Azure OpenAI for measuring-stick evaluation...");
  const { client, deployment } = await createAzureOpenAIClient();

  for (let pass = 1; pass <= maxPasses; pass++) {
    console.log(`\n${"═".repeat(60)}`);
    console.log(`REFINEMENT PASS ${pass}/${maxPasses} — Evaluating draft against measuring stick`);
    console.log("═".repeat(60));

    // Step 1: Evaluate
    const evaluation = await evaluateDraft(client, deployment, draftContent, rubricContent);

    // Save evaluation
    const evalPath = path.join(TEMP_DIR, `evaluation-pass-${pass}.json`);
    fs.writeFileSync(evalPath, JSON.stringify(evaluation, null, 2), "utf-8");
    console.log(`  Evaluation saved → ${evalPath}`);

    const exceptionalCount = printEvaluationSummary(evaluation, pass);

    // Persist the best-known draft as soon as a pass-start evaluation improves.
    // This prevents later merge reverts from falling back to an older, lower score.
    if (exceptionalCount > bestExceptionalCount || pass === 1) {
      if (exceptionalCount > bestExceptionalCount && pass > 1) {
        console.log(`  ✓ New best baseline at pass start: ${bestExceptionalCount} → ${exceptionalCount} Exceptional.`);
      }
      bestExceptionalCount = exceptionalCount;
      bestDraftContent = draftContent;
    }

    // Step 2: Check if target met
    if (exceptionalCount >= TARGET_EXCEPTIONAL_CELLS) {
      console.log(`\n✓ ${exceptionalCount}/12 cells at Exceptional (target: ${TARGET_EXCEPTIONAL_CELLS}). Refinement complete.`);
      bestDraftContent = draftContent;
      fs.writeFileSync(draftPath, bestDraftContent, "utf-8");
      return;
    }

    // Step 3: Collect non-Exceptional cells
    const gaps = evaluation.cells.filter((c) => c.rating !== "Exceptional");
    console.log(`\n  ${gaps.length} cell(s) below Exceptional — launching WorkIQ evidence search...`);

    // Step 3a: Seed gaps with user-context.txt hints (if the file exists)
    // Rotate hints across gaps so each cell gets different user context
    if (fs.existsSync(USER_CONTEXT_FILE)) {
      const userContextLines = fs.readFileSync(USER_CONTEXT_FILE, "utf-8")
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter(Boolean);
      if (userContextLines.length > 0) {
        for (let gi = 0; gi < gaps.length; gi++) {
          // Rotate: each gap starts from a different offset in the hints array
          const rotated = [];
          for (let h = 0; h < userContextLines.length; h++) {
            rotated.push(userContextLines[(gi * 2 + h) % userContextLines.length]);
          }
          const extra = rotated.join(", ");
          gaps[gi].userContext = gaps[gi].userContext ? `${extra}, ${gaps[gi].userContext}` : extra;
        }
        console.log(`  Loaded ${userContextLines.length} hint(s) from user-context.txt, rotated across ${gaps.length} gap(s).`);
      }
    }

    // Step 3b: Clarify ambiguous search terms interactively (pass 1 only)
    if (!noClarify && pass === 1) {
      await clarifySearchTerms(gaps);
    }

    // Step 4: Generate WorkIQ fleet prompts in smaller, jittered batches
    const evidenceFilePath = path.join(TEMP_DIR, `workiq-evidence-pass-${pass}.md`);
    const gapBatches = chunkArray(gaps, workiqBatchSize);
    const batchEvidencePaths = [];

    // Delete any stale pass-level evidence file so we can combine fresh batch files
    if (fs.existsSync(evidenceFilePath)) {
      fs.unlinkSync(evidenceFilePath);
    }

    console.log(`  WorkIQ batching enabled: ${gapBatches.length} batch(es), batch size ${workiqBatchSize}, concurrency cap ${workiqMaxConcurrency}`);
    console.log(`  Jitter window: ${workiqJitterMinMs}-${workiqJitterMaxMs} ms; retries per batch: ${workiqRetries}`);

    for (let batchIndex = 0; batchIndex < gapBatches.length; batchIndex++) {
      const batchGaps = gapBatches[batchIndex];
      const batchNumber = batchIndex + 1;
      const batchLabel = `pass-${pass}-batch-${batchNumber}-of-${gapBatches.length}`;
      const fleetPrompt = generateWorkIQFleetPrompt(batchGaps, pass, quarter, {
        parallelLimit: workiqMaxConcurrency,
        batchLabel,
      });
      const fleetPromptPath = path.join(TEMP_DIR, `workiq-fleet-prompt-pass-${pass}-batch-${batchNumber}.txt`);
      const batchEvidencePath = path.join(TEMP_DIR, `workiq-evidence-pass-${pass}-batch-${batchNumber}.md`);
      batchEvidencePaths.push(batchEvidencePath);

      fs.writeFileSync(fleetPromptPath, fleetPrompt, "utf-8");
      console.log(`\n  Batch ${batchNumber}/${gapBatches.length}: prompt saved → ${fleetPromptPath}`);
      console.log(`  Batch ${batchNumber}/${gapBatches.length}: evidence target → ${batchEvidencePath}`);

      if (fs.existsSync(batchEvidencePath)) {
        fs.unlinkSync(batchEvidencePath);
      }

      const jitterMs = getWorkIQJitterMs();
      if (batchNumber > 1 && jitterMs > 0) {
        console.log(`  Staggering WorkIQ call by ${jitterMs} ms to reduce timeout contention...`);
        await sleep(jitterMs);
      }

      let success = false;
      for (let attempt = 0; attempt <= workiqRetries; attempt++) {
        const attemptNumber = attempt + 1;
        console.log(`  Launching Copilot for batch ${batchNumber}/${gapBatches.length} (attempt ${attemptNumber}/${workiqRetries + 1})...`);
        const result = runCopilotEvidenceCapture(fleetPromptPath, batchEvidencePath);

        if (result.ok) {
          const source = result.fromFile ? "file" : "stdout";
          console.log(`  ✓ Batch ${batchNumber} completed from ${source} (${result.bytes} bytes)`);
          success = true;
          break;
        }

        if (result.timedOut && attempt < workiqRetries) {
          const backoff = workiqRetryBackoffMs * (attempt + 1);
          const retryJitter = getWorkIQJitterMs();
          console.log(`  ⚠ Batch ${batchNumber} hit MCP timeout signature. Retrying after ${backoff + retryJitter} ms...`);
          await sleep(backoff + retryJitter);
          continue;
        }

        if (!result.timedOut && attempt < workiqRetries) {
          const retryJitter = getWorkIQJitterMs();
          console.log(`  ⚠ Batch ${batchNumber} failed. Retrying after ${retryJitter} ms...`);
          await sleep(retryJitter);
          continue;
        }

        console.log(`  ⚠ Batch ${batchNumber} failed after ${workiqRetries + 1} attempt(s).`);
        if (result.message) {
          console.log(`    Last error: ${result.message}`);
        }
      }

      if (!success) {
        console.log(`  ⚠ Continuing with next batch; this batch may have partial evidence only.`);
      }
    }

    // Check if Copilot already wrote the pass-level evidence file directly via its file tools.
    // If so, prefer that over the batch stdout transcripts (which contain tool-call logs, not structured evidence).
    const copilotWroteDirectly = fs.existsSync(evidenceFilePath)
      && fs.statSync(evidenceFilePath).size > 500
      && !/^●\s|ask_work_iq|MCP error/m.test(fs.readFileSync(evidenceFilePath, "utf-8").substring(0, 2000));

    if (copilotWroteDirectly) {
      console.log(`\n  ✓ Copilot wrote evidence directly → ${evidenceFilePath} (${fs.statSync(evidenceFilePath).size} bytes) — skipping batch combination.`);
    } else {
      const combinedBatchEvidence = [];
      for (let i = 0; i < batchEvidencePaths.length; i++) {
        const batchPath = batchEvidencePaths[i];
        if (!fs.existsSync(batchPath)) continue;
        const content = fs.readFileSync(batchPath, "utf-8").trim();
        if (!content) continue;
        combinedBatchEvidence.push(`# Batch ${i + 1}\n\n${content}`);
      }

      if (combinedBatchEvidence.length > 0) {
        const joined = combinedBatchEvidence.join("\n\n---\n\n") + "\n";
        fs.writeFileSync(evidenceFilePath, joined, "utf-8");
        console.log(`\n  ✓ Combined pass evidence saved → ${evidenceFilePath} (${joined.length} chars)`);
      }
    }

    // Step 6: Read gathered evidence
    if (!fs.existsSync(evidenceFilePath)) {
      console.log(`\n  ⚠ Evidence file not found at ${evidenceFilePath}.`);
      console.log(`    If Copilot saved it elsewhere, copy it to that path and run --refine-only.`);
      console.log(`    Skipping merge for this pass (no evidence to integrate).\n`);
    } else {
      let evidenceContent = fs.readFileSync(evidenceFilePath, "utf-8");
      const moreEvidencePath = path.join(__dirname, "more-evidence", "more-eveidence.md");
      if (fs.existsSync(moreEvidencePath)) {
        const moreEvidence = fs.readFileSync(moreEvidencePath, "utf-8");
        evidenceContent += "\n\n# Additional Evidence\n\n" + moreEvidence;
        console.log(`  ✓ Additional evidence loaded (${moreEvidence.length} chars)`);
      }
      console.log(`\n  ✓ Evidence file loaded (${evidenceContent.length} chars)`);

      // Step 7: Merge evidence into draft
      console.log(`  Merging new evidence into Connect draft...`);
      const candidateDraft = await mergeEvidenceIntoDraft(
        client, deployment, draftContent, rubricContent, evidenceContent
      );

      // Step 8: Re-evaluate the candidate draft and only keep if score improved
      console.log(`  Re-evaluating merged draft...`);
      const candidateEval = await evaluateDraft(client, deployment, candidateDraft, rubricContent);
      const candidateCount = printEvaluationSummary(candidateEval, `${pass}-candidate`);

      if (candidateCount > bestExceptionalCount) {
        console.log(`\n  ✓ Score improved: ${bestExceptionalCount} → ${candidateCount} Exceptional. Keeping merged draft.`);
        draftContent = candidateDraft;
        bestDraftContent = candidateDraft;
        bestExceptionalCount = candidateCount;
      } else {
        console.log(`\n  ⚠ Score did not improve (${candidateCount} vs best ${bestExceptionalCount}). Discarding this pass's merge.`);
        draftContent = bestDraftContent; // revert to best known draft
      }
    }

    // Save versioned draft (always the best so far)
    const versionPath = path.join(TEMP_DIR, `Connect-Draft-v${pass}.md`);
    fs.writeFileSync(versionPath, draftContent, "utf-8");
    console.log(`  Refined draft saved → ${versionPath}`);
  }

  // Final evaluation after last pass
  console.log(`\n${"═".repeat(60)}`);
  console.log(`FINAL EVALUATION — Post-refinement check`);
  console.log("═".repeat(60));

  const finalEval = await evaluateDraft(client, deployment, bestDraftContent, rubricContent);
  const finalEvalPath = path.join(TEMP_DIR, `evaluation-final.json`);
  fs.writeFileSync(finalEvalPath, JSON.stringify(finalEval, null, 2), "utf-8");
  const finalCount = printEvaluationSummary(finalEval, "final");

  if (finalCount >= TARGET_EXCEPTIONAL_CELLS) {
    console.log(`\n✓ ${finalCount}/12 cells at Exceptional (target: ${TARGET_EXCEPTIONAL_CELLS}). Refinement complete.`);
  } else {
    console.log(`\n⚠ ${finalCount}/12 cells at Exceptional after ${maxPasses} passes (target: ${TARGET_EXCEPTIONAL_CELLS}).`);
    console.log(`  Review the remaining gaps in ${finalEvalPath} and consider:`);
    console.log(`  - Adding missing evidence manually where [EVIDENCE NEEDED] placeholders appear`);
    console.log(`  - Running again with --refine-only --max-refine-passes <N>`);
  }

  // Save the evidence-ledger version for reference
  const ledgerPath = path.join(TEMP_DIR, "Connect-Draft-ledger.md");
  fs.writeFileSync(ledgerPath, bestDraftContent, "utf-8");
  console.log(`\n  Ledger draft saved → ${ledgerPath}`);

  // Convert to polished prose and validate score holds
  const { prose, score: proseScore } = await convertToProseAndValidate(
    client, deployment, bestDraftContent, rubricContent
  );

  // Save prose evaluation
  const proseEval = await evaluateDraft(client, deployment, prose, rubricContent);
  const proseEvalPath = path.join(TEMP_DIR, `evaluation-prose.json`);
  fs.writeFileSync(proseEvalPath, JSON.stringify(proseEval, null, 2), "utf-8");
  console.log(`  Prose evaluation saved → ${proseEvalPath}`);

  // Overwrite the main draft with the prose version
  fs.writeFileSync(draftPath, prose, "utf-8");
  console.log(`  ✓ Final prose draft saved → ${draftPath} (score: ${proseScore}/12)`);
}

// ── Generate Word doc from markdown ────────────────────────────────────────
function generateWordDoc(mdContent, outputPath) {
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = docx;

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

  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
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
    } else if (/^\s*[-*]\s/.test(line)) {
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
} else if (mergeOnly) {
  // ── Merge-only mode: apply existing evidence to draft and generate Word doc ─
  const draftPath = path.join(TEMP_DIR, "Connect-Draft.md");
  const evidencePath = path.join(TEMP_DIR, `workiq-evidence-pass-${mergePass}.md`);

  if (!fs.existsSync(draftPath)) {
    console.error(`Error: ${draftPath} not found.`);
    process.exit(1);
  }
  if (!fs.existsSync(evidencePath)) {
    console.error(`Error: ${evidencePath} not found.`);
    process.exit(1);
  }

  (async () => {
    const rubricContent = fs.readFileSync(MEASURING_STICK, "utf-8");
    let draftContent = fs.readFileSync(draftPath, "utf-8");
    let evidenceContent = fs.readFileSync(evidencePath, "utf-8");
    const moreEvidencePath = path.join(__dirname, "more-evidence", "more-eveidence.md");
    if (fs.existsSync(moreEvidencePath)) {
      const moreEvidence = fs.readFileSync(moreEvidencePath, "utf-8");
      evidenceContent += "\n\n# Additional Evidence\n\n" + moreEvidence;
      console.log(`  ✓ Additional evidence loaded (${moreEvidence.length} chars)`);
    }

    console.log(`Merging evidence from pass ${mergePass} into Connect draft...`);
    console.log(`  Draft: ${draftContent.length} chars`);
    console.log(`  Evidence: ${evidenceContent.length} chars`);

    const { client, deployment } = await createAzureOpenAIClient();
    draftContent = await mergeEvidenceIntoDraft(
      client, deployment, draftContent, rubricContent, evidenceContent
    );

    // Save ledger version
    const ledgerPath = path.join(TEMP_DIR, "Connect-Draft-ledger.md");
    fs.writeFileSync(ledgerPath, draftContent, "utf-8");
    console.log(`\n✓ Merged ledger draft saved → ${ledgerPath}`);

    // Convert to prose and validate
    const { prose } = await convertToProseAndValidate(
      client, deployment, draftContent, rubricContent
    );
    fs.writeFileSync(draftPath, prose, "utf-8");
    console.log(`✓ Prose draft saved → ${draftPath}`);

    console.log("\nGenerating Word document...");
    const wordPath = path.join(TEMP_DIR, "final.docx");
    await generateWordDoc(prose, wordPath);
    console.log(`✓ Word document saved → ${wordPath}`);
  })().catch((err) => {
    console.error("Merge failed:", err.message);
    process.exit(1);
  });
} else if (refineOnly) {
  // ── Refine-only mode: run measuring-stick loop on existing draft ──────────
  const draftPath = path.join(TEMP_DIR, "Connect-Draft.md");

  (async () => {
    await runRefinementLoop(draftPath, maxRefinePasses);

    console.log("\n" + "═".repeat(60));
    console.log("Generating Word document from refined draft...");
    console.log("═".repeat(60));

    const mdContent = fs.readFileSync(draftPath, "utf-8");
    const wordPath = path.join(TEMP_DIR, "final.docx");
    await generateWordDoc(mdContent, wordPath);
    console.log(`✓ Word document saved → ${wordPath}`);
  })().catch((err) => {
    console.error("Refinement failed:", err.message);
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

// Load any additional evidence files from more-evidence/
let additionalEvidenceBlocks = [];
if (fs.existsSync(MORE_EVIDENCE_DIR)) {
  const evidenceFiles = fs.readdirSync(MORE_EVIDENCE_DIR).filter(f => f.endsWith(".md") || f.endsWith(".txt"));
  for (const file of evidenceFiles) {
    const filePath = path.join(MORE_EVIDENCE_DIR, file);
    const content = fs.readFileSync(filePath, "utf-8");
    additionalEvidenceBlocks.push({ file, content });
    console.log(`  Additional evidence loaded: ${file} (${content.length} chars)`);
  }
}

// Build the merged file: quarter context + full instruction pack + full metrics + additional evidence.
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
for (const { file, content } of additionalEvidenceBlocks) {
  merged += `=== ADDITIONAL EVIDENCE: ${file} ===\n\n`;
  merged += content.trimEnd() + `\n\n`;
  merged += `=== END ADDITIONAL EVIDENCE: ${file} ===\n\n`;
}
merged += `=== OUTPUT INSTRUCTIONS ===\n\n`;
merged += `Write the final Connect in first person, as if I personally wrote it directly to my manager.\n`;
merged += `Use a confident, professional tone that clearly showcases my work, impact, decisions, and leadership.\n`;
merged += `Do NOT include any AI/meta commentary, evidence diagnostics, or statements about what was or was not found.\n`;
merged += `Do NOT mention LLMs, prompts, searches, reasoning steps, rubric scoring, or evaluation logic in the final draft.\n`;
merged += `Exclude personal/non-work content entirely (for example banking, mortgage renewal, personal finance, family, health, travel, or private admin tasks).\n`;
merged += `If the instruction pack includes coverage checks or gap analysis, treat those as internal process only and keep them out of the final submitted Connect narrative.\n`;
merged += `Output only the final Connect content, ready to submit.\n`;
merged += `When complete, save the final Connect draft as the file: temp/Connect-Draft.md\n\n`;
merged += `=== END OUTPUT INSTRUCTIONS ===\n`;

fs.writeFileSync(FLEET_PROMPT_FILE, merged, "utf-8");
console.log(`Merged prompt saved → ${FLEET_PROMPT_FILE}`);
console.log(`  Instructions: ${instructionsContent.length} chars`);
console.log(`  Metrics:      ${metricsContent.length} chars`);
if (additionalEvidenceBlocks.length > 0) {
  const addlChars = additionalEvidenceBlocks.reduce((sum, b) => sum + b.content.length, 0);
  console.log(`  Additional:   ${addlChars} chars across ${additionalEvidenceBlocks.length} file(s)`);
}
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

const setupCommands = `Run the following as individual commands:\n\n/allow-all\nworkiq accepteula\nExecute Prompt: @'${FLEET_PROMPT_FILE}'\n/exit\n`;

try {
  execFileSync("clip", [], { input: setupCommands, cwd: ROOT });
  console.log("\n✓ Setup commands copied to clipboard.");
  console.log("  Once Copilot opens, paste from clipboard (Ctrl+V) into the prompt.\n");
  console.log("  The clipboard contains:");
  console.log("Run the following as individual commands:\n\n")
  console.log("    /allow-all");
  console.log("    workiq accepteula");
  console.log(`    Execute Prompt: @${FLEET_PROMPT_FILE}'\n`);
  console.log("    /exit");
} catch {
  console.log("\nCould not copy to clipboard. Run these commands manually in Copilot:");
  console.log("  /allow-all");
  console.log("  workiq accepteula");
  console.log(`  Execute Prompt: @${FLEET_PROMPT_FILE}'\n`);
  console.log("  /exit");
}

console.log("Launching Copilot CLI...\n");

try {
  execFileSync("copilot", [], { cwd: ROOT, stdio: "inherit", shell: true });
} catch {
  // Copilot exited — not necessarily an error
}

console.log("\n" + "═".repeat(60));
console.log("Copilot session ended.");
console.log("═".repeat(60));

const draftPath = path.join(TEMP_DIR, "Connect-Draft.md");
if (fs.existsSync(draftPath)) {
  (async () => {
    // ── Step 5: Measuring-stick refinement loop ──────────────────────────────
    if (!skipRefine) {
      await runRefinementLoop(draftPath, maxRefinePasses);
    } else {
      console.log("\nSkipping refinement loop (--skip-refine).");
    }

    // ── Step 6: Generate Word document ────────────────────────────────────────
    console.log("\n" + "═".repeat(60));
    console.log("Generating Word document...");
    console.log("═".repeat(60));

    const mdContent = fs.readFileSync(draftPath, "utf-8");
    const wordPath = path.join(TEMP_DIR, "final.docx");
    await generateWordDoc(mdContent, wordPath);
    console.log(`✓ Word document saved → ${wordPath}`);
  })().catch((err) => {
    console.error("Post-Copilot processing failed:", err.message);
    process.exit(1);
  });
} else {
  console.log(`Connect-Draft.md not found at ${draftPath}. Skipping refinement and Word generation.`);
  console.log(`If the draft was saved elsewhere, run: node baby-carlos.js --word-only --quarter ${quarter}`);
}
} // end else (full pipeline)

