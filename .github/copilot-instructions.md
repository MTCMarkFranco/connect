# Copilot Instructions for This Repository

## Repository shape

This is an automated pipeline for producing a quarterly Connect draft. It combines Playwright-based Power BI scraping, Azure OpenAI summarisation, and GitHub Copilot CLI + WorkIQ evidence gathering.

**Key files:**

- `baby-carlos.js` — Main orchestrator. Runs the full pipeline end-to-end (scrape → summarise → merge prompt → launch Copilot CLI). Also generates a Word `.docx` from the final markdown draft.
- `scrape-powerbi.js` — Playwright script that opens a Power BI report in Edge, captures screenshots + text, and sends them to Azure OpenAI for multi-modal summarisation.
- `gh-cli-prompts/quarterly-connect-fleet-instructions.txt` — The agent execution spec (source of truth for the `/fleet` workflow, evidence coverage, evidence schema, quality bar, and output structure).
- `README.md` — Human/operator setup guide.
- `temp/` — Generated artifacts directory (gitignored). Contains scraped metrics, merged prompts, the Connect draft, and the final `.docx`.

Keep `README.md` and the fleet instruction pack aligned: the README explains how to run the workflow; the instruction pack defines how Copilot should execute it.

## Commands

There are no build, lint, or test commands. The repo has these workflow commands:

```powershell
# Full pipeline (headed browser for first-time Power BI login)
node baby-carlos.js --quarter FY26Q3

# Subsequent runs (headless, reuses cached auth)
node baby-carlos.js --quarter FY26Q3 --headless

# Skip scraping, reuse existing metrics
node baby-carlos.js --quarter FY26Q3 --skip-scrape

# Jump straight to Copilot CLI (requires prior run)
node baby-carlos.js --skip-to-copilot --quarter FY26Q3

# Regenerate Word doc from existing Connect-Draft.md
node baby-carlos.js --word-only --quarter FY26Q3

# Scrape only (no Copilot launch)
node scrape-powerbi.js --quarter FY26Q3
node scrape-powerbi.js --quarter FY26Q3 --headless
```

**Install:**

```powershell
npm install
npx playwright install
```

## High-level architecture

The pipeline has three layers:

1. **Scraping & summarisation** (`scrape-powerbi.js`)
   - Launches Edge via Playwright with a persistent auth context (`.auth/`)
   - Scrolls through the Power BI report capturing screenshots + raw text
   - Sends both to Azure OpenAI (GPT-4o-mini via `DefaultAzureCredential`) for structured summarisation → `temp/final-metrics.md`

2. **Prompt assembly & Copilot launch** (`baby-carlos.js`)
   - Merges the fleet instruction pack + summarised metrics into a single prompt file → `temp/fleet-prompt.txt`
   - Copies prompt to clipboard, verifies Copilot CLI auth, accepts WorkIQ EULA
   - Launches `copilot -i` with the merged prompt piped in
   - On exit: captures Copilot stdout → `temp/Connect-Draft.md`, then generates `temp/final.docx`

3. **Evidence-gathering execution** (fleet instruction pack, run inside Copilot CLI)
   - Validates required inputs (quarter + core metrics)
   - Splits into parallel workstreams: metrics analysis, customer impact, community/events, recognition/coaching, growth/setbacks
   - Searches WorkIQ across emails, Teams, documents, and Loop
   - Assembles an evidence ledger, validates coverage, then drafts the final Connect

## Environment

Requires a `.env` file in the project root:

```
AZURE_OPENAI_ENDPOINT=https://YOUR-RESOURCE.openai.azure.com
AZURE_OPENAI_DEPLOYMENT=gpt-4o-mini
```

Authentication uses `DefaultAzureCredential` (Entra ID) — no API key. Browser auth state is persisted in `.auth/` (gitignored).

## Key conventions

- **No invented content.** The instruction pack treats unsupported claims as failures. Every major claim must map back to evidence with source metadata (type, reference, period, people, business value).
- **Evidence-first drafting.** Build an evidence ledger before writing the Connect draft. Prefer target-quarter evidence; label older material as "prior context."
- **Preserve exact metrics.** Do not round or reinterpret numbers from the Power BI extraction. Convert to narrative carefully without overstating causality.
- **Cover all evidence domains:** customer impact, strategic value, community contributions, events, kudos given, coaching delivered, coaching received, awards/recognition, and setbacks/growth.
- **Search all four WorkIQ surfaces:** emails, Teams messages, documents, and Loop content.
- **Paraphrase sensitive content** — do not copy verbatim, especially customer or workplace specifics.
- **Stop and ask** if required inputs (quarter/date range or core metrics) are missing.

## Expected output structure

When following the fleet workflow, produce results in this order:

1. Input and coverage check
2. Evidence ledger
3. Gaps or follow-up questions
4. Final Connect draft (first-person, strategic, evidence-backed)

