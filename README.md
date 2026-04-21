# Quarterly Connect — Automated Pipeline

Scrapes Power BI metrics, summarises them with Azure OpenAI, and feeds the result into GitHub Copilot CLI + WorkIQ to draft your quarterly Connect.

## Install

```powershell
# 1. Clone and enter the repo
git clone https://github.com/CarlosRCHT/connect.git
cd connect

# 2. Install Node.js dependencies
npm install

# 3. Install Playwright browsers (Edge channel)
npx playwright install

# 4. Install GitHub Copilot CLI
winget install GitHub.Copilot

# 5. Install WorkIQ plugin (inside Copilot CLI)
copilot
/plugin install workiq@copilot-plugins
# then exit with /exit
```

### Environment variables

Create a `.env` file in the project root:

```
AZURE_OPENAI_ENDPOINT=https://YOUR-RESOURCE.openai.azure.com
AZURE_OPENAI_DEPLOYMENT=gpt-4o-mini
```

Authentication uses `DefaultAzureCredential` (Entra ID) — no API key needed.

## Run

### Full pipeline (first run — headed browser for Power BI login)

```powershell
node run-connect.js --quarter FY26Q3
```
ℹ️- At the end of the run just CTRL-C Twice to get out of Github Copilot CLI and you are done!

📝- Report will be in the "temp" folder

### Subsequent runs (headless)

```powershell
node run-connect.js --quarter FY26Q3 --headless
```

### Skip scraping (reuse existing metrics)

```powershell
node run-connect.js --quarter FY26Q3 --skip-scrape
```

### What happens

1. Edge opens the Power BI report, scrapes text + screenshot
2. Azure OpenAI (GPT-4o-mini) summarises metrics → `temp/final-metrics.md`
3. Fleet instructions + metrics are merged → `temp/fleet-prompt.txt`
4. Prompt is copied to clipboard
5. Copilot CLI auth is verified (auto-login if needed)
6. Copilot CLI launches with the fleet prompt pre-loaded via `copilot -i`
7. The Connect draft is evaluated against the measuring-stick rubric and iteratively refined until every cell reaches "Exceptional impact" (or max passes reached)

### Refinement-only mode

Re-run the measuring-stick evaluation loop on an existing `Connect-Draft.md`:

```powershell
node run-connect.js --refine-only --quarter FY26Q3
node run-connect.js --refine-only --max-refine-passes 5 --quarter FY26Q3
```

### Skip refinement

Run the full pipeline but skip the post-Copilot refinement loop:

```powershell
node run-connect.js --quarter FY26Q3 --skip-refine
```

### Set target score

By default the loop stops when 10/12 cells reach Exceptional. Override with:

```powershell
node run-connect.js --refine-only --target-score 12 --quarter FY26Q3
```

### Scrape only

```powershell
node scrape-powerbi.js --quarter FY26Q3              # headed
node scrape-powerbi.js --quarter FY26Q3 --headless   # headless
```

## Prerequisites

- Node.js 18+
- Microsoft Edge
- PowerShell 6+
- GitHub account with Copilot subscription
- Azure OpenAI resource with a GPT-4o-mini deployment
- Access to the Power BI report
