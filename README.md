# Quarterly Connect with GitHub Copilot CLI + WorkIQ

This guide shows how to:

1. Install GitHub Copilot CLI
2. Add WorkIQ to Copilot CLI
3. Extract quarterly core metrics from the provided Power BI report as plain text
4. Run the starter `/fleet` command to generate your quarterly Connect

---

## 1. Prerequisites

Before you start, make sure you have:

- A GitHub account with an active Copilot subscription
- Access to GitHub Copilot CLI in your organization
- Access to the Microsoft 365 data you want WorkIQ to query
- Access to the Power BI report:
  - `https://msit.powerbi.com/groups/me/apps/bc16e81a-2071-4f5f-8131-c2e9b7211346/reports/99396eba-3a62-499c-bd28-2165ac0a0737/ReportSectionc481b05a185bb8cce548?experience=power-bi`
- Node.js 18+ if you want to configure WorkIQ as an MCP server manually

On Windows, GitHub Copilot CLI expects PowerShell 6+ or newer.

---

## 2. Install GitHub Copilot CLI

GitHub Copilot CLI supports Windows, macOS, and Linux.

### Windows

Recommended:

```powershell
winget install GitHub.Copilot
```

Prerelease version:

```powershell
winget install GitHub.Copilot.Prerelease
```

### macOS or Linux

Using Homebrew:

```bash
brew install copilot-cli
```

Prerelease version:

```bash
brew install copilot-cli@prerelease
```

Using the install script:

```bash
curl -fsSL https://gh.io/copilot-install | bash
```

or

```bash
wget -qO- https://gh.io/copilot-install | bash
```

### Any platform with npm

```bash
npm install -g @github/copilot
```

Prerelease version:

```bash
npm install -g @github/copilot@prerelease
```

### Verify the install

Open a terminal and start Copilot:

```bash
copilot
```

If prompted:

1. Trust the current folder
2. Run `/login`
3. Complete GitHub authentication

Helpful slash commands once you are inside Copilot CLI:

- `/help` - show command help
- `/model` - choose the model
- `/fleet` - run sub-agents in parallel
- `/mcp` - manage MCP servers
- `/plugin` - manage plugins

---

## 3. Add WorkIQ to GitHub Copilot CLI

There are two good ways to add WorkIQ:

- **Recommended:** install the WorkIQ Copilot plugin
- **Alternative:** add WorkIQ as an explicit MCP server

The plugin route is the fastest. The MCP route is useful if you want the configuration to be explicit and portable.

### Option A - Recommended: install the WorkIQ plugin

Start Copilot CLI:

```bash
copilot
```

Then run:

```text
/plugin install workiq@copilot-plugins
```

After installation:

1. Restart Copilot CLI
2. Ask a simple WorkIQ question such as:

```text
What are my upcoming meetings this week?
```

If this is your first time using WorkIQ, expect:

- EULA acceptance
- Browser sign-in
- Possibly an admin-consent step, depending on your tenant

### Option B - Manual MCP server setup

If you want to add WorkIQ as a normal MCP server in Copilot CLI, make sure Node.js 18+ is installed:

```bash
node --version
npm --version
```

Then start Copilot:

```bash
copilot
```

Run:

```text
/mcp add
```

Fill the form with values like these:

- **Server Name:** `workiq`
- **Server Type:** `STDIO`
- **Command:** `npx`
- **Arguments:** `-y @microsoft/workiq mcp`
- **Environment Variables:** leave blank unless your environment requires tenant-specific settings
- **Tools:** `*`

Press `Ctrl+S` to save.

You can verify the server with:

```text
/mcp show
/mcp show workiq
```

### Optional: configure WorkIQ directly in `~/.copilot/mcp-config.json`

If you prefer editing the file directly, use a config like this:

```json
{
  "mcpServers": {
    "workiq": {
      "type": "local",
      "command": "npx",
      "args": ["-y", "@microsoft/workiq", "mcp"],
      "env": {},
      "tools": ["*"]
    }
  }
}
```

After saving, restart Copilot CLI or use `/mcp show` to confirm it is available.

### Important WorkIQ notes

- WorkIQ queries Microsoft 365 data such as emails, meetings, documents, Teams messages, and more
- The first use may require tenant consent
- If you are not a tenant admin, you may need your administrator to grant access
- If you install WorkIQ standalone outside Copilot CLI, you can also use:

```bash
workiq accept-eula
workiq ask -q "What meetings do I have tomorrow?"
workiq mcp
```

---

## 4. Prepare the workspace

Put these files in your working folder:

- `quarterly-connect-fleet-instructions.txt`
- `core metrics.txt`

If you are using this repository as-is, both files already exist.

Open Copilot CLI in this folder:

```bash
cd C:\Code\connect
copilot
```

---

## 5. Extract the core metrics from Power BI as plain text

The goal is to turn the Power BI report into a clean text file that Copilot can use as input.

### Open the report

Open this URL in your browser:

`https://msit.powerbi.com/groups/me/apps/bc16e81a-2071-4f5f-8131-c2e9b7211346/reports/99396eba-3a62-499c-bd28-2165ac0a0737/ReportSectionc481b05a185bb8cce548?experience=power-bi`

### Filter the report to the right quarter

Before extracting anything:

1. Set the report filters to the correct quarter, for example `FY26Q3`
2. Apply any role, geography, or account filters you normally use for your Connect
3. Make sure the visible page reflects the exact metrics you want summarized

### What to extract

Capture all high-value metrics that could support the Connect narrative, including:

- NSAT or customer satisfaction
- Number of completed engagements
- Outcome documentation coverage
- Geography or hub distribution
- Engagement mix by type
- Solution area mix
- Industry distribution
- Journey counts and statuses
- Opportunities created after engagements
- Any other quarter-specific KPIs visible on the report

### Best practical workflow

Because this Power BI report is tenant-authenticated, the most reliable workflow is:

1. Open the report in the browser
2. Copy the values or visible summaries from the KPI cards, charts, or tables
3. Paste that raw content into Copilot
4. Ask Copilot to rewrite it as plain text
5. Save the result into `core metrics.txt`

### Prompt to extract the metrics as plain text

Use a prompt like this in Copilot:

```text
I am looking at a Power BI report for FY26Q3. Convert the visible report content into plain text.

Requirements:
- Extract every visible key metric
- Include the metric name and value
- Preserve quarter context
- Group related metrics together
- Include trends or distributions only if they are visible in the report
- Do not infer or invent missing numbers
- Write the output as concise business-ready prose and bullets

Be sure to cover:
- NSAT / customer satisfaction
- completed engagements
- outcomes documented
- engagement mix
- solution areas
- industry mix
- journeys and statuses
- opportunities created after engagements
```

### Save the result

Once Copilot has rewritten the report in plain text, save it into:

```text
core metrics.txt
```

If you want, you can ask Copilot for a second pass:

```text
Rewrite this as a stronger quarterly performance summary, but keep every number exact.
```

---

## 6. Start the quarterly Connect command

Once WorkIQ is installed and `core metrics.txt` is ready, start Copilot CLI in this folder and run the starter command.

### Minimal starter command

```text
/fleet Create my quarterly Connect using the instruction pack in @quarterly-connect-fleet-instructions.txt.
Quarter: FY26Q3
Core metrics:
@core metrics.txt
```

### Better starter command with explicit date range and optional guidance

```text
/fleet Create my quarterly Connect using the instruction pack in @quarterly-connect-fleet-instructions.txt.
Quarter: FY26Q3
Date range: Jan 1, 2026 - Mar 31, 2026
Core metrics:
@core metrics.txt

Optional focus themes:
- Strategic business value
- Customer outcomes
- Leadership and multiplier impact

Optional exclusions or sensitivity notes:
- Avoid unsupported revenue claims
- Paraphrase sensitive customer details

Optional current goals or priorities:
- Show stronger executive storytelling
- Highlight coaching delivered and community impact
```

### What Copilot should do next

With the instruction pack in place, Copilot should:

1. Validate that the quarter and core metrics are present
2. Use WorkIQ to gather evidence from emails, Teams, documents, and Loop content
3. Look specifically for:
   - community contributions
   - events
   - kudos given
   - coaching delivered
   - coaching received
   - awards received
4. Build an evidence ledger
5. Draft a full Connect aligned to your sample structure
6. Flag any evidence gaps instead of inventing content

---

## 7. Recommended validation pass

After Copilot produces the draft, review it for:

- exactness of numbers copied from Power BI
- whether every major claim has evidence
- whether quarter boundaries are correct
- whether customer-sensitive details should be softened
- whether community, coaching, kudos, and awards were all included

Useful follow-up prompt:

```text
Review this Connect draft and remove any claim that is not clearly backed by the evidence ledger. Keep the strongest examples and preserve all exact metrics.
```

---

## 8. Quick reference

### Install Copilot CLI

```powershell
winget install GitHub.Copilot
```

### Install WorkIQ plugin

```text
/plugin install workiq@copilot-plugins
```

### Add WorkIQ as MCP server

```text
/mcp add
```

Use:

- Server Name: `workiq`
- Server Type: `STDIO`
- Command: `npx`
- Args: `-y @microsoft/workiq mcp`
- Tools: `*`

### Start the Connect workflow

```text
/fleet Create my quarterly Connect using the instruction pack in @quarterly-connect-fleet-instructions.txt.
Quarter: FY26Q3
Core metrics:
@core metrics.txt
```
