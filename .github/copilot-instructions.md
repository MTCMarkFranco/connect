# Copilot Instructions for This Repository

## Repository shape

This repository is a documentation-driven workflow for producing a quarterly Connect draft with GitHub Copilot CLI and WorkIQ, not an application codebase with source, build artifacts, or tests.

- `README.md` is the human/operator guide. It covers local setup, WorkIQ plugin or MCP configuration, extracting Power BI metrics into `core metrics.txt`, and the starter `/fleet` invocation.
- `quarterly-connect-fleet-instructions.txt` is the agent execution spec. Treat it as the source of truth for the actual `/fleet` workflow, required evidence coverage, evidence schema, quality bar, and final output structure.

Keep those two documents aligned: README explains how to run the workflow, while the instruction pack defines how Copilot should execute it.

## Commands and workflow entry points

There are no repository-defined build, lint, or automated test commands in this repo.

The important commands are the workflow commands documented in `README.md`:

- Start Copilot CLI in this folder:
  - `copilot`
- Install the recommended WorkIQ plugin:
  - `/plugin install workiq@copilot-plugins`
- If using the explicit MCP route instead of the plugin:
  - `/mcp add`
  - `/mcp show`
  - `/mcp show workiq`
- Start the Connect generation flow:

```text
/fleet Create my quarterly Connect using the instruction pack in @quarterly-connect-fleet-instructions.txt.
Quarter: <quarter>
Core metrics:
@core metrics.txt
```

The working folder is expected to contain `quarterly-connect-fleet-instructions.txt` plus a user-provided or user-generated `core metrics.txt`.

## High-level architecture

The effective architecture is a two-layer prompt system plus user-supplied data:

1. **Setup and input preparation layer (`README.md`)**
   - install Copilot CLI
   - connect WorkIQ, preferably through the plugin
   - extract quarter-specific Power BI metrics into plain text
   - launch `/fleet` from this repository folder

2. **Execution layer (`quarterly-connect-fleet-instructions.txt`)**
   - validate required inputs before doing work
   - split the work into parallel evidence-gathering and synthesis streams
   - search WorkIQ across email, Teams, documents, and Loop
   - assemble an evidence ledger
   - draft the final Connect only after validation

3. **Runtime inputs**
   - quarter or date range
   - `core metrics.txt`
   - optional focus themes, exclusions/sensitivity notes, and current goals/priorities

Future changes should preserve this split: onboarding/setup guidance belongs in the README, while evidence and drafting behavior belongs in the fleet instruction pack.

## Repository-specific conventions

- Do not invent facts, examples, praise, or outcomes. The instruction pack treats unsupported claims as failures, not gaps to smooth over.
- Stop and ask for missing required inputs if either the quarter/date range or core metrics are absent.
- Build the final draft from an evidence ledger first. Every major claim should map back to evidence metadata such as source type, reference, period, people, and business value.
- Prefer evidence from the target quarter. If older material is used for context, label it clearly as prior context.
- Preserve exact metric values from the Power BI extraction. Convert them into strategic narrative carefully and do not overstate causality.
- Explicitly cover all required evidence domains from the instruction pack: customer impact, strategic value, community contributions, events, kudos given, coaching delivered, coaching received, awards/recognition, and setbacks/growth.
- Search across all four WorkIQ surfaces named in the pack: emails, Teams messages, documents, and Loop content.
- Paraphrase sensitive workplace or customer content instead of copying it verbatim, especially when the README or pack calls out sensitivity handling.
- When updating the repository docs, keep the quick-start examples in `README.md` consistent with the required input contract and output order in `quarterly-connect-fleet-instructions.txt`.

## Expected output structure

When following this repository's workflow, the final result should be returned in the order defined by the instruction pack:

1. Input and coverage check
2. Evidence ledger
3. Gaps or follow-up questions
4. Final Connect draft
