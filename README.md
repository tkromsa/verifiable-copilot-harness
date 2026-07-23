# Verifiable Copilot Harness

**Turn Microsoft Copilot (Excel + Chat) into a disciplined, evidence-driven project operator — no Cowork, no plugins, no paid add-ons required.**

A behavior harness for Microsoft Copilot in strict corporate environments: 41 skills, frozen acceptance criteria, a deterministic lint gate, revision-controlled project workbooks and evidence-scoped persistence. Copilot stops inventing state, never rewrites files without proof, and can guide a project step by step from intake to verified closeout.

Built for people stuck in locked-down enterprises where Copilot (or a similar LLM chat) is the *only* AI tool allowed.

**Current version: v6.11.0** — adds FOCUS-MODE, an opt-in action-first output skill (ADHD-friendly) adapted from [ayghri/i-have-adhd](https://github.com/ayghri/i-have-adhd). Filenames are version-independent (ADR-010); the version lives inside the artifacts and is reported by `load vch`.

## Why it exists

Out of the box, Copilot in Excel/Chat:

- hallucinates workbook state it never read,
- claims it saved things it never saved,
- silently rewrites files, loses lineage, forgets everything between sessions.

This harness fixes that with a set of sovereign safety rules the model must follow:

| Rule | What it enforces |
| --- | --- |
| **Read-only bootstrap** | Every session starts with a read-only load: version, skill count, capability manifest, resume point. Nothing is created or saved implicitly. |
| **No capability inheritance** | A stored "write verified" from a past session proves nothing. Runtime capability is re-derived live, every time. |
| **Frozen oracle** | Acceptance criteria are immutable during a run. The model never adapts expected values to observed data. |
| **Evidence classes** | INDEPENDENT_READ_BACK vs. conversation claims vs. inference — self-report is not verification. |
| **Persistence router** | Persistent change only via an evidenced mode: copy-on-write, precopied artifact, or in-place checkpoint. Otherwise READ_ONLY / WRITE_BLOCKED. |
| **SCRUB** | No output crosses the trust boundary with raw sensitive values. |
| **Deterministic lint** | Every DETERMINISTIC_SCAN check in the oracle has a real executable implementation in `tools/harness_lint.ps1`. Exit code 0 or it doesn't ship. |

## Repository layout

```
core/
  VCH_HarnessCore.xlsx        # the brain: 41 skills, rules, oracle — attach as knowledge, never write to it
  VCH_ProjectTemplate.xlsx    # immutable project template — fork new projects from it
copilot/
  copilotstart.txt            # session starter instructions — attach with HarnessCore
  copilot_custom_instructions.txt  # paste into Copilot Custom Instructions (one-time)
tools/
  harness_lint.ps1            # integrity validator (PowerShell 5.1+, ImportExcel module)
docs/
  VCH_Cheatsheet_EN.txt       # usage cheatsheet — hand this to your colleagues
```

## Quick start (5 minutes)

1. Paste `copilot/copilot_custom_instructions.txt` into your Copilot custom instructions (one-time).
2. In a Copilot chat, attach `VCH_HarnessCore.xlsx` + `copilotstart.txt` (+ the project template if starting a new project).
3. Type: `load vch` -> read-only bootstrap reports version, 41 skills, capabilities, resume point.
4. Start a project: `[skill: PROJECT-FORK]` with Project ID, Name, Owner, Main Goal -> creates your `v001` project workbook.
5. Work: `guide project` -> PROJECT-GUIDE picks one specialist skill per step, enforces phase gates, and only persists through verified modes.

Full daily-driver guide with commands, failure dictionary and migration playbook: **`docs/VCH_Cheatsheet_EN.txt`**.

## Everyday commands

Triggers are plain English atoms — type them exactly as listed, or route explicitly with `[skill: SKILL-ID]`.

| You type | Skill | What happens |
| --- | --- | --- |
| `load vch` | LOAD-ALL | Read-only bootstrap, resume point |
| `guide project` / `what next` | PROJECT-GUIDE | Guided step-by-step delivery with gates |
| `new project` | PROJECT-FORK | New v001 workbook from template |
| `write plan` | WRITING-PLANS | Implementation plan; every step has VERIFY |
| `debug this` | SYSTEMATIC-DEBUGGING | Root cause first, fix second |
| `review code` | CODE-REVIEW | Findings with severity, no rewriting |
| `security review` | SECURITY-REVIEW | Security assessment + triage |
| `verify fix` | VERIFICATION-CHECKLIST | Proof of result; no-error != proof |
| `update context` | UPDATE-CONTEXT | Checkpoint: writes verified state only |
| `drift check` | DECISION-DRIFT-CHECK | Diff against the original assignment |
| `scrub this` | SCRUB | Sanitizes sensitive values before export |
| `bootstrap check` | BOOTSTRAP-CHECK | Harness + host-mode diagnostics |
| `focus mode` | FOCUS-MODE | Action-first ADHD-friendly output; `focus mode off` deactivates |

## FOCUS-MODE (v6.11.0)

An opt-in output-shaping skill adapted from [ayghri/i-have-adhd](https://github.com/ayghri/i-have-adhd) (MIT): the first line is the next action, multi-step work is numbered, every reply ends with one concrete next step, lists are capped at 5, no preamble or "Hope this helps!" closers.

Activate with `focus mode`, deactivate with `focus mode off`. It never weakens the sovereign rules: evidence classes, `Field=Value` enums, gate verdicts and SCRUB output are never dropped for brevity.

## The three file roles — where you write and where you don't

| File | Role | Write to it? |
| --- | --- | --- |
| HarnessCore | Behavior kernel — HOW Copilot must behave | NEVER |
| ProjectTemplate | Project workbook schema | NEVER — fork only |
| `<ID>_<Name>_v001.xlsx` | Your project workbook | YES — this is where work happens |

## Hacking on the harness

Changing the harness is a governed change:

1. Edit `00_Skills` / rules / oracle in the workbooks.
2. Record an ADR entry.
3. Run the lint gate — it must exit 0:

```powershell
Install-Module ImportExcel -Scope CurrentUser   # once
.\tools\harness_lint.ps1 -Path .\core\VCH_HarnessCore.xlsx, .\core\VCH_ProjectTemplate.xlsx
```

The lint executes every DETERMINISTIC_SCAN declared in `__TEST_ORACLE`: skill count, version agreement, trigger duplicates/prefix collisions, non-ASCII cells, chain resolution, enum declarations, state schema, probe cell, and more.

## Migrating an old project

Never rewrite an old-version workbook. Fork a new revision from the current template, port **verified** content only (map by column names, not positions), verify on reopen, archive the old line. Full playbook in the cheatsheet, section 11.

## Not just for Copilot

The harness is model-agnostic: it was iterated across several LLMs (Copilot/GPT, Claude, Gemini, Grok, GLM, Kimi). The starter + custom instructions adapt to any chat that supports file grounding or custom instructions — the workbooks and the lint stay the same.

## Inspiration and credits

Standing on the shoulders of the open agent-skills community — thanks for the ideas, patterns and prior art:

- **obra / Superpowers** — the original Claude skills harness that proved skills-in-files work
- **GrillMe** — adversarial requirement-grilling pattern
- **Shaping Skills** — skill-shape design ideas
- **Gentle AI** — human-in-the-loop agent workflows
- **[ayghri/i-have-adhd](https://github.com/ayghri/i-have-adhd)** — ADHD-friendly output rules adapted as FOCUS-MODE

This project is an independent implementation for the Microsoft Copilot + Excel world, built and hardened iteratively by a human operator working through a small parliament of LLMs. No corporate resources, names or data inside — the harness is fully anonymized.

## License

MIT — take it, fork it, adapt it to your corporate reality. See [LICENSE](LICENSE).
