## What's new

**This is a renumbering, not a downgrade.** The 6.x line ends at v6.15.1 and the harness continues as pre-1.0 software (ADR-016). v0.7.0 is the biggest simplification in the project's history:

- **Single kernel (ADR-017):** `VCH_ProjectTemplate.xlsx` is gone. One workbook, `VCH_HarnessCore.xlsx`, carries the complete behavior catalog plus a new `__DELIVERY_SCHEMA` sheet. PROJECT-FORK now *generates* the nine delivery sheets (`01_Plan` .. `09_Worklog`) into revision 001 instead of copying a template. Workbook modes collapse from five to three: HARNESS, PROJECT, MIGRATION.
- **NEXT footer (ADR-018):** every reply ends with a recommended next action and skill, drawn deterministically from `May_Chain_To` filtered by `Allowed_Modes` — never inferred, never auto-chained. Lifecycle replies carry it in the extended Status Card `Next:` line.
- **Renumbering (ADR-016):** version markers, oracles, lint and docs all moved to v0.7.0.

**Upgrade from v6.15.1:** replace all files, re-paste custom instructions. Starting a project no longer needs a second file — attach the kernel, `load vch`, `new project`. Old projects migrate via MIGRATION mode against the current kernel (`Template_Version` is now `Schema_Version`).

## Governance

- 43 skills (unchanged), routing corpus 55/55 fixtures (3 rewritten: RT-031, AV-002, AV-009), 3 new ADRs (016/017/018, status Proposed until gates pass)
- `harness_lint.ps1` gains the `Test-DeliverySchema` check: a forked v001 must match `__DELIVERY_SCHEMA` exactly
- New tooling: `tools/build_v070.py`, `tools/verify_v070.py` (33/33 read-back checks PASS), `tools/build_v001_sample.py` (FORK_GATE PASS)
- sha256 of every release file in `VCH_release_manifest.json`

## Gates status (honest)

- Python read-back verification: **PASS (33/33)**; fork simulation: **PASS**
- Structural gate (`harness_lint.ps1` on Windows + Excel COM): **NOT_TESTED** — pending a Windows run
- Behavioral routing gate (55 fixtures, isolated Copilot contexts): **NOT_TESTED** — pending
- This release is marked as a prerelease until both gates report PASS.
