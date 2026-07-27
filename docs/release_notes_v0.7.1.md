## What's new in v0.7.1

**One attached file per session.** The starter contract is now embedded in the kernel as the `_STARTER` sheet (ADR-019). Attach only `VCH_HarnessCore.xlsx`, type `load vch` — no second file to forget, and starter/kernel version drift is impossible by construction.

- `copilot/copilotstart.txt` stays in the repo as the **source of truth**; the sheet is generated at build time, never hand-edited
- `harness_lint.ps1` gains the **STARTER check**: every non-empty `_STARTER` row must reproduce the txt line-for-line
- `copilot_custom_instructions.txt` remains outside — it configures Copilot itself and cannot be attached; the bootstrap read-set now starts with `_STARTER`

**Upgrade from v0.7.0:** replace all files, re-paste custom instructions. The attach flow drops the separate starter file.

v0.7.0 (superseded) introduced the single-kernel architecture (ADR-017), the NEXT footer (ADR-018) and the renumbering from 6.x (ADR-016).

## Governance

- 43 skills (unchanged), routing corpus 55/55 fixtures, ADR-019 recorded (Proposed until gates pass)
- Python read-back verification: **PASS (35/35)** including the new `_STARTER` equality check; fork simulation: **PASS**
- Structural gate (`harness_lint.ps1` on Windows + Excel COM): **NOT_TESTED** — pending a Windows run (see `docs/WIN_GUIDE.md`)
- Behavioral routing gate (55 fixtures, isolated Copilot contexts): **NOT_TESTED** — pending
- This release is marked as a prerelease until both gates report PASS.
- sha256 of every release file in `VCH_release_manifest.json`
