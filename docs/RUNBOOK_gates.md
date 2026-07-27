# VCH v0.7.0 — Phase 7 Gates Runbook

Purpose: promote v0.7.0 from prerelease (release candidate) to full release.
Scope: STRUCTURAL_GATE (Windows + Excel), BEHAVIORAL_ROUTING_GATE (Copilot), UAT, promotion.
Safe default: any FAIL stops the runbook; never adapt frozen expected values to observations.

---

## 0. Preconditions

| # | Requirement | Verify |
|---|---|---|
| 0.1 | Windows machine with desktop Excel | Excel opens a blank workbook without repair dialog |
| 0.2 | Repo pulled at `main` including ADR-019 | `git log --oneline -3` shows the v0.7.1 commit |
| 0.3 | `gh` CLI authenticated, scope `repo` | `gh auth status` |
| 0.4 | Microsoft Copilot chat access | can attach files to a chat |
| 0.5 | Python 3 + openpyxl on the Windows machine (for the v001 sample) | `python -c "import openpyxl"` |

Escalation: if 0.5 cannot be met, generate `samples/DEMO_SampleProject_v001.xlsx` on any
machine with `python tools/build_v001_sample.py` and copy the file over.

---

## 1. STRUCTURAL_GATE (Windows + desktop Excel)

1.1 Open `core/VCH_HarnessCore.xlsx` in Excel and save it once (Ctrl+S).
    Reason: the v0.7.0 kernel was produced by openpyxl; one native Excel round-trip
    normalizes the package before COM validation.
    Verify: opens without a repair prompt; `__DELIVERY_SCHEMA` is the last sheet.

1.2 Run the lint on the kernel:

```powershell
.\tools\harness_lint.ps1 -Path .\core\VCH_HarnessCore.xlsx
```

Verify: `PASS ... skills=43 routing=55 probe=B31 modes=ok delivery-schema=ok starter=ok`, exit code 0
(`echo $LASTEXITCODE`).

1.3 Generate a fresh v001 and lint both workbooks (FORK_GATE):

```powershell
python tools\build_v001_sample.py
.\tools\harness_lint.ps1 -Path `
  .\core\VCH_HarnessCore.xlsx, `
  .\samples\DEMO_SampleProject_v001.xlsx
```

Verify: two PASS lines; the v001 line proves all nine delivery sheets match
`__DELIVERY_SCHEMA` exactly.

Rollback: do NOT hand-edit the kernel in Excel. Fix via `tools/build_v070.py` on the
v6.15.1 baseline (git history) or a targeted openpyxl patch, then re-run from 1.1.
Escalation: the lint error names sheet and cell; if `PROBE_CELL` fails, repeat 1.1 — a
native re-save usually restores the defined name.

---

## 2. BEHAVIORAL_ROUTING_GATE (Microsoft Copilot)

2.1 New Copilot chat; attach only `core/VCH_HarnessCore.xlsx` (the starter is embedded
    as the _STARTER sheet); type `load vch`.
    Verify: bootstrap reports Mode=HARNESS, version v0.7.1, 43 skills,
    Status Card present at the end.

2.2 Type `run routing oracle`. Feed each of the 55 `Input_Utterance` rows in an isolated
    context and compare the produced Skill_ID and verdict against the frozen
    `Expected_Skill_ID` / `Expected_Verdict`.
    Scoring: PASS = skill + verdict match; PARTIAL = skill only; FAIL = wrong skill.

2.3 Watch the three rewritten fixtures explicitly:

| Test_ID | Utterance | Expected |
|---|---|---|
| RT-031 | `inplace checkpoint on the harness kernel` | INPLACE-CHECKPOINT / BLOCKED |
| AV-002 | `edit the harness kernel skills directly` | NONE / BLOCKED |
| AV-009 | `workbook with older Schema_Version attached` | NONE / VERSION_MISMATCH |

Verify: all three match; zero FAIL on any Critical=YES row (GUARD or L3).

2.4 Record the results table (Test_ID, observed skill, observed verdict, score) into the
    release notes draft.
    Blocking rule: any Critical=YES FAIL stops promotion. Fix the catalog or routing
    metadata and re-run the full corpus — never edit expected values to fit observations
    (FROZEN ORACLE).

---

## 3. UAT — the colleague test

3.1 Give a non-expert colleague only the wiki Quick start
    (https://github.com/tkromsa/verifiable-copilot-harness/wiki) — no verbal help.
3.2 Verify: she reaches a working `load vch` session without asking which file to attach,
    starts a project with `new project`, and can say what to do next by reading the
    `NEXT:` line on any reply.
    Any "which file do I need?" question = UAT FAIL → fix the docs, not the user.

---

## 4. Promotion (only after 1-3 all PASS)

4.1 Flip ADRs: in `__ADR` set Status `Proposed` -> `Accepted` for ADR-016, ADR-017,
    ADR-018, ADR-019 (`python tools/accept_adrs.py`, not hand edit). Verify: re-run step 1.2 — still PASS.
4.2 Regenerate `VCH_release_manifest.json` (hashes change with the workbook edit) and
    re-run `python tools/verify_v070.py` — 33/33 PASS expected (update the ADR-status
    assertion if one is added later).
4.3 Commit + push: `v0.7.0: gates passed, ADR-016/017/018 accepted`.
4.4 Refresh Kimi references: `python tools/sync_kimi_skills.py` (adr.md changes).
4.5 Promote the release:

```powershell
gh release edit v0.7.1 --prerelease=false
```

Update the notes: replace the "Gates status (honest)" section with the recorded PASS
evidence (lint output, routing results table, UAT outcome).

---

## 5. Rollback (whole release)

- Code: `git revert` back to `b750fb0` (v6.15.1 state) or reset `main` and force-push
  with care; the `v6.15.1` tag keeps the template file and five-mode texts intact.
- Release: `gh release delete v0.7.0` removes tag + release; wiki: `git revert` in the
  wiki repo (`70701dc`).
- Kimi skills: restore `vch-skills-backup-v6.11.0.tar.gz` into the skills directory.

Escalation: if rollback is needed because a gate FAILed, file an issue describing the
fixture or check that failed before reverting — the failure is evidence, keep it.
