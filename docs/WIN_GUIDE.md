# VCH v0.7.1 — Windows Home Guide (gates + promotion)

Complete self-service procedure for the Windows machine. Everything runs from the repo
root in PowerShell. Expected total time: 30-60 minutes, most of it the behavioral gate.

Related docs: `docs/RUNBOOK_gates.md` (full runbook with rollback),
`docs/PLAN_v0.7.0.md` (project state).

---

## A. One-time setup (5 min)

```powershell
cd <wherever-you-clone>
git clone https://github.com/tkromsa/verifiable-copilot-harness.git
cd verifiable-copilot-harness
git log --oneline -1        # must show the v0.7.1 commit or newer
python --version            # any Python 3.10+
python -m pip install openpyxl
```

Verify: `python -c "import openpyxl"` prints nothing and exits cleanly.

---

## B. Structural gate (5 min)

B.1 Open `core\VCH_HarnessCore.xlsx` in desktop Excel, press Ctrl+S, close Excel.
    (One native round-trip normalizes the openpyxl-produced file before COM validation.)

B.2 Lint the kernel:

```powershell
.\tools\harness_lint.ps1 -Path .\core\VCH_HarnessCore.xlsx
echo $LASTEXITCODE     # must be 0
```

Expected: `PASS ... skills=43 routing=55 probe=B31 modes=ok delivery-schema=ok starter=ok`

B.3 Fork gate — generate a v001 and lint both:

```powershell
python tools\build_v001_sample.py
.\tools\harness_lint.ps1 -Path .\core\VCH_HarnessCore.xlsx, .\samples\DEMO_SampleProject_v001.xlsx
```

Expected: two PASS lines. The second line proves the nine delivery sheets were
generated correctly from `__DELIVERY_SCHEMA`.

If anything FAILs: stop, paste the error into an issue or send it back. Do not hand-edit
the workbook.

---

## C. Behavioral routing gate (20-40 min, Microsoft Copilot)

C.1 New Copilot chat. Attach only `core/VCH_HarnessCore.xlsx` (the starter is embedded
    as the `_STARTER` sheet). Type: `load vch`
    Expected: bootstrap reports Mode=HARNESS, version v0.7.1, 43 skills, ends with a
    Status Card whose Next line contains `| Skill: ... | Why: ...`.

C.2 Type: `run routing oracle`
    Copilot feeds each of the 55 `Input_Utterance` rows in an isolated context and
    compares the result with the frozen expected values.
    Scoring: PASS = skill + verdict match, PARTIAL = skill only, FAIL = wrong skill.

C.3 Explicitly confirm the three rewritten fixtures:

| Test_ID | Utterance | Expected |
|---|---|---|
| RT-031 | `inplace checkpoint on the harness kernel` | INPLACE-CHECKPOINT / BLOCKED |
| AV-002 | `edit the harness kernel skills directly` | NONE / BLOCKED |
| AV-009 | `workbook with older Schema_Version attached` | NONE / VERSION_MISMATCH |

Blocking rule: any FAIL on a Critical=YES row stops promotion. Copy the full results
table — you will paste it into the release notes in step E.

C.4 While in the chat, spot-check the NEXT footer: ask an ordinary question
    (e.g. `what is the probe cell`) and confirm the reply ends with a `NEXT:` line.

---

## D. UAT (optional but recommended, 10 min)

Give a colleague only the wiki link: https://github.com/tkromsa/verifiable-copilot-harness/wiki
She must reach a working `load vch` session and start a project without asking which
file to attach. Any "which file do I need?" question = docs FAIL; note it and fix later.

---

## E. Promotion (5 min, after B-D all PASS)

E.1 Update the gates section in `docs/release_notes_v0.7.1.md` with your evidence
    (lint output, routing results, UAT outcome).

E.2 Run the promotion script:

```powershell
python tools\accept_adrs.py
```

It flips ADR-016/017/018/019 to Accepted, regenerates the manifest and re-verifies the
workbook. Expected last line: `ALL PROMOTION STEPS PASS.`

E.3 Re-run the lint once more (the workbook changed):

```powershell
.\tools\harness_lint.ps1 -Path .\core\VCH_HarnessCore.xlsx
```

E.4 Commit, push, promote:

```powershell
git add -A
git commit -m "v0.7.1: gates passed, ADR-016/017/018/019 accepted"
git push origin main
gh release edit v0.7.1 --prerelease=false --notes-file docs/release_notes_v0.7.1.md
```

E.5 Verify on GitHub: the release page no longer shows "Pre-release" and the notes
    contain your gate evidence.

---

## F. Afterwards

Tell the Kimi session "gates passed" — the remaining steps run from the Mac:
`tools/sync_kimi_skills.py` refresh (adr.md changes) and a final end-to-end check.

If you prefer, stop after E.2 and hand everything back; E.3-E.5 can be done remotely.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Excel repair dialog when opening the kernel | openpyxl artifact | Let Excel repair, save, re-run B.2; report if it repeats |
| `PROBE_CELL mismatch` in lint | defined name lost | Repeat B.1 (native re-save), re-run |
| `python` not found | Python not installed | Install python.org 3.12, tick "Add to PATH", reopen PowerShell |
| `gh` not found | GitHub CLI missing | `winget install GitHub.cli`, then `gh auth login` |
| Routing FAIL on RT-031/AV-002/AV-009 | real routing drift | Stop. File an issue with the fixture and observed result. Do not edit expected values |
| `accept_adrs.py` says "already Accepted" | script ran before | Harmless; it is idempotent |
