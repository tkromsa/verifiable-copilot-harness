# VCH Release Checklist

How to cut a new VCH release (e.g. v6.16.0) without creating inconsistencies between the workbooks, README, manifest, tag, release and wiki.

> This whole procedure can be delegated to an AI agent with GitHub CLI access (`gh auth login`, scope `repo`). Steps marked 🤖 are agent-friendly; steps marked 🪟 require Windows + desktop Excel.

---

## 0. Preconditions

- 🪟 Windows machine with desktop Excel (for the structural gate).
- `gh` CLI authenticated (`gh auth status` shows scope `repo`).
- Working tree: latest `main` pulled locally, or agent-side API access.

## 1. Make the change

- Apply the catalog / schema / rule change to the workbooks.
- Record an **ADR** in the workbook (`__ADR` sheet).
- Update `SKILLCOUNT` if skills were added or removed.
- Add or update `__ROUTING_ORACLE` fixtures for any new or changed trigger / precedence.

## 2. Structural gate 🪟

```powershell
.\tools\harness_lint.ps1 -Path `
  .\core\VCH_HarnessCore.xlsx, `
  .\core\VCH_ProjectTemplate.xlsx
```

- Must exit `0` with `PASS` on **both** workbooks.
- Any `FAIL` → fix and re-run. Do not proceed.

## 3. Behavioral routing gate 🤖

- Run all **55** `__ROUTING_ORACLE` fixtures in isolated contexts against the frozen `Expected_Skill_ID` / `Expected_Verdict`.
- Scoring: PASS = skill + verdict match; PARTIAL = skill only; FAIL = wrong skill.
- **Any FAIL on a Critical=YES GUARD or L3 row blocks the release.**
- Structural PASS never implies routing PASS.

## 4. Update version markers 🤖

- Bump the version **inside the artifacts** (filenames never change — ADR-010).
- README: update `**Current version: vX.Y.Z**` and add a one-line changelog bullet at the top of the list.

## 5. Rebuild the manifest 🤖

Every release file must have an up-to-date entry in `VCH_release_manifest.json` (sha256 + bytes). Order matters: hash the **final** file contents, after all edits above.

```bash
# macOS / Linux
for f in core/VCH_HarnessCore.xlsx core/VCH_ProjectTemplate.xlsx \
         copilot/copilotstart.txt copilot/copilot_custom_instructions.txt \
         tools/harness_lint.ps1 docs/VCH_Cheatsheet_EN.txt \
         docs/RELEASING.md README.md LICENSE; do
  echo "$f: $(shasum -a 256 "$f" | cut -d' ' -f1) $(stat -f %z "$f")"
done
```

```powershell
# Windows PowerShell
Get-ChildItem core\*.xlsx, copilot\*.txt, tools\*.ps1, docs\*.md, docs\*.txt, README.md, LICENSE |
  ForEach-Object { "{0}: {1} {2}" -f $_.FullName, (Get-FileHash $_ -Algorithm SHA256).Hash.ToLower(), $_.Length }
```

> ⚠️ If you edit a file **after** hashing it, re-hash. The classic failure is editing README last and shipping a stale README hash.

## 6. Commit 🤖

- Use a conventional message, never "Add files via upload":

```text
v6.16.0: <short change title> (ADR-0NN)
```

## 7. Tag & GitHub Release 🤖

```bash
gh release create v6.16.0 \
  --repo tkromsa/verifiable-copilot-harness \
  --target main \
  --title "v6.16.0 — <short change title>" \
  --notes "<notes, see template below>"
```

Release notes template:

```markdown
## What's new
<user-facing description of the change, ADR reference>

## Governance
- <SKILLCOUNT, routing corpus counts, ADR recorded, glossary terms>
- Deterministic scan on reopened workbooks: <lint results summary>

**Upgrade from vX.Y.Z:** replace all files, re-paste custom instructions.
sha256 of every file in `VCH_release_manifest.json`.
```

## 8. Update the wiki 🤖

- If user-facing behavior changed (new skill, new mode, new status value), update the affected wiki page(s): Usage Guide, Skills Reference, Release & Integrity.
- Wiki is a git repo: `https://github.com/tkromsa/verifiable-copilot-harness.wiki.git` (branch `master`).

## 9. Post-release verification 🤖

1. Download the release files fresh from GitHub.
2. Re-hash them and compare against `VCH_release_manifest.json`.
3. Confirm the repo description and README skill count / version agree with the workbooks.
4. Confirm the new tag points at the release commit.

```bash
curl -sL https://raw.githubusercontent.com/tkromsa/verifiable-copilot-harness/main/README.md | shasum -a 256
# must equal the README.md entry in the manifest
```

---

## Common failure modes this checklist prevents

| Failure | Prevented by step |
|---|---|
| README claims a version with no tag/release | 4 + 7 |
| Description / README skill count ≠ workbook SKILLCOUNT | 1 + 9 |
| Manifest hash stale after a late edit | 5 (order + warning) |
| Structural PASS treated as full verification | 3 |
| "Add files via upload" commit noise | 6 |
| Wiki drifting from user-facing behavior | 8 |
