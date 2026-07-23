# Verifiable Copilot Harness

**Turn Microsoft Copilot in Excel and Chat into a disciplined, evidence-driven project operator without requiring a custom agent runtime or paid add-on.**

Verifiable Copilot Harness (VCH) is a behavior harness for strict corporate environments. It provides 43 governed skills, frozen structural acceptance criteria, a deterministic lint gate, a frozen behavioral routing corpus, revision-controlled project workbooks, evidence-scoped persistence and conservative capability handling.

The harness is designed to prevent inferred state from being treated as verified state and to block persistence when the required evidence or runtime capability is unavailable.

📖 **Full documentation: [the Wiki](https://github.com/tkromsa/verifiable-copilot-harness/wiki)** — usage guide, all 43 skills, release gates and integrity checks.

**Current version: v6.15.1**

- v6.15.1 hardens the release gate, removes the stale active B32 probe instruction, restores historical ADR release identity and aligns all active routing counts to 55.
- v6.15.0 adds the `Allowed_Modes` column for deterministic mode filtering (ADR-014).
- v6.14.0 adds `STATUS`, `CAPABILITY-DISCOVERY` and the Status Card (ADR-013).
- v6.13.0 adds the embedded `__ROUTING_ORACLE` behavioral corpus (ADR-012).
- v6.12.0 adds one-active-workbook modes (ADR-011).
- Filenames are version-independent (ADR-010). The version lives inside the artifacts and is reported by `load vch`.

## Why it exists

Out of the box, a chat model can:

- infer workbook state that was never read back,
- claim persistence without reopening the output,
- lose artifact lineage across revisions,
- confuse stored capability evidence with current runtime capability,
- route similar requests inconsistently.

VCH addresses these risks with explicit rules and verifiable artifacts.

| Rule | What it enforces |
|---|---|
| Read-only bootstrap | Every session starts with a complete read-only load. Nothing is created or saved implicitly. |
| One active workbook | Normal operation uses one active workbook. Compatible duplicate catalogs are validated but read once, never merged. |
| No capability inheritance | Stored capability evidence never proves the current runtime capability. |
| Frozen structural oracle | Acceptance criteria in `__TEST_ORACLE` do not change during a run. |
| Frozen routing corpus | `__ROUTING_ORACLE` contains 55 L2/L3 routing and adversarial fixtures and is excluded from the bootstrap read-set. |
| Evidence classes | Independent read-back, tool reports, conversation claims, simulated input and inference remain distinct. |
| Persistence router | Persistent change requires an evidenced persistence mode. Otherwise the result is `READ_ONLY` or `WRITE_BLOCKED`. |
| SCRUB | Sensitive output must pass deterministic sanitization before crossing a trust boundary. |
| Status Card | Key lifecycle events end with Mode, Resume, Gate, Verified and Next. |
| Deterministic lint | `tools/harness_lint.ps1` performs the structural checks required for release promotion. |

## Repository layout

```text
core/
  VCH_HarnessCore.xlsx
  VCH_ProjectTemplate.xlsx
copilot/
  copilotstart.txt
  copilot_custom_instructions.txt
tools/
  harness_lint.ps1
docs/
  VCH_Cheatsheet_EN.txt
  RELEASING.md
VCH_release_manifest.json
```

The validator is a plain Windows PowerShell script using desktop Excel via COM. No modules or add-ons are required.

## Workbook modes

Normal operation uses one active workbook.

| Mode | Active file | Use |
|---|---|---|
| HARNESS | `VCH_HarnessCore.xlsx` | Everyday non-project work |
| TEMPLATE | `VCH_ProjectTemplate.xlsx` | Start a project directly |
| PROJECT-CREATION | HarnessCore plus Template | Convert non-project work into a governed project |
| PROJECT | `<ID>_<Name>_vNNN.xlsx` | Continue a self-contained project |
| MIGRATION | Old project plus current Template | Port verified content to the current schema |

Multiple workbooks are used only for a controlled transition, migration, lineage verification or explicit comparison.

## Quick start

### Everyday non-project work

1. Attach `core/VCH_HarnessCore.xlsx` and `copilot/copilotstart.txt`.
2. Type `load vch`.
3. Ask normally or use an explicit skill trigger.

### Start a project directly

1. Attach `core/VCH_ProjectTemplate.xlsx` and `copilot/copilotstart.txt`.
2. Type `load vch`.
3. Use `[skill: PROJECT-FORK]` with Project ID, Project Name, Owner and Main Goal.

### Convert current harness work into a project

1. Keep `VCH_HarnessCore.xlsx` available and attach `VCH_ProjectTemplate.xlsx`.
2. Use `[skill: PROJECT-FORK]` with complete project metadata.
3. After revision 001 is reopened and verified, continue with the new project workbook only.

### Continue an existing project

1. Attach the current verified project workbook and only the supporting files needed for the task.
2. Type `load vch`.
3. Type `guide project` or invoke a specialist skill directly.

## Everyday commands

| You type | Skill | Result |
|---|---|---|
| `load vch` | LOAD-ALL | Complete read-only bootstrap |
| `guide project` / `what next` | PROJECT-GUIDE | Next governed project step |
| `new project` | PROJECT-FORK | Create revision 001 from the template |
| `write plan` | WRITING-PLANS | Implementation plan with VERIFY evidence |
| `debug this` | SYSTEMATIC-DEBUGGING | Root cause before fix |
| `review code` | CODE-REVIEW | Findings without silent rewrite |
| `verify fix` | VERIFICATION-CHECKLIST | Evidence-based verification |
| `status` / `where are we` | STATUS | Read-only Status Card |
| `what can you do` | CAPABILITY-DISCOVERY | Skills currently allowed by mode, gates and capability |
| `focus mode` | FOCUS-MODE | Action-first output shaping |

## Status Card

After key lifecycle events, VCH emits a compact final block:

```text
STATUS
Mode: HARNESS|TEMPLATE|PROJECT-CREATION|PROJECT|MIGRATION
Resume: <literal __STATE Resume_From value>
Gate: <gate> -> PASS|BLOCKED|NOT_TESTED|n/a
Verified: <up to three evidence-qualified items>
Next: <one concrete action>
```

Conversation claims are never listed as verified evidence.

## Allowed modes

`00_Skills.Allowed_Modes` is a machine-readable allowlist. Each value is either:

```text
ALL
```

or a canonical comma-space separated subset of:

```text
HARNESS, TEMPLATE, PROJECT-CREATION, PROJECT, MIGRATION
```

Routing and `CAPABILITY-DISCOVERY` must not propose a skill outside the current workbook mode. The deterministic lint validates the column and rejects empty values, unknown modes, combined `ALL`, duplicate modes and non-canonical separator spacing.

## Routing oracle

Both workbooks embed a frozen `__ROUTING_ORACLE` sheet with 55 fixtures:

- 43 L2 routing-correctness fixtures,
- 12 L3 adversarial and safe-failure fixtures,
- explicit fixture context and candidate skill sets for synthetic precedence ties.

The routing oracle is intentionally excluded from `load vch`.

### Structural gate

The lint validates:

- routing sheet presence,
- exactly 55 rows,
- unique `Test_ID` values,
- resolvable `Expected_Skill_ID` values,
- complete tie fixtures.

### Behavioral routing gate

A behavioral run feeds each `Input_Utterance` to the model in an isolated context and compares the observed skill and verdict with the frozen `Expected_Skill_ID` and `Expected_Verdict`.

Structural PASS never implies routing PASS. A mismatch on a critical guard or L3 row blocks catalog promotion.

## Persistence

| Mode | Requirement |
|---|---|
| `READ_ONLY` | Default when persistence is not requested or capability is unproven |
| `ARTIFACT_COPY_ON_WRITE` | Create, write, export, reopen and verify a new physical revision |
| `PRECOPIED_ARTIFACT` | Activate the exact user-created next revision after verifying identity and parentage |
| `EXCEL_NATIVE_VERSIONING` | In-place logical checkpoint in a verified `OPEN_WORKBOOK` host |
| `WRITE_BLOCKED` | Required persistence capability or prerequisite is not proven |

HarnessCore and ProjectTemplate are immutable for project delivery.

## Release gates

A catalog or schema change is promoted only when both gates pass.

### 1. Structural gate

Run the validator on Windows with desktop Excel installed:

```powershell
.\tools\harness_lint.ps1 -Path `
  .\core\VCH_HarnessCore.xlsx, `
  .\core\VCH_ProjectTemplate.xlsx
```

The validator continues with the second workbook even if the first fails and exits nonzero when any result is `FAIL`.

### 2. Behavioral routing gate

Run the 55 routing fixtures in isolated contexts against the frozen expected values. The behavioral runner is separate from the structural lint.

## Migrating an old project

Never rewrite an old-version project workbook.

1. Load the old workbook read-only and report `VERSION_MISMATCH`.
2. Fork from the current template.
3. Port verified project content by field or column identity.
4. Never transfer stored runtime capability PASS state.
5. Reopen and verify identity, lineage and intended values.
6. Archive the old project line after the new revision is promoted.

## Trust boundaries and limitations

- A workbook instruction constrains model behavior but cannot independently guarantee model compliance.
- Structural lint does not prove behavioral routing.
- Conversation claims are not independent read-back.
- A manifest generated by the build process is release metadata, not independent verification.
- Direct editing and persistence capabilities depend on live host evidence.

## Inspiration and credits

VCH is an independent implementation for Microsoft Copilot and Excel, informed by open skill and human-in-the-loop patterns, including:

- [obra/superpowers](https://github.com/obra/superpowers)
- [mattpocock/skills - Grill Me](https://github.com/mattpocock/skills/blob/main/skills/productivity/grill-me/SKILL.md)
- [rjs/shaping-skills](https://github.com/rjs/shaping-skills)
- [Gentleman-Programming/gentle-ai](https://github.com/Gentleman-Programming/gentle-ai)
- [ayghri/i-have-adhd](https://github.com/ayghri/i-have-adhd)

## License

MIT. See [LICENSE](LICENSE).
