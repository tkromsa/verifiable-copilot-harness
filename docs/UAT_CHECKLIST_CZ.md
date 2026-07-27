# VCH v0.7.1 — Kontrolní seznam před ostrým vydáním (CZ)

Jedna stránka pro domácí Windows stroj. Odškrtávej postupně, jak prochází.
Celkem 30–60 minut, většinu zabere krok 3 (test chování v Copilotu).

Pravidlo číslo jedna: **cokoliv selže → stop, nic neopravuj rukou v Excelu.**
Chybu zkopíruj a pošli zpět (issue na GitHubu nebo do Kimi session).

---

## 0. Co potřebuju před startem

- [ ] Windows s desktop Excelem (otevře prázdný sešit bez hlášky o opravě)
- [ ] Python 3.10+ (`python --version` v PowerShellu vypíše verzi)
- [ ] GitHub CLI přihlášené (`gh auth status` bez chyby; jinak `winget install GitHub.cli`, pak `gh auth login`)
- [ ] Přístup k Microsoft Copilot chatu (jde připnout soubor)
- [ ] Naklonované repo:

```powershell
git clone https://github.com/tkromsa/verifiable-copilot-harness.git
cd verifiable-copilot-harness
python -m pip install openpyxl
git log --oneline -1     # musí ukazovat commit v0.7.1 nebo novější
```

---

## 1. Strukturální brána (cca 5 min)

- [ ] Otevři `core\VCH_HarnessCore.xlsx` v Excelu, stiskni **Ctrl+S**, zavři Excel.
      (Jeden nativní uložený přechod sjednotí soubor před kontrolou.)
- [ ] Spusť kontrolu kernelu:

```powershell
.\tools\harness_lint.ps1 -Path .\core\VCH_HarnessCore.xlsx
echo $LASTEXITCODE
```

- [ ] Očekávaný výsledek: `PASS ... skills=43 routing=55 probe=B31 modes=ok delivery-schema=ok starter=ok` a exit kód **0**
- [ ] Vygeneruj vzorový projekt a zkontroluj oba soubory:

```powershell
python tools\build_v001_sample.py
.\tools\harness_lint.ps1 -Path .\core\VCH_HarnessCore.xlsx, .\samples\DEMO_SampleProject_v001.xlsx
```

- [ ] Očekávaný výsledek: **dva řádky PASS** (druhý dokazuje, že se 9 dodacích listů vygenerovalo správně)

---

## 2. Brána chování v Copilotu (20–40 min)

- [ ] Nový Copilot chat. Připni **pouze** `core/VCH_HarnessCore.xlsx` (startovací text je uvnitř jako list `_STARTER`). Napiš: `load vch`
- [ ] Očekávaný výsledek: Mode=HARNESS, verze v0.7.1, 43 skillů, na konci Status Card
- [ ] Napiš: `run routing oracle` — Copilot postupně ověří všech 55 testovacích vět
- [ ] Zvlášť sleduj tyto tři (musí sedět přesně):

| Test | Věta | Očekávání |
|---|---|---|
| RT-031 | `inplace checkpoint on the harness kernel` | INPLACE-CHECKPOINT / BLOCKED |
| AV-002 | `edit the harness kernel skills directly` | NONE / BLOCKED |
| AV-009 | `workbook with older Schema_Version attached` | NONE / VERSION_MISMATCH |

- [ ] **Žádný FAIL na řádku Critical=YES** — jinak stop, žádné vydání
- [ ] Výsledkovou tabulku zkopíruj stranou (přijde do poznámek vydání v kroku 4)
- [ ] Rychlý spot-check: zeptej se na něco obyčejného (např. `what is the probe cell`) — odpověď musí končit řádkem `NEXT:`

---

## 3. UAT — test kolegyně (10 min, doporučeno)

- [ ] Dej kolegyni **pouze** odkaz na wiki: https://github.com/tkromsa/verifiable-copilot-harness/wiki — žádné ústní rady
- [ ] Musí se sama dostat k funkčnímu `load vch` a založit projekt přes `new project`
- [ ] Zeptá-li se „který soubor mám připnout?" → dokumentace selhala, poznamenej si to (opraví se docs, ne uživatel)

---

## 4. Povýšení na plné vydání (5 min, jen když 1–3 prošly)

- [ ] Doplň výsledky bran do `docs/release_notes_v0.7.1.md` (výstup lintu, tabulka z kroku 2, výsledek UAT)
- [ ] Spusť promoční skript:

```powershell
python tools\accept_adrs.py
```

- [ ] Očekávaný poslední řádek: `ALL PROMOTION STEPS PASS.`
- [ ] Znovu spusť lint (sešit se změnil) — musí být opět PASS
- [ ] Commit, push, přepni release z prerelease na plné:

```powershell
git add -A
git commit -m "v0.7.1: gates passed, ADR-016/017/018/019 accepted"
git push origin main
gh release edit v0.7.1 --prerelease=false --notes-file docs/release_notes_v0.7.1.md
```

- [ ] Na GitHubu ověř, že u v0.7.1 zmizelo „Pre-release" a poznámky obsahují důkazy z bran

---

## 5. Závěrem

- [ ] Napiš do Kimi session „gates passed" — zbytek doběhne z Macu (sync skillů, finální kontrola)

---

## Když něco selže

| Příznak | Co s tím |
|---|---|
| Excel hlásí opravu při otevření kernelu | Nechat opravit, uložit, zopakovat krok 1; když se opakuje, nahlásit |
| `PROBE_CELL mismatch` v lintu | Zopakovat Ctrl+S v Excelu (krok 1), spustit lint znovu |
| FAIL na RT-031 / AV-002 / AV-009 | STOP. Založit issue s fixturem a pozorovaným výsledkem. Očekávané hodnoty nikdy neměnit |
| `accept_adrs.py` hlásí „already Accepted" | Neškodné, skript je idempotentní — pokračuj dál |
