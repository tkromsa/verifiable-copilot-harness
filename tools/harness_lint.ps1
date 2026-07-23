<#
.SYNOPSIS
    harness-lint v6.11.0 - executable implementation of every DETERMINISTIC_SCAN
    check declared in __TEST_ORACLE of the VCH harness workbooks.

.DESCRIPTION
    Read-only validator for VCH_HarnessCore and VCH_ProjectTemplate
    workbooks. Executes the frozen-oracle deterministic checks so that
    DETERMINISTIC_SCAN is real code, not a label:

      SKILLCOUNT   Skill count equals ExpectedSkillCount
      VERSION      All declared version fields agree; no foreign bare version cells
      NORMALIZATION Trigger duplicates / prefix collisions / non-ASCII cells /
                   duplicate identity keys / canonical NONE token / Current_Artifact
      CHAINS       Every May_Chain_To Skill_ID resolves
      ENUMS        Every value of a Lists-declared field is declared in Lists
      STATE        Probe cell field __WRITE_PROBE at row 32; PROBE_CELL named range;
                   Failed_Artifact empty in a clean artifact; all canonical
                   __STATE fields incl. the ADR-008 lifecycle extension present
      GUIDE        Next_Recommended_Skill is NONE or an existing Skill_ID

.ASSUMPTIONS
    - PowerShell 5.1+ on Windows.
    - Module ImportExcel installed (Install-Module ImportExcel -Scope CurrentUser).
      Its bundled EPPlus assembly is used for full-fidelity cell access.
    - Workbook schema: 00_Skills, Project.Rules, __STATE, Lists, __TEST_ORACLE,
      __ADR, __GLOSSARY as defined by harness v6.11.0.

.NOTES
    Safety: read-only. No writes, no secrets, no rollback needed.
    Evidence: emits one record per check (PASS/FAIL + detail) - TOOL_TRANSACTION_REPORT.
    Exit code: 0 = all PASS, 1 = at least one FAIL or unreadable file.

.EXAMPLE
    .\harness_lint.ps1 -Path .\core\VCH_HarnessCore.xlsx, .\core\VCH_ProjectTemplate.xlsx
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$Path,
    [string]$ExpectedVersion = 'v6.11.0',
    [int]$ExpectedSkillCount = 41,
    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- prerequisite: ImportExcel (provides EPPlus) -------------------------------
$mod = Get-Module -ListAvailable -Name ImportExcel | Sort-Object Version -Descending | Select-Object -First 1
if (-not $mod) {
    throw "Module ImportExcel not found. Install once: Install-Module ImportExcel -Scope CurrentUser"
}
$epplus = Get-ChildItem -Path $mod.ModuleBase -Recurse -Filter 'EPPlus.dll' | Select-Object -First 1
if (-not $epplus) { throw "EPPlus.dll not found inside ImportExcel module at $($mod.ModuleBase)" }
Add-Type -Path $epplus.FullName
Write-Verbose "Using EPPlus from $($epplus.FullName)"

$results = New-Object System.Collections.Generic.List[object]

function Add-Result {
    param([string]$File, [string]$TestId, [string]$Check, [bool]$Ok, [string]$Detail)
    $script:results.Add([pscustomobject]@{
        File    = $File
        Test_ID = $TestId
        Check   = $Check
        Result  = $(if ($Ok) { 'PASS' } else { 'FAIL' })
        Detail  = $Detail
    })
}

function Get-Sheet {
    param($Package, [string]$Name)
    $ws = $Package.Workbook.Worksheets[$Name]
    if (-not $ws) { throw "Missing sheet '$Name'" }
    return $ws
}

function Get-Rows {
    # Returns non-empty rows as string[] arrays (cells trimmed of nothing; nulls preserved)
    param($Worksheet)
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $Worksheet.Dimension) { return $rows }
    for ($r = $Worksheet.Dimension.Start.Row; $r -le $Worksheet.Dimension.End.Row; $r++) {
        $vals = @()
        $hasValue = $false
        for ($c = $Worksheet.Dimension.Start.Column; $c -le $Worksheet.Dimension.End.Column; $c++) {
            $v = $Worksheet.Cells[$r, $c].Value
            if ($null -ne $v -and "$v" -ne '') { $hasValue = $true }
            $vals += $v
        }
        if ($hasValue) { $rows.Add([pscustomobject]@{ Row = $r; Cells = $vals }) }
    }
    return $rows
}

function Find-HeaderRow {
    param($Rows, [string]$FirstColumnName)
    foreach ($row in $Rows) { if ($row.Cells[0] -eq $FirstColumnName) { return $row.Row } }
    return 0
}

foreach ($file in $Path) {
    $fname = Split-Path $file -Leaf
    Write-Verbose "Linting $fname"
    try {
        if (-not (Test-Path $file)) { throw "File not found: $file" }
        $fi = Get-Item $file
        $pkg = New-Object OfficeOpenXml.ExcelPackage($fi)
        try {
            # ---- gather sheet data -------------------------------------------------
            $wsSkills = Get-Sheet $pkg '00_Skills'
            $wsState  = Get-Sheet $pkg '__STATE'
            $wsLists  = Get-Sheet $pkg 'Lists'
            $wsOracle = Get-Sheet $pkg '__TEST_ORACLE'

            $skillRows = Get-Rows $wsSkills
            $hRow = Find-HeaderRow $skillRows 'Skill_ID'
            if ($hRow -eq 0) { throw "00_Skills header row not found" }
            $hdr = ($skillRows | Where-Object Row -eq $hRow).Cells
            $colTrigger = [Array]::IndexOf($hdr, 'Trigger')
            $colChain   = [Array]::IndexOf($hdr, 'May_Chain_To')
            $colVersion = [Array]::IndexOf($hdr, 'Version')
            $skills = $skillRows | Where-Object { $_.Row -gt $hRow -and $_.Cells[0] }
            $skillIds = @($skills | ForEach-Object { "$($_.Cells[0])" })

            # ---- SKILLCOUNT --------------------------------------------------------
            Add-Result $fname 'SKILLCOUNT' 'Skill_Count' ($skills.Count -eq $ExpectedSkillCount) "found $($skills.Count), expected $ExpectedSkillCount"

            # ---- VERSION: skill versions + no foreign bare version cells -----------
            $badVer = @($skills | Where-Object { "$($_.Cells[$colVersion])" -ne $ExpectedVersion } | ForEach-Object { "$($_.Cells[0])" })
            Add-Result $fname 'VERSION' 'Skill_Version_Format' ($badVer.Count -eq 0) ($badVer -join ',')
            $foreign = New-Object System.Collections.Generic.List[string]
            foreach ($ws in $pkg.Workbook.Worksheets) {
                if (-not $ws.Dimension) { continue }
                foreach ($cell in $ws.Cells[$ws.Dimension.Address]) {
                    $v = $cell.Value
                    if ($v -is [string] -and $v.Trim() -match '^v\d+\.\d+(\.\d+)?$' -and $v.Trim() -ne $ExpectedVersion) {
                        $foreign.Add("$($ws.Name)!$($cell.Address)=$v")
                    }
                }
            }
            Add-Result $fname 'VERSION' 'Harness_Version_Agreement' ($foreign.Count -eq 0) ($foreign -join ';')

            # ---- NORMALIZATION: trigger duplicates ---------------------------------
            $atoms = @{}
            foreach ($s in $skills) {
                foreach ($a in ("$($s.Cells[$colTrigger])" -split ' / ')) {
                    $key = $a.Trim().ToLowerInvariant()
                    if (-not $atoms.ContainsKey($key)) { $atoms[$key] = New-Object System.Collections.Generic.List[string] }
                    $atoms[$key].Add("$($s.Cells[0])")
                }
            }
            $dups = @($atoms.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 } | ForEach-Object { "$($_.Key) <- $($_.Value -join ',')" })
            Add-Result $fname 'NORMALIZATION' 'Trigger_Duplicates' ($dups.Count -eq 0) ($dups -join ';')

            # ---- NORMALIZATION: prefix collisions ----------------------------------
            $keys = @($atoms.Keys)
            $pref = New-Object System.Collections.Generic.List[string]
            foreach ($a in $keys) { foreach ($b in $keys) {
                if ($a -ne $b -and $b.StartsWith($a)) { $pref.Add("'$a' prefix of '$b'") }
            } }
            Add-Result $fname 'NORMALIZATION' 'Trigger_Prefix_Collisions' ($pref.Count -eq 0) ($pref -join ';')

            # ---- NORMALIZATION: hidden / non-ASCII characters ----------------------
            $badCells = New-Object System.Collections.Generic.List[string]
            foreach ($ws in $pkg.Workbook.Worksheets) {
                if (-not $ws.Dimension) { continue }
                foreach ($cell in $ws.Cells[$ws.Dimension.Address]) {
                    $v = $cell.Value
                    if ($v -is [string] -and $v.Length -gt 0) {
                        foreach ($ch in $v.ToCharArray()) {
                            $code = [int]$ch
                            if ($code -gt 126 -or ($code -lt 32 -and $ch -notin "`n","`r","`t") -or
                                [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -eq 'Format' -or $code -eq 160) {
                                $badCells.Add("$($ws.Name)!$($cell.Address) U+$('{0:X4}' -f $code)"); break
                            }
                        }
                    }
                }
            }
            Add-Result $fname 'NORMALIZATION' 'Hidden_Or_Non_ASCII_Characters' ($badCells.Count -eq 0) (($badCells | Select-Object -First 5) -join ';')

            # ---- NORMALIZATION: canonical NONE token -------------------------------
            $oracleRows = Get-Rows $wsOracle
            $oHRow = Find-HeaderRow $oracleRows 'Test_ID'
            $noneViol = @($oracleRows | Where-Object { $_.Row -gt $oHRow -and ($_.Cells[4] -eq 'None' -or $_.Cells[4] -eq 'none') } | ForEach-Object { "oracle row $($_.Row)" })
            Add-Result $fname 'NORMALIZATION' 'Canonical_Token_Violations' ($noneViol.Count -eq 0) ($noneViol -join ';')

            # ---- NORMALIZATION: duplicate identity keys -----------------------------
            $dupList = New-Object System.Collections.Generic.List[string]
            foreach ($sid in ($skillIds | Group-Object | Where-Object Count -gt 1)) { $dupList.Add("Skill_ID:$($sid.Name)") }
            foreach ($sheetKey in @(@('Project.Rules','Rule'), @('__GLOSSARY','Term'), @('__ADR','ADR_ID'))) {
                $ws = $pkg.Workbook.Worksheets[$sheetKey[0]]
                if (-not $ws) { continue }
                $rowsK = Get-Rows $ws
                $hk = Find-HeaderRow $rowsK $sheetKey[1]
                if ($hk -eq 0) { continue }
                $vals = @($rowsK | Where-Object { $_.Row -gt $hk -and $_.Cells[0] } | ForEach-Object { "$($_.Cells[0])" })
                foreach ($g in ($vals | Group-Object | Where-Object Count -gt 1)) { $dupList.Add("$($sheetKey[0]):$($g.Name)") }
            }
            $oIds = @($oracleRows | Where-Object { $_.Row -gt $oHRow -and $_.Cells[0] } | ForEach-Object { "$($_.Cells[0])+$($_.Cells[1])+$($_.Cells[2])" })
            foreach ($g in ($oIds | Group-Object | Where-Object Count -gt 1)) { $dupList.Add("oracle:$($g.Name)") }
            Add-Result $fname 'NORMALIZATION' 'Duplicate_Identity_Keys' ($dupList.Count -eq 0) ($dupList -join ';')

            # ---- NORMALIZATION: Current_Artifact equals physical filename -----------
            $stateRows = Get-Rows $wsState
            $state = @{}
            foreach ($r in $stateRows) { if ($r.Cells[0]) { $state["$($r.Cells[0])"] = $r.Cells[1] } }
            $caState = "$($state['Current_Artifact'])"
            $caOracle = @($oracleRows | Where-Object { $_.Row -gt $oHRow -and $_.Cells[2] -eq 'Current_Artifact' } | ForEach-Object { "$($_.Cells[3])" })
            $caOk = ($caState -eq $fname) -and ($caOracle.Count -eq 1) -and ($caOracle[0] -eq $fname)
            Add-Result $fname 'NORMALIZATION' 'Current_Artifact' $caOk "state=$caState oracle=$($caOracle -join ',') file=$fname"

            # ---- CHAINS: unresolved references --------------------------------------
            $badChains = New-Object System.Collections.Generic.List[string]
            foreach ($s in $skills) {
                $chain = "$($s.Cells[$colChain])"
                if ($chain) {
                    foreach ($t in ($chain -split ', ')) {
                        if ($skillIds -notcontains $t.Trim()) { $badChains.Add("$($s.Cells[0])->$t") }
                    }
                }
            }
            Add-Result $fname 'CHAINS' 'Unresolved_Chain_References' ($badChains.Count -eq 0) ($badChains -join ';')

            # ---- ENUMS: values of Lists-declared fields ------------------------------
            $listRows = Get-Rows $wsLists
            $lHRow = Find-HeaderRow $listRows 'Host_Mode'
            $lHdr = ($listRows | Where-Object Row -eq $lHRow).Cells
            $declared = @{}
            for ($c = 0; $c -lt $lHdr.Count; $c++) {
                if ($lHdr[$c]) {
                    $declared["$($lHdr[$c])"] = @($listRows | Where-Object { $_.Row -gt $lHRow -and $_.Cells[$c] } | ForEach-Object { "$($_.Cells[$c])" })
                }
            }
            $enumBad = New-Object System.Collections.Generic.List[string]
            foreach ($r in $stateRows) {
                $f = "$($r.Cells[0])"; $v = $r.Cells[1]
                if ($declared.ContainsKey($f) -and $null -ne $v -and "$v" -ne '' -and $declared[$f] -notcontains "$v") {
                    $enumBad.Add("__STATE $f=$v")
                }
            }
            $wsLand = $pkg.Workbook.Worksheets['00_Landing']
            if ($wsLand) {
                foreach ($r in (Get-Rows $wsLand)) {
                    for ($c = 0; $c -lt $r.Cells.Count; $c++) {
                        $f = "$($r.Cells[$c])"
                        if ($declared.ContainsKey($f)) {
                            $v = $null
                            for ($d = $c + 1; $d -lt $r.Cells.Count; $d++) { if ($r.Cells[$d]) { $v = "$($r.Cells[$d])"; break } }
                            if ($v -and $v -notmatch '^(PROBE_CELL|__STATE)' -and $declared[$f] -notcontains $v) { $enumBad.Add("00_Landing $f=$v") }
                        }
                    }
                }
            }
            Add-Result $fname 'ENUMS' 'Undeclared_Enum_Values' ($enumBad.Count -eq 0) ($enumBad -join ';')

            # ---- STATE: probe cell + named range + Failed_Artifact -------------------
            $a32 = $wsState.Cells[32, 1].Text; $b32 = $wsState.Cells[32, 2].Text
            Add-Result $fname 'STATE' 'Probe_Cell_Field' ($a32 -eq '__WRITE_PROBE' -and $b32 -eq '') "A32=$a32 B32=$b32"
            $probeName = $pkg.Workbook.Names['PROBE_CELL']
            $probeOk = ($null -ne $probeName) -and ($probeName.Address -match '\$B\$32')
            Add-Result $fname 'STATE' 'PROBE_CELL_Named_Range' $probeOk $(if ($probeName) { $probeName.FullAddress } else { 'missing' })
            $fa = "$($state['Failed_Artifact'])"
            Add-Result $fname 'STATE' 'Failed_Artifact_Empty' ($fa -eq '') "Failed_Artifact='$fa'"

            # ---- STATE: canonical schema fields (incl. ADR-008 lifecycle extension) ---
            $requiredFields = @('Harness version','File_Role','Template_Version','Project_ID','Project_Safe_Name',
                'Revision','Logical_Revision','Revision_Type','Parent_Artifact','Current_Artifact','Active_Artifact',
                'Physical_Artifact','Physical_Artifact_Changed','Artifact_Status','Created_From','Last_Export_Verified',
                'Output_Verification_Status','Host_Mode','Persistence_Mode','Storage_Version_ID','Native_Version_History_Status',
                'Precreated_Target_Status','Last_Completed_Action','Last_Attempted_Action','Last_Failed_Action',
                'Last_Verified_Action','Failed_Artifact','__WRITE_PROBE','Failure_Stage','Failure_Code','Current_Skill',
                'Current_Phase','Next step','Resume_From','State_Status','Capability_Mode','Source_Refs',
                'Phase_Gate_Status','Selected_Shape','Current_Plan_ID','Last_Verified_Gate','Next_Recommended_Skill','Next_Executable_Step')
            $missingFields = @($requiredFields | Where-Object { -not $state.ContainsKey($_) })
            Add-Result $fname 'STATE' 'Missing_Required_State_Fields' ($missingFields.Count -eq 0) ($missingFields -join ',')

            # ---- GUIDE: Next_Recommended_Skill is NONE or an existing Skill_ID --------
            $nrs = "$($state['Next_Recommended_Skill'])"
            $nrsOk = ($nrs -eq '' -or $nrs -eq 'NONE' -or $skillIds -contains $nrs)
            Add-Result $fname 'GUIDE' 'Invalid_Next_Recommended_Skill' $nrsOk "Next_Recommended_Skill='$nrs'"
        }
        finally { $pkg.Dispose() }
    }
    catch {
        Add-Result $fname 'FILE' 'Readable_And_Schema' $false $_.Exception.Message
    }
}

# ---- report -------------------------------------------------------------------
if ($AsJson) { $results | ConvertTo-Json -Depth 4 }
else {
    $results | Format-Table File, Test_ID, Check, Result, Detail -AutoSize
    $fail = @($results | Where-Object Result -eq 'FAIL')
    $pass = @($results | Where-Object Result -eq 'PASS')
    Write-Host ""
    Write-Host ("harness-lint {0}: {1} PASS, {2} FAIL across {3} file(s)" -f $ExpectedVersion, $pass.Count, $fail.Count, $Path.Count)
}
if (@($results | Where-Object Result -eq 'FAIL').Count -gt 0) { exit 1 } else { exit 0 }
