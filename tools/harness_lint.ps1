[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ExpectedVersion = 'v6.15.1'
$ExpectedSkillCount = 43
$ExpectedRoutingCount = 55
$RequiredSheets = @(
    '00_Skills', 'Project.Rules', '__STATE', '__TEST_ORACLE',
    'Lists', '__ADR', '__GLOSSARY', '__ROUTING_ORACLE'
)
$ValidModes = @('HARNESS', 'TEMPLATE', 'PROJECT-CREATION', 'PROJECT', 'MIGRATION')
$ExpectedStateFields = @(
    'Harness version', 'File_Role', 'Template_Version', 'Project_ID', 'Project_Safe_Name',
    'Revision', 'Logical_Revision', 'Revision_Type', 'Parent_Artifact', 'Current_Artifact',
    'Active_Artifact', 'Physical_Artifact', 'Physical_Artifact_Changed', 'Artifact_Status',
    'Created_From', 'Last_Export_Verified', 'Output_Verification_Status', 'Host_Mode',
    'Persistence_Mode', 'Storage_Version_ID', 'Native_Version_History_Status',
    'Precreated_Target_Status', 'Last_Completed_Action', 'Last_Attempted_Action',
    'Last_Failed_Action', 'Last_Verified_Action', 'Failed_Artifact', '__WRITE_PROBE',
    'Failure_Stage', 'Failure_Code', 'Current_Skill', 'Current_Phase', 'Next step',
    'Resume_From', 'State_Status', 'Capability_Mode', 'Source_Refs', 'Phase_Gate_Status',
    'Selected_Shape', 'Current_Plan_ID', 'Last_Verified_Gate', 'Next_Recommended_Skill',
    'Next_Executable_Step'
)

function Get-Headers {
    param($Values, [string]$FirstHeader)
    $HeaderRow = 0
    for ($Row = 1; $Row -le $Values.GetLength(0); $Row++) {
        if ([string]$Values[$Row, 1] -eq $FirstHeader) {
            $HeaderRow = $Row
            break
        }
    }
    if (-not $HeaderRow) { throw "Header not found: $FirstHeader" }

    $Headers = @{}
    for ($Column = 1; $Column -le $Values.GetLength(1); $Column++) {
        $Name = [string]$Values[$HeaderRow, $Column]
        if ($Name) {
            if ($Headers.ContainsKey($Name)) { throw "Duplicate header: $Name" }
            $Headers[$Name] = $Column
        }
    }
    [pscustomobject]@{ Row = $HeaderRow; Map = $Headers }
}

function Assert-RequiredHeaders {
    param([hashtable]$Headers, [string[]]$Required)
    foreach ($Name in $Required) {
        if (-not $Headers.ContainsKey($Name)) { throw "Missing column: $Name" }
    }
}

function Get-ColumnValues {
    param($Values, [int]$Column)
    $Result = @()
    for ($Row = 1; $Row -le $Values.GetLength(0); $Row++) {
        $Value = [string]$Values[$Row, $Column]
        if ($Value) { $Result += $Value }
    }
    $Result
}

function Assert-PrintableAscii {
    param($Workbook, [string[]]$SheetNames)
    foreach ($SheetName in $SheetNames) {
        $Values = $Workbook.Worksheets.Item($SheetName).UsedRange.Value2
        for ($Row = 1; $Row -le $Values.GetLength(0); $Row++) {
            for ($Column = 1; $Column -le $Values.GetLength(1); $Column++) {
                $Value = [string]$Values[$Row, $Column]
                foreach ($Character in $Value.ToCharArray()) {
                    $Code = [int][char]$Character
                    if ($Code -lt 32 -or $Code -gt 126) {
                        throw "Non-ASCII character in ${SheetName}!R${Row}C${Column}: U+$($Code.ToString('X4'))"
                    }
                }
            }
        }
    }
}

function Assert-NoStaleProbeAddress {
    param($Workbook)
    $ActiveSheets = @('00_Skills', 'Project.Rules', '00_Landing', '__TEST_ORACLE', '__GLOSSARY')
    foreach ($SheetName in $ActiveSheets) {
        if (@($Workbook.Worksheets | ForEach-Object { $_.Name }) -notcontains $SheetName) { continue }
        $Values = $Workbook.Worksheets.Item($SheetName).UsedRange.Value2
        for ($Row = 1; $Row -le $Values.GetLength(0); $Row++) {
            for ($Column = 1; $Column -le $Values.GetLength(1); $Column++) {
                $Value = [string]$Values[$Row, $Column]
                if ($Value -match '(?i)B32\s+only\s+for\s+explicit\s+probe') {
                    throw "Stale active B32 probe instruction in ${SheetName}!R${Row}C${Column}"
                }
            }
        }
    }
}

function Test-Workbook {
    param($Excel, [string]$Item)

    $FullPath = (Resolve-Path -LiteralPath $Item).Path
    $Workbook = $Excel.Workbooks.Open($FullPath, 0, $true)
    try {
        $SheetNames = @($Workbook.Worksheets | ForEach-Object { $_.Name })
        foreach ($Name in $RequiredSheets) {
            if ($SheetNames -notcontains $Name) { throw "Missing sheet: $Name" }
        }

        $Skills = $Workbook.Worksheets.Item('00_Skills').UsedRange.Value2
        $SkillHeader = Get-Headers -Values $Skills -FirstHeader 'Skill_ID'
        $SkillHeaders = $SkillHeader.Map
        Assert-RequiredHeaders -Headers $SkillHeaders -Required @(
            'Skill_ID', 'Trigger', 'Version', 'May_Chain_To', 'Allowed_Modes'
        )

        $SkillIds = @()
        $TriggerAtoms = @()
        $SkillVersions = @()
        $AllowedModesErrors = @()
        $ChainTargets = @()

        for ($Row = $SkillHeader.Row + 1; $Row -le $Skills.GetLength(0); $Row++) {
            $SkillId = [string]$Skills[$Row, $SkillHeaders.Skill_ID]
            if (-not $SkillId) { continue }

            $SkillIds += $SkillId
            $SkillVersions += [string]$Skills[$Row, $SkillHeaders.Version]
            $TriggerAtoms += @(
                ([string]$Skills[$Row, $SkillHeaders.Trigger] -split ' / ') |
                    ForEach-Object { $_.Trim().ToLowerInvariant() } |
                    Where-Object { $_ }
            )
            $Chain = [string]$Skills[$Row, $SkillHeaders.May_Chain_To]
            if ($Chain -and $Chain -ne 'NONE') {
                $ChainTargets += @($Chain -split ', ' | Where-Object { $_ })
            }

            $ModeValue = [string]$Skills[$Row, $SkillHeaders.Allowed_Modes]
            if (-not $ModeValue) {
                $AllowedModesErrors += "${SkillId}:empty"
            }
            elseif ($ModeValue -ne 'ALL') {
                $Parts = @($ModeValue -split ', ')
                if (($Parts -join ', ') -ne $ModeValue) {
                    $AllowedModesErrors += "${SkillId}:non-canonical delimiter spacing"
                }
                if ($Parts -contains 'ALL') {
                    $AllowedModesErrors += "${SkillId}:ALL cannot be combined"
                }
                if (@($Parts | Group-Object | Where-Object Count -gt 1).Count) {
                    $AllowedModesErrors += "${SkillId}:duplicate mode"
                }
                foreach ($Mode in $Parts) {
                    if ($ValidModes -notcontains $Mode) {
                        $AllowedModesErrors += "${SkillId}:$Mode"
                    }
                }
            }
        }

        if ($SkillIds.Count -ne $ExpectedSkillCount) {
            throw "Skill count $($SkillIds.Count), expected $ExpectedSkillCount"
        }
        if (@($SkillIds | Group-Object | Where-Object Count -gt 1).Count) {
            throw 'Duplicate Skill_ID'
        }
        if (@($SkillVersions | Where-Object { $_ -ne $ExpectedVersion }).Count) {
            throw 'Skill version mismatch'
        }
        if (@($TriggerAtoms | Group-Object | Where-Object Count -gt 1).Count) {
            throw 'Duplicate trigger atom'
        }
        foreach ($Left in $TriggerAtoms) {
            foreach ($Right in $TriggerAtoms) {
                if ($Left -ne $Right -and $Right.StartsWith($Left)) {
                    throw "Trigger prefix collision: $Left -> $Right"
                }
            }
        }
        $UnresolvedChains = @($ChainTargets | Where-Object { $SkillIds -notcontains $_ } | Sort-Object -Unique)
        if ($UnresolvedChains.Count) {
            throw "Unresolved chain references: $($UnresolvedChains -join ', ')"
        }
        if ($AllowedModesErrors.Count) {
            throw "Invalid Allowed_Modes: $($AllowedModesErrors -join ', ')"
        }

        $State = $Workbook.Worksheets.Item('__STATE').UsedRange.Value2
        $StateHeader = Get-Headers -Values $State -FirstHeader 'Field'
        $StateHeaders = $StateHeader.Map
        Assert-RequiredHeaders -Headers $StateHeaders -Required @('Field', 'Value')
        $StateMap = @{}
        for ($Row = $StateHeader.Row + 1; $Row -le $State.GetLength(0); $Row++) {
            $Field = [string]$State[$Row, $StateHeaders.Field]
            if (-not $Field) { continue }
            if ($StateMap.ContainsKey($Field)) { throw "Duplicate state field: $Field" }
            $StateMap[$Field] = [string]$State[$Row, $StateHeaders.Value]
        }
        foreach ($Field in $ExpectedStateFields) {
            if (-not $StateMap.ContainsKey($Field)) { throw "Missing state field: $Field" }
        }
        if ($StateMap['Harness version'] -ne $ExpectedVersion) { throw 'State Harness version mismatch' }
        if ($StateMap['Template_Version'] -ne $ExpectedVersion) { throw 'State Template_Version mismatch' }
        if ($StateMap['Next_Recommended_Skill'] -ne 'NONE' -and $SkillIds -notcontains $StateMap['Next_Recommended_Skill']) {
            throw "Invalid Next_Recommended_Skill: $($StateMap['Next_Recommended_Skill'])"
        }

        $PhysicalName = [IO.Path]::GetFileName($FullPath)
        if ($StateMap['Current_Artifact'] -ne $PhysicalName) {
            throw "Current_Artifact mismatch: state=$($StateMap['Current_Artifact']); file=$PhysicalName"
        }

        $Probe = $Workbook.Names.Item('PROBE_CELL').RefersTo
        $ProbeRow = 0
        for ($Row = $StateHeader.Row + 1; $Row -le $State.GetLength(0); $Row++) {
            if ([string]$State[$Row, $StateHeaders.Field] -eq '__WRITE_PROBE') { $ProbeRow = $Row }
        }
        if ($Probe -notmatch '__STATE!\$B\$31') { throw "PROBE_CELL mismatch: $Probe" }
        if ($ProbeRow -ne 31) { throw "__WRITE_PROBE row $ProbeRow, expected 31" }

        $Routing = $Workbook.Worksheets.Item('__ROUTING_ORACLE').UsedRange.Value2
        $RoutingHeader = Get-Headers -Values $Routing -FirstHeader 'Test_ID'
        $RoutingHeaders = $RoutingHeader.Map
        Assert-RequiredHeaders -Headers $RoutingHeaders -Required @(
            'Test_ID', 'Layer', 'Category', 'Expected_Skill_ID', 'Critical',
            'Fixture_Context', 'Candidate_Skill_Set'
        )
        $RoutingIds = @()
        $RoutingRefs = @()
        $RoutingCount = 0
        for ($Row = $RoutingHeader.Row + 1; $Row -le $Routing.GetLength(0); $Row++) {
            $TestId = [string]$Routing[$Row, $RoutingHeaders.Test_ID]
            if (-not $TestId) { continue }
            $RoutingCount++
            $RoutingIds += $TestId
            $RoutingRefs += [string]$Routing[$Row, $RoutingHeaders.Expected_Skill_ID]
            if ([string]$Routing[$Row, $RoutingHeaders.Category] -eq 'TIE') {
                if (-not [string]$Routing[$Row, $RoutingHeaders.Fixture_Context] -or
                    -not [string]$Routing[$Row, $RoutingHeaders.Candidate_Skill_Set]) {
                    throw "Incomplete tie fixture: $TestId"
                }
            }
        }
        if ($RoutingCount -ne $ExpectedRoutingCount) {
            throw "Routing rows $RoutingCount, expected $ExpectedRoutingCount"
        }
        if (@($RoutingIds | Group-Object | Where-Object Count -gt 1).Count) {
            throw 'Duplicate routing Test_ID'
        }
        $BadRoutingRefs = @(
            $RoutingRefs | Where-Object { $_ -ne 'NONE' -and $SkillIds -notcontains $_ } | Sort-Object -Unique
        )
        if ($BadRoutingRefs.Count) {
            throw "Unresolved routing skill refs: $($BadRoutingRefs -join ', ')"
        }

        $Oracle = $Workbook.Worksheets.Item('__TEST_ORACLE').UsedRange.Value2
        $OracleHeader = Get-Headers -Values $Oracle -FirstHeader 'Test_ID'
        $OracleHeaders = $OracleHeader.Map
        Assert-RequiredHeaders -Headers $OracleHeaders -Required @(
            'Test_ID', 'Oracle_Version', 'Required_Field', 'Expected_Value',
            'Evidence_Class_Required', 'Critical', 'Mismatch_Result'
        )
        $OracleKeys = @()
        for ($Row = $OracleHeader.Row + 1; $Row -le $Oracle.GetLength(0); $Row++) {
            $TestId = [string]$Oracle[$Row, $OracleHeaders.Test_ID]
            if (-not $TestId) { continue }
            $Key = '{0}|{1}|{2}' -f $TestId,
                [string]$Oracle[$Row, $OracleHeaders.Oracle_Version],
                [string]$Oracle[$Row, $OracleHeaders.Required_Field]
            $OracleKeys += $Key
        }
        if (@($OracleKeys | Group-Object | Where-Object Count -gt 1).Count) {
            throw 'Duplicate __TEST_ORACLE identity key'
        }

        Assert-PrintableAscii -Workbook $Workbook -SheetNames @(
            '00_Skills', 'Project.Rules', '__STATE', '00_Landing', '__TEST_ORACLE',
            'Lists', '__ADR', '__GLOSSARY', '__ROUTING_ORACLE'
        )
        Assert-NoStaleProbeAddress -Workbook $Workbook

        Write-Host "PASS $FullPath skills=43 routing=55 probe=B31 modes=ok" -ForegroundColor Green
        return [pscustomobject]@{ Path = $FullPath; Result = 'PASS'; Error = $null }
    }
    catch {
        return [pscustomobject]@{ Path = $FullPath; Result = 'FAIL'; Error = $_.Exception.Message }
    }
    finally {
        $Workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook)
    }
}

$Excel = $null
$Results = @()
try {
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false

    foreach ($Item in $Path) {
        $Results += Test-Workbook -Excel $Excel -Item $Item
    }
}
finally {
    if ($Excel) {
        $Excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$Results | Format-Table -AutoSize
if (@($Results | Where-Object Result -eq 'FAIL').Count) { exit 1 }
exit 0
