[CmdletBinding()]param([Parameter(Mandatory=$true)][string[]]$Path)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$expectedVersion='v6.13.0'
$excel=$null
$failed=$false
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false
 foreach($item in $Path){
  $full=(Resolve-Path -LiteralPath $item).Path
  $wb=$excel.Workbooks.Open($full,0,$true)
  try {
   $names=@($wb.Worksheets|ForEach-Object{$_.Name})
   $required=@('00_Skills','Project.Rules','__STATE','__TEST_ORACLE','Lists','__ADR','__GLOSSARY','__ROUTING_ORACLE')
   foreach($n in $required){if($names -notcontains $n){throw "Missing sheet: $n"}}
   $skills=$wb.Worksheets.Item('00_Skills').UsedRange.Value2
   $hr=0;for($r=1;$r -le $skills.GetLength(0);$r++){if($skills[$r,1]-eq 'Skill_ID'){$hr=$r;break}}
   if(!$hr){throw 'Skill header not found'}
   $headers=@{};for($c=1;$c -le $skills.GetLength(1);$c++){$headers[[string]$skills[$hr,$c]]=$c}
   $ids=@();$atoms=@();$versions=@()
   for($r=$hr+1;$r -le $skills.GetLength(0);$r++){if($skills[$r,$headers.Skill_ID]){$ids+=[string]$skills[$r,$headers.Skill_ID];$versions+=[string]$skills[$r,$headers.Version];$atoms+=([string]$skills[$r,$headers.Trigger]-split ' / '|ForEach-Object{$_.Trim().ToLowerInvariant()}|Where-Object{$_})}}
   if($ids.Count-ne 41){throw "Skill count $($ids.Count), expected 41"}
   if(@($ids|Group-Object|Where-Object Count -gt 1).Count){throw 'Duplicate Skill_ID'}
   if(@($versions|Where-Object{$_-ne $expectedVersion}).Count){throw 'Skill version mismatch'}
   if(@($atoms|Group-Object|Where-Object Count -gt 1).Count){throw 'Duplicate trigger atom'}
   foreach($a in $atoms){foreach($b in $atoms){if($a-ne$b -and $b.StartsWith($a)){throw "Trigger prefix collision: $a -> $b"}}}
   $ro=$wb.Worksheets.Item('__ROUTING_ORACLE').UsedRange.Value2
   $rr=0;for($r=1;$r-le$ro.GetLength(0);$r++){if($ro[$r,1]-eq 'Test_ID'){$rr=$r;break}}
   $rh=@{};for($c=1;$c-le$ro.GetLength(1);$c++){$rh[[string]$ro[$rr,$c]]=$c}
   foreach($key in @('Fixture_Context','Candidate_Skill_Set')){if(!$rh.ContainsKey($key)){throw "Missing routing column: $key"}}
   $tids=@();$refs=@();$routingCount=0
   for($r=$rr+1;$r-le$ro.GetLength(0);$r++){if($ro[$r,$rh.Test_ID]){$routingCount++;$tids+=[string]$ro[$r,$rh.Test_ID];$refs+=[string]$ro[$r,$rh.Expected_Skill_ID];if($ro[$r,$rh.Category]-eq'TIE' -and (!$ro[$r,$rh.Fixture_Context] -or !$ro[$r,$rh.Candidate_Skill_Set])){throw "Incomplete tie fixture: $($ro[$r,$rh.Test_ID])"}}}
   if($routingCount-ne 52){throw "Routing rows $routingCount, expected 52"}
   if(@($tids|Group-Object|Where-Object Count -gt 1).Count){throw 'Duplicate routing Test_ID'}
   $bad=@($refs|Where-Object{$_-ne'NONE' -and $ids-notcontains$_});if($bad.Count){throw "Unresolved routing skill refs: $($bad-join', ')"}
   $probe=$wb.Names.Item('PROBE_CELL').RefersTo
   if($probe-notmatch "__STATE!\`?\$B\$31") {throw "PROBE_CELL mismatch: $probe"}
   $state=$wb.Worksheets.Item('__STATE').UsedRange.Value2;$probeRow=0
   for($r=1;$r-le$state.GetLength(0);$r++){if($state[$r,1]-eq'__WRITE_PROBE'){$probeRow=$r}}
   if($probeRow-ne31){throw "__WRITE_PROBE row $probeRow, expected 31"}
   Write-Host "PASS $full skills=41 routing=52 probe=B31" -ForegroundColor Green
  } finally {$wb.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)}
 }
} catch {$failed=$true;Write-Error $_} finally {if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
if($failed){exit 1}else{exit 0}
