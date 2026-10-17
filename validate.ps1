<#
.SYNOPSIS
  validate.ps1 - Sysmon validation harness for the Synthetic Human Agent (team05, P3).

.DESCRIPTION
  Reads Sysmon events (ID 1 process create, 5 process terminate, 11 file create,
  22 DNS query) for one agent run and checks it against the guardrails in
  MASTER_BUILD_PLAN.md:

    1. Every app launch has ParentImage = explorer.exe (Chrome, Writer, Notepad)
    2. No process is parented by python / cmd / powershell / wscript ...
    3. Cold-launch discipline: each app fully closed before it is launched again
    4. Every launched app was closed again (and nothing is left running)
    5. No project-folder / .py trace in the app processes' command line or cwd
    6. No Task Scheduler / Start-Process / startfile traces
    7. Run length <= 15 minutes
    8. Timing is not metronomic (gap statistics + histogram of action gaps)
    9. Office (Writer or Calc) saved a document (Sysmon ID 11, or file on disk as fallback)
   10. The agent's own interpreter (pythonw.exe running agent.pyw) was started by explorer.exe
   11. The agent interpreter created no child processes
   12. Agent lifetime is about 15 minutes (15:00 +/- 30 s)
   13. No cmd / powershell / conhost / wsl / script-interpreter process was created during the
       run (WARN, because judges look for exactly these)
   Timing also reports p95, repeated (identical-to-the-second) gaps and the longest gap, and
   the verdict needs only 4 gaps. Targets are the team's own heuristics, not judge thresholds.

  MUST RUN ELEVATED (Get-WinEvent on the Sysmon log needs admin).

.PARAMETER Since / Until   Explicit local time window.
.PARAMETER Minutes         Look back N minutes from Until.
                           With none of the three: window starts 5 s before the most
                           recent python.exe whose command line mentions agent/sequencer/
                           office/web/engine .py (last 3 h), else the last 20 minutes.
.PARAMETER MaxRunMinutes   Run-length cap (default 15).
.PARAMETER OutDir          Where the CSV evidence is written (default .\validation).

.EXAMPLE
  .\validate.ps1
  .\validate.ps1 -Minutes 30
  .\validate.ps1 -Since "2026-10-10 21:00" -Until "2026-10-10 21:20"

  Exit code: 0 = no FAIL, 1 = at least one FAIL, 2 = could not run.
#>
[CmdletBinding()]
param(
    [datetime]$Since,
    [datetime]$Until,
    [int]$Minutes = 0,
    [int]$MaxRunMinutes = 15,
    [string]$OutDir
)

$ErrorActionPreference = 'Stop'
$ScriptStart = Get-Date
if (-not $Until) { $Until = $ScriptStart }
if (-not $OutDir) {
    $base = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $OutDir = Join-Path $base 'validation'
}

# ---------------------------------------------------------------- preflight
$ident = [Security.Principal.WindowsIdentity]::GetCurrent()
$princ = New-Object Security.Principal.WindowsPrincipal($ident)
if (-not $princ.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "validate.ps1 must run in an ELEVATED PowerShell (Run as administrator)." -ForegroundColor Red
    exit 2
}
$LogName = 'Microsoft-Windows-Sysmon/Operational'
try { $null = Get-WinEvent -ListLog $LogName } catch {
    Write-Host "Sysmon log '$LogName' not found. Is Sysmon installed and running?" -ForegroundColor Red
    exit 2
}

# ---------------------------------------------------------------- helpers
function Get-Sysmon {
    param([int[]]$Ids, [datetime]$Start, [datetime]$End)
    $evs = Get-WinEvent -FilterHashtable @{ LogName = $LogName; Id = $Ids; StartTime = $Start; EndTime = $End } -ErrorAction SilentlyContinue
    foreach ($e in $evs) {
        $xml = [xml]$e.ToXml()
        $h = [ordered]@{ EventId = $e.Id; TimeCreated = $e.TimeCreated }
        foreach ($d in $xml.Event.EventData.Data) { $h[$d.Name] = [string]$d.'#text' }
        $h['ImgName'] = if ($h['Image'])       { [IO.Path]::GetFileName($h['Image']).ToLower() }       else { '' }
        $h['ParName'] = if ($h['ParentImage']) { [IO.Path]::GetFileName($h['ParentImage']).ToLower() } else { '' }
        [pscustomobject]$h
    }
}

$Results = New-Object System.Collections.ArrayList
function Add-Check {
    param([string]$Name, [string]$Status, [string]$Detail)
    [void]$Results.Add([pscustomobject]@{ Check = $Name; Status = $Status; Detail = $Detail })
}

function Get-Stats {
    param([double[]]$Values)
    $v = @($Values)
    $n = $v.Count
    if ($n -eq 0) { return $null }
    $s = @($v | Sort-Object)
    $mean = ($v | Measure-Object -Average).Average
    $ss = 0.0
    foreach ($x in $v) { $ss += ($x - $mean) * ($x - $mean) }
    $sd = if ($n -gt 1) { [math]::Sqrt($ss / ($n - 1)) } else { 0.0 }
    $med = if ($n % 2 -eq 1) { $s[[int](($n - 1) / 2)] } else { ($s[[int]($n / 2 - 1)] + $s[[int]($n / 2)]) / 2 }
    [pscustomobject]@{
        N = $n; Min = $s[0]; Median = $med; Mean = $mean; SD = $sd
        CV = $(if ($mean -gt 0) { $sd / $mean } else { 0 }); Max = $s[$n - 1]
        P95 = $s[[math]::Min($n - 1, [int][math]::Ceiling(0.95 * $n) - 1)]
    }
}

function Show-Histogram {
    param([double[]]$Gaps)
    $bins = @(@(0, 5), @(5, 10), @(10, 20), @(20, 40), @(40, 80), @(80, 1e9))
    foreach ($b in $bins) {
        $c = @($Gaps | Where-Object { $_ -ge $b[0] -and $_ -lt $b[1] }).Count
        $label = if ($b[1] -ge 1e9) { "{0}+ s" -f $b[0] } else { "{0}-{1} s" -f $b[0], $b[1] }
        Write-Host ("  {0,-10} {1,3}  {2}" -f $label, $c, ('#' * $c))
    }
}

# ---------------------------------------------------------------- time window
$anchorNote = ''
if ($Since) {
    $anchorNote = 'explicit -Since'
} elseif ($Minutes -gt 0) {
    $Since = $Until.AddMinutes(-$Minutes)
    $anchorNote = "last $Minutes min"
} else {
    $probe = @(Get-Sysmon -Ids 1 -Start $Until.AddMinutes(-180) -End $Until |
        Where-Object { $_.ImgName -match '^pythonw?\.exe$' -and $_.CommandLine -match '(agent|sequencer|office|web|engine)\.py' } |
        Sort-Object TimeCreated) | Select-Object -Last 1
    if ($probe) {
        $Since = $probe.TimeCreated.AddSeconds(-5)
        $anchorNote = 'anchored on latest agent python.exe start'
    } else {
        $Since = $Until.AddMinutes(-20)
        $anchorNote = 'no agent python.exe found; last 20 min'
    }
}

Write-Host ""
Write-Host "=== Sysmon validation ===" -ForegroundColor Cyan
Write-Host ("Window : {0} -> {1}  ({2})" -f $Since.ToString('yyyy-MM-dd HH:mm:ss'), $Until.ToString('yyyy-MM-dd HH:mm:ss'), $anchorNote)
Write-Host ("Clock  : VM now = {0}  (if this is wrong the window is wrong - see clock note in plan)" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))

# ---------------------------------------------------------------- fetch
$procs = @(Get-Sysmon -Ids 1  -Start $Since -End $Until | Where-Object { $_.ProcessId -ne "$PID" } | Sort-Object TimeCreated)
$terms = @(Get-Sysmon -Ids 5  -Start $Since -End $Until)
$files = @(Get-Sysmon -Ids 11 -Start $Since -End $Until)
$dns   = @(Get-Sysmon -Ids 22 -Start $Since -End $Until)

Write-Host ("Events : {0} process-create, {1} terminate, {2} file-create, {3} dns" -f $procs.Count, $terms.Count, $files.Count, $dns.Count)
if ($procs.Count -eq 0) {
    Write-Host "No Sysmon process-create events in the window. Wrong window, or Sysmon not logging?" -ForegroundColor Red
    exit 2
}

$termByGuid = @{}
foreach ($t in $terms) { $termByGuid[$t.ProcessGuid] = $t.TimeCreated }

# ---------------------------------------------------------------- app families
$notChromeChild = { param($e) $e.CommandLine -notmatch '--type=' }
$always         = { param($e) $true }
$notSofficeKid  = { param($e) $e.ParName -notin @('soffice.exe', 'soffice.bin') }

$families = @(
    @{ Name = 'Chrome';  Launch = 'chrome.exe';  Life = 'chrome.exe';  LaunchMain = $notChromeChild; LifeMain = $notChromeChild },
    @{ Name = 'Writer/Calc'; Launch = 'soffice.exe'; Life = 'soffice.bin'; LaunchMain = $notSofficeKid;  LifeMain = $always },
    @{ Name = 'Notepad'; Launch = 'notepad.exe'; Life = 'notepad.exe'; LaunchMain = $always;         LifeMain = $always }
)

$allMain = New-Object System.Collections.ArrayList
$famLaunches = @{}
foreach ($f in $families) {
    $L = @($procs | Where-Object { $_.ImgName -eq $f.Launch -and (& $f.LaunchMain $_) } | Sort-Object TimeCreated)
    $famLaunches[$f.Name] = $L
    foreach ($e in $L) { [void]$allMain.Add($e) }
}

Write-Host ""
Write-Host "--- App launches ---" -ForegroundColor Cyan
if ($allMain.Count -eq 0) {
    Write-Host "  (none found)"
} else {
    $allMain | Sort-Object TimeCreated | ForEach-Object {
        $cmd = if ($_.CommandLine.Length -gt 70) { $_.CommandLine.Substring(0, 70) + '...' } else { $_.CommandLine }
        Write-Host ("  {0}  {1,-12} parent={2,-14} cwd={3}  {4}" -f $_.TimeCreated.ToString('HH:mm:ss'), $_.ImgName, $_.ParName, $_.CurrentDirectory, $cmd)
    }
}

# ---------------------------------------------------------------- CHECK 1: explorer parent
$script:assocStub = $null
foreach ($f in $families) {
    $L = $famLaunches[$f.Name]
    if ($L.Count -eq 0) { continue }
    $bad = @($L | Where-Object { $_.ParName -ne 'explorer.exe' })
    # A document opened by double-click / Enter in File Explorer goes through the OpenOffice
    # launcher stub: explorer.exe -> swriter.exe (scalc.exe ...) -> soffice.exe.  Tell that
    # apart from a real bad parent.
    $viaAssoc = @()
    $hard = @()
    foreach ($b in $bad) {
        $isStub = $false
        if ($b.ParName -match '^s(writer|calc|impress|base|draw|math)\.exe$') {
            $stub = @($procs | Where-Object { $_.ProcessId -eq $b.ParentProcessId -and $_.ImgName -eq $b.ParName }) | Select-Object -First 1
            if ($stub -and $stub.ParName -eq 'explorer.exe') { $isStub = $true }
        }
        if ($isStub) { $viaAssoc += $b } else { $hard += $b }
    }
    if ($viaAssoc.Count -gt 0) { $script:assocStub = $viaAssoc[0].ParName }
    if ($bad.Count -eq 0) {
        Add-Check "Explorer parent: $($f.Name)" 'PASS' ("{0}/{0} launches have ParentImage = explorer.exe" -f $L.Count)
    } elseif ($hard.Count -eq 0) {
        Add-Check "Explorer parent: $($f.Name)" 'WARN' ("{0}/{1} launches went explorer.exe -> {2} -> soffice.exe (opened from File Explorer via the file association: explorer.exe is the grandparent, not the direct parent)" -f $viaAssoc.Count, $L.Count, $script:assocStub)
    } else {
        $who = ($hard | ForEach-Object { $_.ParName } | Sort-Object -Unique) -join ', '
        Add-Check "Explorer parent: $($f.Name)" 'FAIL' ("{0}/{1} launches NOT from explorer.exe (parents: {2})" -f $hard.Count, $L.Count, $who)
    }
}
if ($allMain.Count -eq 0) {
    Add-Check 'App launches found' 'FAIL' 'No Chrome / Writer / Notepad launch in the window'
}

# soffice chain: soffice.bin should hang off soffice.exe
$bins = @($procs | Where-Object { $_.ImgName -eq 'soffice.bin' })
if ($bins.Count -gt 0) {
    $badBin = @($bins | Where-Object { $_.ParName -notin @('soffice.exe', 'soffice.bin') })
    if ($badBin.Count -eq 0) {
        $chainTxt = if ($script:assocStub) { "explorer.exe -> $($script:assocStub) -> soffice.exe -> soffice.bin" } else { 'explorer.exe -> soffice.exe -> soffice.bin' }
        Add-Check 'Writer chain' 'PASS' $chainTxt
    } else {
        Add-Check 'Writer chain' 'WARN' ("soffice.bin has unexpected parent(s): {0}" -f (($badBin | ForEach-Object { $_.ParName } | Sort-Object -Unique) -join ', '))
    }
}

# ---------------------------------------------------------------- CHECK 2: bad parents
$badParents = @('python.exe', 'pythonw.exe', 'py.exe', 'cmd.exe', 'powershell.exe', 'pwsh.exe', 'wscript.exe', 'cscript.exe')
# Scope: only the AGENT's apps count. Anything else in the window that has a script/shell
# parent (monitor.ps1, wevtutil, validate.ps1 itself, the python launcher, conhost...) is
# tooling around the run, not the agent, and must not fail the check.
$agentApps = @('chrome.exe', 'soffice.exe', 'soffice.bin', 'winword.exe', 'excel.exe', 'powerpnt.exe',
               'notepad.exe', 'calc.exe', 'calculator.exe', 'calculatorapp.exe')
$flagged = New-Object System.Collections.ArrayList
$info = 0
foreach ($e in $procs) {
    if ($e.ParName -notin $badParents) { continue }
    if ($e.ImgName -notin $agentApps) { $info++; continue }   # not an agent app: ignore (tooling / start-up noise)
    [void]$flagged.Add($e)
}
if ($flagged.Count -eq 0) {
    Add-Check 'No script/shell parents' 'PASS' ("No agent app (chrome/soffice/winword/excel/notepad/calc) has a python/cmd/powershell parent ({0} non-agent tooling process(es) ignored)" -f $info)
} else {
    $top = ($flagged | Select-Object -First 4 | ForEach-Object { "{0} <- {1}" -f $_.ImgName, $_.ParName }) -join '; '
    Add-Check 'No script/shell parents' 'FAIL' ("{0} process(es) with a script/shell parent, e.g. {1}" -f $flagged.Count, $top)
    Write-Host ""
    Write-Host "--- Agent apps with a script/shell parent ---" -ForegroundColor Yellow
    $flagged | ForEach-Object { Write-Host ("  {0}  {1} <- {2}  {3}" -f $_.TimeCreated.ToString('HH:mm:ss'), $_.ImgName, $_.ParName, $_.CommandLine) }
}

# ---------------------------------------------------------------- CHECK 3: cold-launch discipline
$coldViol = New-Object System.Collections.ArrayList
foreach ($f in $families) {
    $L = @($famLaunches[$f.Name])
    for ($i = 0; $i -lt $L.Count - 1; $i++) {
        $a = $L[$i]; $b = $L[$i + 1]
        $life = @($procs | Where-Object {
            $_.ImgName -eq $f.Life -and $_.TimeCreated -ge $a.TimeCreated.AddSeconds(-1) -and
            $_.TimeCreated -lt $b.TimeCreated -and (& $f.LifeMain $_) })
        foreach ($p in $life) {
            $tt = $termByGuid[$p.ProcessGuid]
            if (-not $tt -or $tt -gt $b.TimeCreated) {
                [void]$coldViol.Add(("{0}: relaunch at {1} while earlier instance (started {2}) was still running" -f $f.Name, $b.TimeCreated.ToString('HH:mm:ss'), $p.TimeCreated.ToString('HH:mm:ss')))
            }
        }
    }
}
if ($allMain.Count -gt 0) {
    if ($coldViol.Count -eq 0) {
        Add-Check 'Cold-launch discipline' 'PASS' 'Every app was closed before it was launched again'
    } else {
        Add-Check 'Cold-launch discipline' 'FAIL' (($coldViol | Select-Object -First 3) -join ' | ')
    }
}

# ---------------------------------------------------------------- CHECK 4: closed again
$notClosed = New-Object System.Collections.ArrayList
foreach ($f in $families) {
    $life = @($procs | Where-Object { $_.ImgName -eq $f.Life -and (& $f.LifeMain $_) })
    foreach ($p in $life) { if (-not $termByGuid.ContainsKey($p.ProcessGuid)) { [void]$notClosed.Add($f.Name) } }
}
$left = @(Get-Process -Name 'chrome', 'soffice', 'soffice.bin', 'notepad' -ErrorAction SilentlyContinue)
if ($allMain.Count -gt 0) {
    if ($notClosed.Count -eq 0 -and $left.Count -eq 0) {
        Add-Check 'Apps closed at end' 'PASS' 'All launched apps have a terminate event; nothing left running'
    } else {
        $names = ((@($notClosed.ToArray()) + @($left | ForEach-Object { $_.ProcessName })) | Sort-Object -Unique) -join ', '
        Add-Check 'Apps closed at end' 'WARN' ("Still running / no terminate event: {0} (run may still be in progress, or an app was never closed)" -f $names)
    }
}

# ---------------------------------------------------------------- CHECK 5: project-folder trace
$traceImgsSkip = @('python.exe', 'pythonw.exe', 'py.exe', 'cmd.exe', 'powershell.exe', 'pwsh.exe', 'conhost.exe')
$trace = @($procs | Where-Object {
    $_.ImgName -notin $traceImgsSkip -and
    ($_.CommandLine -match 'synth\w*-agent|\.py\b|\.pyw\b' -or $_.CurrentDirectory -match 'synth\w*-agent') })   # matches synthetic-agent AND the real folder name synthentic-agent
if ($trace.Count -eq 0) {
    Add-Check 'No project-folder trace' 'PASS' 'No app process mentions the project folder (synthetic/synthentic-agent) or .py/.pyw in cmdline / cwd'
} else {
    Add-Check 'No project-folder trace' 'FAIL' ("{0} process(es), e.g. {1} cwd={2}" -f $trace.Count, $trace[0].ImgName, $trace[0].CurrentDirectory)
}

# ---------------------------------------------------------------- CHECK 6: forbidden mechanisms
$forbidden = @($procs | Where-Object {
    $_.ImgName -in @('schtasks.exe', 'at.exe', 'taskeng.exe') -or
    ($_.ImgName -in @('powershell.exe', 'pwsh.exe', 'cmd.exe') -and $_.CommandLine -match 'Start-Process|startfile|\bstart\s+"?\w') })
if ($forbidden.Count -eq 0) {
    Add-Check 'No Task Scheduler / Start-Process' 'PASS' 'No schtasks / Start-Process / start / startfile traces'
} else {
    Add-Check 'No Task Scheduler / Start-Process' 'FAIL' ("{0} hit(s), e.g. {1}: {2}" -f $forbidden.Count, $forbidden[0].ImgName, $forbidden[0].CommandLine)
}

# ---------------------------------------------------------------- CHECK 7: run length
$famEvents = New-Object System.Collections.ArrayList
foreach ($e in $allMain) { [void]$famEvents.Add($e.TimeCreated) }
foreach ($f in $families) {
    foreach ($p in @($procs | Where-Object { $_.ImgName -eq $f.Life -and (& $f.LifeMain $_) })) {
        if ($termByGuid.ContainsKey($p.ProcessGuid)) { [void]$famEvents.Add($termByGuid[$p.ProcessGuid]) }
    }
}
if ($famEvents.Count -ge 2) {
    $sorted = @($famEvents | Sort-Object)
    $durSec = ($sorted[-1] - $sorted[0]).TotalSeconds
    $durTxt = "{0:N1} min ({1} .. {2})" -f ($durSec / 60), $sorted[0].ToString('HH:mm:ss'), $sorted[-1].ToString('HH:mm:ss')
    if ($durSec -le ($MaxRunMinutes * 60 + 30)) {
        Add-Check "Run length <= $MaxRunMinutes min" 'PASS' $durTxt
    } else {
        Add-Check "Run length <= $MaxRunMinutes min" 'FAIL' $durTxt
    }
} else {
    Add-Check "Run length <= $MaxRunMinutes min" 'WARN' 'Not enough app events to measure'
}

# ---------------------------------------------------------------- CHECK 8: timing variability
# Action timestamps: app launches + closes + document saves + Chrome DNS lookups.
$stamps = New-Object System.Collections.ArrayList
foreach ($t in $famEvents) { [void]$stamps.Add($t) }
foreach ($d in $dns)   { if ($d.ImgName -eq 'chrome.exe') { [void]$stamps.Add($d.TimeCreated) } }
foreach ($x in $files) { if ($x.ImgName -like 'soffice*' -and $x.TargetFilename -match '\.(odt|ods|docx?|xlsx?|rtf|txt)$') { [void]$stamps.Add($x.TimeCreated) } }
# merge bursts (<3 s apart = one action), then gaps between burst starts
$starts = New-Object System.Collections.ArrayList
$prev = $null
foreach ($t in @($stamps | Sort-Object)) {
    if ($prev -eq $null -or ($t - $prev).TotalSeconds -ge 3) { [void]$starts.Add($t) }
    $prev = $t
}
$gaps = New-Object System.Collections.ArrayList
for ($i = 1; $i -lt $starts.Count; $i++) { [void]$gaps.Add(($starts[$i] - $starts[$i - 1]).TotalSeconds) }
Write-Host ""
Write-Host ("--- Action gaps ({0} actions, {1} gaps) ---" -f $starts.Count, $gaps.Count) -ForegroundColor Cyan
if ($gaps.Count -ge 1) {
    $st = Get-Stats -Values ([double[]]$gaps.ToArray())
    Write-Host ("  min {0:N1}s  median {1:N1}s  mean {2:N1}s  p95 {3:N1}s  sd {4:N1}s  max {5:N1}s  CV {6:N2}" -f $st.Min, $st.Median, $st.Mean, $st.P95, $st.SD, $st.Max, $st.CV)
    # identical-to-the-second gaps: a scripted sleep(n) shows up as repeated whole-second gaps
    $roundedGaps = @($gaps | ForEach-Object { [int][math]::Round($_) })
    $dupExtra = 0
    foreach ($g in @($roundedGaps | Group-Object | Where-Object { $_.Count -gt 1 })) { $dupExtra += ($g.Count - 1) }
    Write-Host ("  identical-to-the-second gaps: {0} repeat(s);  longest gap {1:N0}s" -f $dupExtra, $st.Max)
    Show-Histogram -Gaps ([double[]]$gaps.ToArray())
    if ($gaps.Count -ge 4) {
        $nearMed = @($gaps | Where-Object { [math]::Abs($_ - $st.Median) -le 0.05 * $st.Median }).Count
        if ($st.CV -ge 0.6 -and $nearMed -le ($gaps.Count / 2)) {
            Add-Check 'Timing not metronomic' 'PASS' ("{0} gaps, CV {1:N2} (team target >= 0.60), median {2:N1}s, p95 {3:N1}s, range {4:N1}-{5:N1}s" -f $gaps.Count, $st.CV, $st.Median, $st.P95, $st.Min, $st.Max)
        } elseif ($st.CV -lt 0.6) {
            Add-Check 'Timing not metronomic' 'WARN' ("CV {0:N2} is below the team target of 0.60 ({1} gaps, median {2:N1}s) - our own heuristic, not a judge threshold" -f $st.CV, $gaps.Count, $st.Median)
        } else {
            Add-Check 'Timing not metronomic' 'WARN' ("{0} of {1} gaps are within 5% of the median - delays look too regular" -f $nearMed, $gaps.Count)
        }
        if ($dupExtra -le 2) {
            Add-Check 'Repeated gap lengths' 'PASS' ("{0} repeated gap length(s) at 1 s resolution (target <= 2)" -f $dupExtra)
        } else {
            Add-Check 'Repeated gap lengths' 'WARN' ("{0} repeated gap length(s) at 1 s resolution (target <= 2) - looks scripted" -f $dupExtra)
        }
        if ($st.Max -ge 60) {
            Add-Check 'Long pause visible' 'PASS' ("longest gap {0:N0}s (target >= 60s)" -f $st.Max)
        } else {
            Add-Check 'Long pause visible' 'WARN' ("longest gap is only {0:N0}s - no pause over 60 s is visible in Sysmon" -f $st.Max)
        }
    } else {
        Add-Check 'Timing not metronomic' 'INFO' ("Only {0} gaps - run a longer session for meaningful statistics" -f $gaps.Count)
    }
} else {
    Add-Check 'Timing not metronomic' 'INFO' 'Not enough actions to compute gaps'
}

# ---------------------------------------------------------------- CHECK 9: document saved
$docs = @($files | Where-Object { $_.ImgName -like 'soffice*' -and $_.TargetFilename -match '\.(odt|ods|docx?|xlsx?)$' })
if ($docs.Count -gt 0) {
    Add-Check 'Office saved a document' 'PASS' ("Sysmon ID 11: {0}" -f (($docs | Select-Object -ExpandProperty TargetFilename -Unique | Select-Object -First 2) -join ', '))
} else {
    $dirs = @("$env:USERPROFILE\Documents", "$env:USERPROFILE\Desktop") | Where-Object { Test-Path $_ }
    $disk = @()
    foreach ($d in $dirs) {
        $disk += @(Get-ChildItem -Path $d -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in '.odt', '.ods', '.docx', '.doc', '.xlsx' -and $_.LastWriteTime -ge $Since -and $_.LastWriteTime -le $Until.AddMinutes(2) })
    }
    if ($disk.Count -gt 0) {
        $newest = $disk | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        Add-Check 'Office saved a document' 'PASS' ("{0} ({1} bytes) - on disk; FileCreate not in Sysmon config" -f $newest.FullName, $newest.Length)
    } else {
        $wr = @($famLaunches['Writer/Calc'])
        if ($wr.Count -gt 0) { Add-Check 'Office saved a document' 'FAIL' 'Office ran but no .odt/.ods/.docx found (Sysmon or Documents/Desktop)' }
    }
}

# ---------------------------------------------------------------- CHECKS 10-13: the agent itself + stray shells
# The agent interpreter: pythonw.exe (or python.exe) running agent.pyw / agent.py / sequencer.py.
$agentProc = @($procs | Where-Object { $_.ImgName -match '^pythonw?\.exe$' -and $_.CommandLine -match '(agent|sequencer)\.pyw?\b' } |
    Sort-Object TimeCreated) | Select-Object -First 1
$agentEnd = $null
if ($agentProc -and $termByGuid.ContainsKey($agentProc.ProcessGuid)) { $agentEnd = $termByGuid[$agentProc.ProcessGuid] }
$appTimes = @($famEvents | Sort-Object)
# the run span: agent start .. agent end (fallback: first app launch .. last app event + 90 s)
$runLo = if ($agentProc) { $agentProc.TimeCreated } elseif ($appTimes.Count -gt 0) { $appTimes[0] } else { $Since }
$runHi = if ($agentEnd) { $agentEnd } elseif ($appTimes.Count -gt 0) { $appTimes[$appTimes.Count - 1].AddSeconds(90) } else { $Until }

# CHECK 10: the agent's own parent
if ($agentProc) {
    $agentCmd = if ($agentProc.CommandLine.Length -gt 90) { $agentProc.CommandLine.Substring(0, 90) + '...' } else { $agentProc.CommandLine }
    if ($agentProc.ParName -eq 'explorer.exe') {
        Add-Check 'Agent parent' 'PASS' ("explorer.exe -> {0}  ({1})" -f $agentProc.ImgName, $agentCmd)
    } else {
        Add-Check 'Agent parent' 'FAIL' ("agent {0} was started by {1}, not explorer.exe  ({2})" -f $agentProc.ImgName, $agentProc.ParName, $agentCmd)
    }
} else {
    Add-Check 'Agent parent' 'INFO' 'Agent interpreter (pythonw/python running agent.pyw / agent.py / sequencer.py) is not in the window - start -Since a minute before the run'
}

# CHECK 11: the agent interpreter should not spawn anything (apps hang off explorer.exe)
if ($agentProc) {
    $agentKids = @($procs | Where-Object { $_.ParentProcessGuid -eq $agentProc.ProcessGuid })
    if ($agentKids.Count -eq 0) {
        Add-Check 'Agent child processes' 'PASS' 'The agent interpreter created no child processes'
    } else {
        $kidNames = ($agentKids | ForEach-Object { $_.ImgName } | Sort-Object -Unique) -join ', '
        Add-Check 'Agent child processes' 'WARN' ("The agent interpreter created {0} child process(es): {1}" -f $agentKids.Count, $kidNames)
    }
}

# CHECK 12: agent lifetime (live run target 15:00 +/- 30 s)
if ($agentProc -and $agentEnd) {
    $life = ($agentEnd - $agentProc.TimeCreated).TotalSeconds
    $lifeTxt = "{0:N0} s = {1:N1} min ({2} .. {3})" -f $life, ($life / 60), $agentProc.TimeCreated.ToString('HH:mm:ss'), $agentEnd.ToString('HH:mm:ss')
    if ($life -ge 870 -and $life -le 930) {
        Add-Check 'Agent lifetime ~15 min' 'PASS' $lifeTxt
    } else {
        Add-Check 'Agent lifetime ~15 min' 'WARN' ("{0} - outside 15:00 +/- 30 s (expected for a rehearsal with another budget)" -f $lifeTxt)
    }
} elseif ($agentProc) {
    Add-Check 'Agent lifetime ~15 min' 'INFO' 'No terminate event for the agent: still running, or the run was cut by the window'
}

# CHECK 13: shells / interpreters created during the run (judges look for exactly these)
$shellNames = @('cmd.exe', 'powershell.exe', 'pwsh.exe', 'conhost.exe', 'wsl.exe', 'wslhost.exe', 'wscript.exe', 'cscript.exe', 'python.exe', 'py.exe')
$shellHits = @($procs | Where-Object { $_.ImgName -in $shellNames -and $_.TimeCreated -ge $runLo -and $_.TimeCreated -le $runHi } | Sort-Object TimeCreated)
if ($shellHits.Count -eq 0) {
    Add-Check 'No shell / interpreter processes' 'PASS' ("No cmd / powershell / conhost / wsl / script-interpreter process was created between {0} and {1}" -f $runLo.ToString('HH:mm:ss'), $runHi.ToString('HH:mm:ss'))
} else {
    $topShell = ($shellHits | Select-Object -First 3 | ForEach-Object { "{0} {1} <- {2}" -f $_.TimeCreated.ToString('HH:mm:ss'), $_.ImgName, $_.ParName }) -join '; '
    Add-Check 'No shell / interpreter processes' 'WARN' ("{0} created during the run, e.g. {1} - close terminals and monitoring tools before the live run" -f $shellHits.Count, $topShell)
    Write-Host ""
    Write-Host "--- Shell / interpreter processes created during the run ---" -ForegroundColor Yellow
    $shellHits | ForEach-Object { Write-Host ("  {0}  {1} <- {2}  {3}" -f $_.TimeCreated.ToString('HH:mm:ss'), $_.ImgName, $_.ParName, $_.CommandLine) }
}

# ---------------------------------------------------------------- process chains (evidence for Q&A)
Write-Host ""
Write-Host "--- Process chains (evidence) ---" -ForegroundColor Cyan
function Get-Kids { param($Guid) @($procs | Where-Object { $_.ParentProcessGuid -eq $Guid }) }
if ($agentProc) {
    Write-Host ("  {0}  {1} -> {2}  (the agent itself)" -f $agentProc.TimeCreated.ToString('HH:mm:ss'), $agentProc.ParName, $agentProc.ImgName)
}
foreach ($e in @($allMain | Sort-Object TimeCreated)) {
    $line = "{0}  {1} -> {2}" -f $e.TimeCreated.ToString('HH:mm:ss'), $e.ParName, $e.ImgName
    if ($e.ImgName -eq 'chrome.exe') {
        $kids = @(Get-Kids $e.ProcessGuid)
        $line += ("  (+{0} child processes, normal multiprocess noise)" -f $kids.Count)
    } else {
        foreach ($k in (Get-Kids $e.ProcessGuid)) {
            $line += " -> " + $k.ImgName
            foreach ($k2 in (Get-Kids $k.ProcessGuid)) { $line += " -> " + $k2.ImgName }
        }
    }
    Write-Host ("  " + $line)
}

# ---------------------------------------------------------------- report
Write-Host ""
Write-Host "=== Checklist ===" -ForegroundColor Cyan
foreach ($r in $Results) {
    $color = switch ($r.Status) { 'PASS' { 'Green' } 'FAIL' { 'Red' } 'WARN' { 'Yellow' } default { 'Gray' } }
    Write-Host ("  [{0}] {1,-38} {2}" -f $r.Status, $r.Check, $r.Detail) -ForegroundColor $color
}
$nFail = @($Results | Where-Object { $_.Status -eq 'FAIL' }).Count
$nWarn = @($Results | Where-Object { $_.Status -eq 'WARN' }).Count
Write-Host ""
if ($nFail -eq 0) {
    Write-Host ("RESULT: PASS  ({0} warning(s))" -f $nWarn) -ForegroundColor Green
} else {
    Write-Host ("RESULT: FAIL  ({0} failure(s), {1} warning(s))" -f $nFail, $nWarn) -ForegroundColor Red
}

# ---------------------------------------------------------------- evidence files
try {
    if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
    $stamp = $ScriptStart.ToString('yyyyMMdd_HHmmss')
    $procs | Select-Object TimeCreated, ImgName, ParName, ProcessId, ParentProcessId, CurrentDirectory, CommandLine |
        Export-Csv -Path (Join-Path $OutDir "sysmon_process_$stamp.csv") -NoTypeInformation -Encoding UTF8
    $Results | Export-Csv -Path (Join-Path $OutDir "checklist_$stamp.csv") -NoTypeInformation -Encoding UTF8
    if ($gaps.Count -gt 0) {
        $gaps | ForEach-Object { [pscustomobject]@{ GapSeconds = [math]::Round($_, 2) } } |
            Export-Csv -Path (Join-Path $OutDir "gaps_$stamp.csv") -NoTypeInformation -Encoding UTF8
    }
    Write-Host ("Evidence written to {0}" -f $OutDir)
} catch {
    Write-Host ("Could not write evidence files: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
}

if ($nFail -gt 0) { exit 1 } else { exit 0 }