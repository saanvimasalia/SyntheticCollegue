param([string]$Script = "test.py")
$cores = [Environment]::ProcessorCount
Write-Host "Launching: python $Script  (close Chrome first; then click away)"
$p = Start-Process python -ArgumentList $Script -PassThru
Start-Sleep -Milliseconds 500
$rows = @(); $prev = $null; $prevT = Get-Date
while (-not $p.HasExited) {
    Start-Sleep -Seconds 1
    $p.Refresh()
    if ($p.HasExited) { break }
    $now = Get-Date
    $ram = [math]::Round($p.WorkingSet64/1MB, 1)
    $cpuSec = $p.TotalProcessorTime.TotalSeconds
    $cpuPct = $null
    if ($prev -ne $null) {
        $dt = ($now - $prevT).TotalSeconds
        if ($dt -gt 0) { $cpuPct = [math]::Round((($cpuSec - $prev)/$dt)/$cores*100, 1) }
    }
    $rows += [pscustomobject]@{ t=$now.ToString("HH:mm:ss"); cpuPct=$cpuPct; ramMB=$ram }
    $prev = $cpuSec; $prevT = $now
}
$csv = ".\validation\resource_$(Get-Date -Format yyyyMMdd_HHmmss).csv"
$rows | Export-Csv $csv -NoTypeInformation
$c = $rows | Where-Object {$_.cpuPct -ne $null} | Measure-Object cpuPct -Average -Maximum
$r = $rows | Measure-Object ramMB -Average -Maximum
Write-Host ("`nsamples {0} -> {1}" -f $rows.Count, $csv)
Write-Host ("agent CPU%  avg {0:N1}  peak {1:N1}" -f $c.Average, $c.Maximum)
Write-Host ("agent RAM MB avg {0:N1}  peak {1:N1}" -f $r.Average, $r.Maximum)