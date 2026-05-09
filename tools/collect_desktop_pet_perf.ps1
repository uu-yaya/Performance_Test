param(
  [Parameter(Mandatory=$true)]
  [string]$ProductId,

  [Parameter(Mandatory=$true)]
  [string]$ProductName,

  [Parameter(Mandatory=$true)]
  [string[]]$ProcessName,

  [Parameter(Mandatory=$true)]
  [string]$Scenario,

  [int]$DurationSec = 600,
  [int]$IntervalSec = 1,
  [string]$OutputDir = "D:\desktop_pet\perf_runs",
  [string]$Notes = ""
)

$ErrorActionPreference = "Continue"

function Normalize-ProcessName {
  param([string]$Name)
  return [System.IO.Path]::GetFileNameWithoutExtension($Name.Trim())
}

function Percentile {
  param([double[]]$Values, [double]$P)
  $clean = @($Values | Where-Object { $_ -ne $null -and -not [double]::IsNaN($_) } | Sort-Object)
  if ($clean.Count -eq 0) { return $null }
  if ($clean.Count -eq 1) { return [math]::Round($clean[0], 3) }
  $rank = ($P / 100.0) * ($clean.Count - 1)
  $low = [math]::Floor($rank)
  $high = [math]::Ceiling($rank)
  if ($low -eq $high) { return [math]::Round($clean[$low], 3) }
  $weight = $rank - $low
  return [math]::Round(($clean[$low] * (1 - $weight)) + ($clean[$high] * $weight), 3)
}

function AverageOrNull {
  param([double[]]$Values)
  $clean = @($Values | Where-Object { $_ -ne $null -and -not [double]::IsNaN($_) })
  if ($clean.Count -eq 0) { return $null }
  return [math]::Round((($clean | Measure-Object -Average).Average), 3)
}

function MaxOrNull {
  param([double[]]$Values)
  $clean = @($Values | Where-Object { $_ -ne $null -and -not [double]::IsNaN($_) })
  if ($clean.Count -eq 0) { return $null }
  return [math]::Round((($clean | Measure-Object -Maximum).Maximum), 3)
}

function Get-GpuCountersForPids {
  param([int[]]$Pids)
  $result = @{
    GpuUtilPercent = 0.0
    DedicatedBytes = 0.0
    SharedBytes = 0.0
    EngineTypes = @()
  }
  if ($Pids.Count -eq 0) { return $result }

  $enginePaths = @()
  $memoryPaths = @()
  try {
    $enginePaths = (Get-Counter -ListSet "GPU Engine" -ErrorAction Stop).PathsWithInstances
  } catch {}
  try {
    $memoryPaths = (Get-Counter -ListSet "GPU Process Memory" -ErrorAction Stop).PathsWithInstances
  } catch {}

  foreach ($processId in $Pids) {
    $pidPattern = "pid_${processId}_"
    $matchedEngine = @($enginePaths | Where-Object { $_ -like "*$pidPattern*" -and $_ -like "*\Utilization Percentage" })
    if ($matchedEngine.Count -gt 0) {
      try {
        $samples = (Get-Counter -Counter $matchedEngine -ErrorAction Stop).CounterSamples
        foreach ($sample in $samples) {
          $result.GpuUtilPercent += [double]$sample.CookedValue
          if ($sample.Path -match "engtype_([^)\s]+)") {
            $result.EngineTypes += $Matches[1]
          }
        }
      } catch {}
    }

    $dedicated = @($memoryPaths | Where-Object { $_ -like "*$pidPattern*" -and $_ -like "*\Dedicated Usage" })
    if ($dedicated.Count -gt 0) {
      try {
        $samples = (Get-Counter -Counter $dedicated -ErrorAction Stop).CounterSamples
        foreach ($sample in $samples) { $result.DedicatedBytes += [double]$sample.CookedValue }
      } catch {}
    }

    $shared = @($memoryPaths | Where-Object { $_ -like "*$pidPattern*" -and $_ -like "*\Shared Usage" })
    if ($shared.Count -gt 0) {
      try {
        $samples = (Get-Counter -Counter $shared -ErrorAction Stop).CounterSamples
        foreach ($sample in $samples) { $result.SharedBytes += [double]$sample.CookedValue }
      } catch {}
    }
  }
  $result.GpuUtilPercent = [math]::Round($result.GpuUtilPercent, 3)
  return $result
}

function Get-PerfProcByPid {
  param([int[]]$Pids)
  if ($Pids.Count -eq 0) { return @{} }
  $map = @{}
  try {
    $items = Get-CimInstance Win32_PerfFormattedData_PerfProc_Process |
      Where-Object { $Pids -contains [int]$_.IDProcess }
    foreach ($item in $items) {
      $map[[int]$item.IDProcess] = $item
    }
  } catch {}
  return $map
}

$normalizedNames = @($ProcessName | ForEach-Object { Normalize-ProcessName $_ } | Where-Object { $_ })
$logicalCpu = [Environment]::ProcessorCount
$runStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$safeProduct = ($ProductName -replace '[\\/:*?"<>| ]+', '_')
$runDir = Join-Path $OutputDir "${runStamp}_${ProductId}_${safeProduct}_${Scenario}"
New-Item -ItemType Directory -Path $runDir -Force | Out-Null

$rawPath = Join-Path $runDir "raw_samples.csv"
$summaryPath = Join-Path $runDir "summary.csv"
$metaPath = Join-Path $runDir "metadata.json"

$metadata = [ordered]@{
  product_id = $ProductId
  product_name = $ProductName
  process_names = $normalizedNames
  scenario = $Scenario
  duration_sec = $DurationSec
  interval_sec = $IntervalSec
  started_at = (Get-Date).ToString("s")
  computer_name = $env:COMPUTERNAME
  logical_cpu_count = $logicalCpu
  notes = $Notes
}
$metadata | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $metaPath -Encoding UTF8

$rows = New-Object System.Collections.Generic.List[object]
$previousCpuByPid = @{}
$previousTsByPid = @{}
$endTime = (Get-Date).AddSeconds($DurationSec)
$sampleIndex = 0

Write-Host "Collecting $ProductName / $Scenario for $DurationSec seconds..."
Write-Host "Process names: $($normalizedNames -join ', ')"
Write-Host "Output: $runDir"

while ((Get-Date) -lt $endTime) {
  $timestamp = Get-Date
  $processes = @()
  foreach ($name in $normalizedNames) {
    $processes += @(Get-Process -Name $name -ErrorAction SilentlyContinue)
  }
  $processes = @($processes | Sort-Object Id -Unique)
  $pids = @($processes | ForEach-Object { [int]$_.Id })
  $gpu = Get-GpuCountersForPids -Pids $pids
  $perfProc = Get-PerfProcByPid -Pids $pids

  $totalCpuPct = 0.0
  $totalWorkingSetMb = 0.0
  $totalPrivateMb = 0.0
  $totalHandles = 0
  $totalThreads = 0
  $totalIoReadBps = 0.0
  $totalIoWriteBps = 0.0
  $tcpConnectionCount = 0
  $processDetails = @()

  foreach ($p in $processes) {
    $processId = [int]$p.Id
    $cpuPct = 0.0
    if ($p.CPU -ne $null -and $previousCpuByPid.ContainsKey($processId)) {
      $elapsed = ($timestamp - $previousTsByPid[$processId]).TotalSeconds
      if ($elapsed -gt 0) {
        $cpuPct = (($p.CPU - $previousCpuByPid[$processId]) / $elapsed) * 100.0 / $logicalCpu
      }
    }
    if ($p.CPU -ne $null) {
      $previousCpuByPid[$processId] = [double]$p.CPU
      $previousTsByPid[$processId] = $timestamp
    }

    $totalCpuPct += [math]::Max(0, $cpuPct)
    $totalWorkingSetMb += $p.WorkingSet64 / 1MB
    $totalPrivateMb += $p.PrivateMemorySize64 / 1MB
    $totalHandles += $p.HandleCount
    $totalThreads += $p.Threads.Count

    if ($perfProc.ContainsKey($processId)) {
      $totalIoReadBps += [double]$perfProc[$processId].IOReadBytesPerSec
      $totalIoWriteBps += [double]$perfProc[$processId].IOWriteBytesPerSec
    }

    try {
      $tcpConnectionCount += @(Get-NetTCPConnection -OwningProcess $processId -ErrorAction SilentlyContinue).Count
    } catch {}

    $processDetails += "$($p.ProcessName):$processId"
  }

  $row = [ordered]@{
    sample_index = $sampleIndex
    timestamp = $timestamp.ToString("s")
    product_id = $ProductId
    product_name = $ProductName
    scenario = $Scenario
    process_count = $processes.Count
    process_details = ($processDetails -join ";")
    cpu_percent = [math]::Round($totalCpuPct, 3)
    gpu_percent = $gpu.GpuUtilPercent
    memory_working_set_mb = [math]::Round($totalWorkingSetMb, 3)
    memory_private_mb = [math]::Round($totalPrivateMb, 3)
    dedicated_vram_mb = [math]::Round($gpu.DedicatedBytes / 1MB, 3)
    shared_vram_mb = [math]::Round($gpu.SharedBytes / 1MB, 3)
    handles = $totalHandles
    threads = $totalThreads
    io_read_kbps = [math]::Round($totalIoReadBps / 1KB, 3)
    io_write_kbps = [math]::Round($totalIoWriteBps / 1KB, 3)
    tcp_connections = $tcpConnectionCount
    gpu_engine_types = (($gpu.EngineTypes | Sort-Object -Unique) -join ";")
  }
  $rows.Add([pscustomobject]$row) | Out-Null
  $sampleIndex += 1
  Start-Sleep -Seconds $IntervalSec
}

$rows | Export-Csv -LiteralPath $rawPath -NoTypeInformation -Encoding UTF8

$cpu = @($rows | ForEach-Object { [double]$_.cpu_percent })
$gpuVals = @($rows | ForEach-Object { [double]$_.gpu_percent })
$mem = @($rows | ForEach-Object { [double]$_.memory_working_set_mb })
$priv = @($rows | ForEach-Object { [double]$_.memory_private_mb })
$vram = @($rows | ForEach-Object { [double]$_.dedicated_vram_mb })
$handles = @($rows | ForEach-Object { [double]$_.handles })
$threads = @($rows | ForEach-Object { [double]$_.threads })
$ioRead = @($rows | ForEach-Object { [double]$_.io_read_kbps })
$ioWrite = @($rows | ForEach-Object { [double]$_.io_write_kbps })
$tcp = @($rows | ForEach-Object { [double]$_.tcp_connections })

$summary = [pscustomobject][ordered]@{
  product_id = $ProductId
  product_name = $ProductName
  scenario = $Scenario
  started_at = $metadata.started_at
  duration_sec = $DurationSec
  interval_sec = $IntervalSec
  sample_count = $rows.Count
  cpu_avg_percent = AverageOrNull $cpu
  cpu_p95_percent = Percentile $cpu 95
  cpu_max_percent = MaxOrNull $cpu
  gpu_avg_percent = AverageOrNull $gpuVals
  gpu_p95_percent = Percentile $gpuVals 95
  gpu_max_percent = MaxOrNull $gpuVals
  memory_avg_mb = AverageOrNull $mem
  memory_max_mb = MaxOrNull $mem
  private_memory_avg_mb = AverageOrNull $priv
  dedicated_vram_avg_mb = AverageOrNull $vram
  dedicated_vram_max_mb = MaxOrNull $vram
  handles_start = if ($handles.Count -gt 0) { $handles[0] } else { $null }
  handles_end = if ($handles.Count -gt 0) { $handles[-1] } else { $null }
  handles_delta = if ($handles.Count -gt 0) { $handles[-1] - $handles[0] } else { $null }
  threads_start = if ($threads.Count -gt 0) { $threads[0] } else { $null }
  threads_end = if ($threads.Count -gt 0) { $threads[-1] } else { $null }
  threads_delta = if ($threads.Count -gt 0) { $threads[-1] - $threads[0] } else { $null }
  io_read_avg_kbps = AverageOrNull $ioRead
  io_write_avg_kbps = AverageOrNull $ioWrite
  tcp_connections_max = MaxOrNull $tcp
  raw_csv = $rawPath
  metadata_json = $metaPath
  notes = $Notes
}

$summary | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8
Write-Host "Done."
Write-Host "Raw samples: $rawPath"
Write-Host "Summary: $summaryPath"
