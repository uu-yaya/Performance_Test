param(
  [string]$ConfigPath = "D:\desktop_pet\tools\perf_collect_config.example.json",
  [string]$OutputDir = "D:\desktop_pet\perf_runs"
)

$ErrorActionPreference = "Stop"
$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$collector = Join-Path $PSScriptRoot "collect_desktop_pet_perf.ps1"

foreach ($scenario in $config.scenarios) {
  & $collector `
    -ProductId $config.product_id `
    -ProductName $config.product_name `
    -ProcessName $config.process_names `
    -Scenario $scenario.name `
    -DurationSec ([int]$scenario.duration_sec) `
    -IntervalSec ([int]$scenario.interval_sec) `
    -OutputDir $OutputDir `
    -Notes $scenario.notes
}
