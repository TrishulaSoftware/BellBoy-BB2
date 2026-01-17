param(
  [string]$Root = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

$shaFile = Join-Path $Root "docs\SHA256SUMS.txt"
if (-not (Test-Path -LiteralPath $shaFile)) { throw "Missing $shaFile" }

$expected = @{}
Get-Content -LiteralPath $shaFile | ForEach-Object {
  if ($_ -match '^\s*([0-9a-fA-F]{64})\s\s+(.+?)\s*$') {
    $expected[$Matches[2]] = $Matches[1].ToLowerInvariant()
  }
}

$fail = $false
foreach ($k in $expected.Keys) {
  $p = Join-Path $Root $k
  if (-not (Test-Path -LiteralPath $p)) {
    Write-Host "[MISSING] $k"
    $fail = $true
    continue
  }
  $h = (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash.ToLowerInvariant()
  if ($h -ne $expected[$k]) {
    Write-Host "[BADHASH] $k"
    Write-Host "  expected: $($expected[$k])"
    Write-Host "  actual:   $h"
    $fail = $true
  } else {
    Write-Host "[OK] $k"
  }
}

if ($fail) { exit 2 } else { Write-Host "ALL OK" }
