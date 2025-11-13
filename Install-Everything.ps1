[CmdletBinding()]
param()

Write-Host "[INFO] Installing Everything (Voidtools) for fast sample search..." -ForegroundColor Cyan

# Try winget install (silent)
$winget = Get-Command winget -ErrorAction SilentlyContinue
if ($winget) {
  Write-Host "[INFO] Using winget to install Everything..." -ForegroundColor Cyan
  try {
    winget install -e --id voidtools.Everything --accept-package-agreements --accept-source-agreements | Out-Null
  } catch {
    Write-Warning "winget install reported an error: $_"
  }
} else {
  Write-Warning "winget not found. Please install Everything from https://www.voidtools.com/ or add winget."
}

# Locate es.exe
$candidateDirs = @(
  'C:\\Program Files\\Everything',
  'C:\\Program Files (x86)\\Everything',
  'C:\\Program Files\\Voidtools\\Everything',
  'C:\\Program Files (x86)\\Voidtools\\Everything'
)
$EverythingDir = $candidateDirs | Where-Object { Test-Path (Join-Path $_ 'es.exe') } | Select-Object -First 1
if (-not $EverythingDir) {
  $esCmd = Get-Command es -ErrorAction SilentlyContinue
  if ($esCmd) { $EverythingDir = Split-Path -Parent $esCmd.Source }
}

if ($EverythingDir) {
  Write-Host "[INFO] Found es.exe in: $EverythingDir" -ForegroundColor Green
  # Ensure es.exe folder is on PATH for the current user
  $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
  if (-not $userPath) { $userPath = '' }
  if (-not ($userPath -split ';' | Where-Object { $_ -ieq $EverythingDir })) {
    [Environment]::SetEnvironmentVariable('Path', ($userPath + ';' + $EverythingDir).Trim(';'), 'User')
    Write-Host "[INFO] Added to PATH for current user. Open a new terminal to pick it up." -ForegroundColor Green
  } else {
    Write-Host "[INFO] es.exe directory already in PATH." -ForegroundColor DarkGreen
  }
} else {
  Write-Warning "Could not locate es.exe. If Everything is installed, add its folder to PATH manually."
}

# Start Everything service if available
$svc = Get-Service -Name 'Everything' -ErrorAction SilentlyContinue
if ($svc) {
  if ($svc.Status -ne 'Running') {
    try { Start-Service 'Everything'; Write-Host "[INFO] Started Everything service." -ForegroundColor Green } catch { Write-Warning $_ }
  } else {
    Write-Host "[INFO] Everything service already running." -ForegroundColor Green
  }
} else {
  Write-Host "[INFO] Everything service not found. Launch Everything.exe and enable 'Everything Service' in Options for best performance." -ForegroundColor Yellow
}

# Verify
$esVersion = $null
try { $esVersion = & es -version 2>$null } catch {}
if ($esVersion) {
  Write-Host "[OK] es.exe is available: $esVersion" -ForegroundColor Green
} else {
  Write-Host "[WARN] 'es' is not recognized yet. After reopening your terminal, run: es -version" -ForegroundColor Yellow
}

Write-Host "[DONE] Everything setup attempt complete." -ForegroundColor Cyan
