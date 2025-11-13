# Install Everything (es.exe) with PowerShell

Everything (by Voidtools) indexes your drives so searches are instant. Bento Project Packer uses the `es.exe` CLI to find samples quickly. Installing it is a big speed-up on large libraries.

## Quick Install (winget)

Run these commands in PowerShell (new window recommended):

```
# 1) Install Everything via winget (silent)
winget install -e --id voidtools.Everything --accept-package-agreements --accept-source-agreements

# 2) Verify es.exe is available
es -version
where es

# 3) If es.exe isn’t found, add the install folder to PATH (pick one that exists)
$dirs = @(
  'C:\\Program Files\\Everything',
  'C:\\Program Files (x86)\\Everything',
  'C:\\Program Files\\Voidtools\\Everything',
  'C:\\Program Files (x86)\\Voidtools\\Everything'
)
$EverythingDir = $dirs | Where-Object { Test-Path (Join-Path $_ 'es.exe') } | Select-Object -First 1
if ($EverythingDir) {
  [Environment]::SetEnvironmentVariable('Path', ($env:Path + ';' + $EverythingDir), 'User')
  Write-Host "Added to PATH: $EverythingDir (open a new terminal)" -ForegroundColor Green
} else {
  Write-Host "Could not locate es.exe automatically. Check Everything install folder and add to PATH." -ForegroundColor Yellow
}

# 4) Ensure the Everything service is running (optional but recommended)
$svc = Get-Service -Name 'Everything' -ErrorAction SilentlyContinue
if ($svc) {
  if ($svc.Status -ne 'Running') { Start-Service 'Everything' }
  Write-Host "Everything service is running." -ForegroundColor Green
} else {
  Write-Host "Everything service not found. Launch Everything.exe and enable 'Everything Service' in Options." -ForegroundColor Yellow
}

# 5) Test a scoped search (replace D:\\Samples and filename)
es path:"D:\\Samples" "kick.wav"
```

Notes
- If `es` still isn’t recognized, close and reopen your terminal so PATH changes take effect.
- Bento Project Packer auto-detects `es.exe` and falls back to a normal scan if not found.

## Portable (optional)
If you prefer portable, download the portable ZIP from the Voidtools site, extract to a folder (e.g., `C:\\Tools\\Everything`), then add that folder to PATH as shown above and run `Everything.exe` once to build the index.
