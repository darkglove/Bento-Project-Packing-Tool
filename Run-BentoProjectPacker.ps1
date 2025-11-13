<#
Bento Project Packer — Runner Wrapper

Usage:
  .\Run-BentoProjectPacker.ps1 <ProjectXmlOrDir> [-DryRun] [-SearchRoot <path>] [-Report <path>]

Notes:
  - <ProjectXmlOrDir> can be a project XML file or a directory containing one (e.g., 'project.xml').
  - If -SearchRoot is not provided, roots are read from 'SampleRoots.txt' (one path per line) in this folder; otherwise falls back to %USERPROFILE%\Music.
  - A timestamped report is always written beside the project XML (and a copy in the current directory). You can override with -Report.
  - This tool is beta. Run on copies of your projects and keep backups. Use -DryRun first to preview changes.
#>

$ProjectXmlPath = $null
$SearchRoot = $null
$Simulate = $false
$ReportPath = $null

for ($i = 0; $i -lt $args.Count; $i++) {
  $a = $args[$i]
  if ($a -eq '-DryRun') {
    $Simulate = $true; continue
  } elseif ($a.StartsWith('-SearchRoot=')) {
    $SearchRoot = $a.Substring(12); continue
  } elseif ($a -eq '-SearchRoot') {
    if ($i + 1 -lt $args.Count) { $i++; $SearchRoot = $args[$i]; continue }
    Write-Host "-SearchRoot requires a value" -ForegroundColor Red; exit 1
  } elseif ($a.StartsWith('-Report=')) {
    $ReportPath = $a.Substring(8); continue
  } elseif ($a -eq '-Report') {
    if ($i + 1 -lt $args.Count) { $i++; $ReportPath = $args[$i]; continue }
    Write-Host "-Report requires a value" -ForegroundColor Red; exit 1
  } else {
    if (-not $ProjectXmlPath) { $ProjectXmlPath = $a; continue }
    Write-Host ("Unknown argument: {0}" -f $a) -ForegroundColor Yellow
  }
}

if (-not $ProjectXmlPath) {
  Write-Host "Usage: .\Run-BentoProjectPacker.ps1 <ProjectXmlOrDir> [-DryRun] [-SearchRoot <path>] [-Report <path>]" -ForegroundColor Yellow
  exit 1
}

# Accept a project directory or an explicit XML path
try {
  $resolvedInput = (Resolve-Path -LiteralPath $ProjectXmlPath -ErrorAction Stop).Path
} catch {
  Write-Host ("Project path not found: {0}" -f $ProjectXmlPath) -ForegroundColor Red
  exit 1
}
if (Test-Path -LiteralPath $resolvedInput -PathType Container) {
  $dir = $resolvedInput
  $candidate = Join-Path $dir 'project.xml'
  if (Test-Path -LiteralPath $candidate -PathType Leaf) {
    $ProjectXmlPath = $candidate
  } else {
    $xmls = Get-ChildItem -LiteralPath $dir -Filter *.xml -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
    if ($xmls.Count -eq 1) {
      $ProjectXmlPath = $xmls[0]
    } elseif ($xmls.Count -gt 1) {
      Write-Host ("Multiple XML files found in '{0}'. Please specify the XML file explicitly." -f $dir) -ForegroundColor Yellow
      $xmls | ForEach-Object { Write-Host " - $_" -ForegroundColor DarkGray }
      exit 1
    } else {
      Write-Host ("No XML file found in '{0}'. Please provide the project XML path." -f $dir) -ForegroundColor Red
      exit 1
    }
  }
} else {
  $ProjectXmlPath = $resolvedInput
}

# Resolve search roots (SampleRoots.txt or -SearchRoot or default Music)
$SearchRoots = @()
if ($SearchRoot) {
  $SearchRoots = @($SearchRoot)
} else {
  $rootsFile = Join-Path $PSScriptRoot 'SampleRoots.txt'
  if (Test-Path -LiteralPath $rootsFile) {
    $lines = Get-Content -LiteralPath $rootsFile -ErrorAction SilentlyContinue | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') -and -not $_.StartsWith(';') }
    foreach ($line in $lines) { if (Test-Path -LiteralPath $line) { $SearchRoots += $line } }
  }
  if ($SearchRoots.Count -eq 0) {
    $userProfile = $env:USERPROFILE
    if ($userProfile) {
      $music = Join-Path $userProfile 'Music'
      if (Test-Path -LiteralPath $music) { $SearchRoots += $music } else { $SearchRoots += $userProfile }
    } else {
      $SearchRoots += (Get-Location).Path
    }
  }
}

# Safety notice
Write-Host "Bento Project Packer (beta): Run on copies of your projects. Keep backups. Use -DryRun first to preview changes." -ForegroundColor Yellow

. "$PSScriptRoot\Bento-ProjectPacker.ps1"
$exts = @(".wav", ".aif", ".aiff", ".flac", ".mp3", ".ogg")

$options = @{ Simulate = $Simulate }
if ($ReportPath) { $options.ReportPath = $ReportPath }

Invoke-BentoProjectPacker -XmlPath $ProjectXmlPath -SearchRoots $SearchRoots -Exts $exts @options

