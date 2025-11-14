# This file defines functions only; use Run-BentoProjectPacker.ps1 to execute.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Warning $msg }
function Write-Err($msg)  { Write-Host "[ERROR] $msg" -ForegroundColor Red }

function New-ReportPath {
    param(
        [Parameter(Mandatory=$true)][string]$ProjectDir,
        [string]$ReportPath
    )
    if ($ReportPath) { return $ReportPath }
    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    return (Join-Path $ProjectDir ("{0} Bento Project Packer Report.txt" -f $stamp))
}

function Get-UniqueFileName {
    param(
        [Parameter(Mandatory=$true)][string]$Directory,
        [Parameter(Mandatory=$true)][string]$BaseName,  # without extension
        [Parameter(Mandatory=$true)][string]$Extension  # with dot
    )
    $candidate = Join-Path $Directory ("{0}{1}" -f $BaseName, $Extension)
    if (-not (Test-Path -LiteralPath $candidate)) { return [System.IO.Path]::GetFileName($candidate) }
    $i = 1
    while ($true) {
        $name = "{0} ({1}){2}" -f $BaseName, $i, $Extension
        $candidate = Join-Path $Directory $name
        if (-not (Test-Path -LiteralPath $candidate)) { return $name }
        $i++
    }
}

function Get-SafeFolderName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $trimmed = $Name.Trim()
    if (-not $trimmed) { return '' }
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $builder = New-Object System.Text.StringBuilder
    foreach ($ch in $trimmed.ToCharArray()) {
        if ($invalid -contains $ch) {
            $null = $builder.Append('_')
        } else {
            $null = $builder.Append($ch)
        }
    }
    $candidate = $builder.ToString().Trim()
    $candidate = $candidate.TrimEnd('.').Trim()
    if (-not $candidate) { return '' }
    return $candidate
}

function Build-NameIndex {
    param(
        [Parameter(Mandatory=$true)][string]$Root,
        [string[]]$Exts
    )
    $index = @{}
    Write-Info "Scanning '$Root' for audio files..."
    $files = Get-ChildItem -LiteralPath $Root -Recurse -File -ErrorAction SilentlyContinue
    foreach ($f in $files) {
        $ext = [System.IO.Path]::GetExtension($f.Name).ToLowerInvariant()
        if ($Exts -and ($Exts.Count -gt 0) -and ($Exts -notcontains $ext)) { continue }
        $key = $f.Name.ToLowerInvariant()
        if (-not $index.ContainsKey($key)) { $index[$key] = @() }
        $index[$key] += $f.FullName
    }
    Write-Info ("Indexed {0} files" -f ($files | Where-Object { -not $Exts -or ($Exts -contains ([System.IO.Path]::GetExtension($_.Name).ToLowerInvariant())) } | Measure-Object).Count)
    return $index
}

function Build-DirIndex {
    param(
        [Parameter(Mandatory=$true)][string]$Root
    )
    $dirIndex = @{}
    Write-Info "Scanning '$Root' for directories..."
    $dirs = Get-ChildItem -LiteralPath $Root -Recurse -Directory -ErrorAction SilentlyContinue
    foreach ($d in $dirs) {
        $leaf = $d.Name.ToLowerInvariant()
        if (-not $dirIndex.ContainsKey($leaf)) { $dirIndex[$leaf] = @() }
        $dirIndex[$leaf] += $d.FullName
    }
    Write-Info ("Indexed {0} directories" -f ($dirs | Measure-Object).Count)
    return $dirIndex
}

function Get-TailDirNames {
    param([string]$OriginalPath)
    if ([string]::IsNullOrWhiteSpace($OriginalPath)) { return @() }
    $p = $OriginalPath.Replace('/', '\\')
    $parts = $p -split "[\\/]+" | Where-Object { $_ -ne '' }
    if ($parts.Count -lt 2) { return @() }
    $dirs = $parts[0..($parts.Count-2)]
    # Return last 3 directory names (deepest first)
    $take = [Math]::Min(3, $dirs.Count)
    return @($dirs[($dirs.Count-$take)..($dirs.Count-1)] | Sort-Object { -($_.Length) } | ForEach-Object { $_ })
}

function Find-CandidatesByFolder {
    param(
        [Parameter(Mandatory=$true)][hashtable]$DirIndex,
        [Parameter(Mandatory=$true)][string[]]$TailDirNames,
        [Parameter(Mandatory=$true)][string]$LeafFileName
    )
    $results = @()
    foreach ($dn in $TailDirNames) {
        $key = $dn.ToLowerInvariant()
        if (-not $DirIndex.ContainsKey($key)) { continue }
        foreach ($dirPath in $DirIndex[$key]) {
            # Search inside the matched directory tree for the exact leaf file name
            $found = Get-ChildItem -LiteralPath $dirPath -Recurse -File -Filter $LeafFileName -ErrorAction SilentlyContinue
            foreach ($f in $found) { $results += $f.FullName }
        }
        if ($results.Count -gt 0) { break }
    }
    return $results
}

function Find-CandidatesWithEverything {
    param(
        [Parameter(Mandatory=$true)][string]$EsPath,
        [Parameter(Mandatory=$true)][string]$SearchRoot,
        [Parameter()][string[]]$TailDirNames,
        [Parameter(Mandatory=$true)][string]$LeafFileName
    )
    $all = @()
    $tail = @()
    if ($TailDirNames) { $tail = $TailDirNames }

    # Build strict-to-relaxed queries: all tail dirs -> drop one -> just filename
    for ($k = $tail.Count; $k -ge 0; $k--) {
        $terms = @([string]::Format('path:"{0}"', $SearchRoot))
        if ($k -gt 0) {
            $terms += ($tail[($tail.Count-$k)..($tail.Count-1)] | ForEach-Object { [string]::Format('path:"{0}"', $_) })
        }
        # Include filename literal
        $terms += @([string]::Format('"{0}"', $LeafFileName))
        $query = ($terms -join ' ')
        try {
            $raw = & $EsPath $query 2>$null
        } catch {
            $raw = $null
        }
        if ($raw) {
            foreach ($line in $raw) {
                $p = $line.Trim()
                if ([string]::IsNullOrWhiteSpace($p)) { continue }
                # Ensure it is under SearchRoot and leaf matches exactly
                # Ensure it is under the current search root
                if (-not $p.StartsWith($SearchRoot, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                if ([System.IO.Path]::GetFileName($p) -ieq $LeafFileName) { $all += $p }
            }
        }
        if ($all.Count -gt 0) { break }
    }
    return $all
}

function Choose-CandidatePath {
    param(
        [Parameter(Mandatory=$true)][string[]]$Candidates,
        [Parameter(Mandatory=$true)][string]$OriginalValue
    )
    if ($Candidates.Count -eq 1) { return $Candidates[0] }
    # Prefer the candidate with the longest common suffix to the original string
    $orig = $OriginalValue.Replace('/', '\\')
    $scores = foreach ($c in $Candidates) {
        $a = ($orig.ToLowerInvariant()).ToCharArray()
        $b = ($c.ToLowerInvariant()).ToCharArray()
        $i = $a.Length - 1
        $j = $b.Length - 1
        $score = 0
        while ($i -ge 0 -and $j -ge 0 -and $a[$i] -eq $b[$j]) { $score++; $i--; $j-- }
        [pscustomobject]@{ Path = $c; Score = $score }
    }
    ($scores | Sort-Object -Property @{Expression='Score';Descending=$true}, @{Expression={ $_.Path.Length };Ascending=$true} | Select-Object -First 1).Path
}

function Invoke-BentoProjectPacker {
    param(
        [string]$XmlPath,
        [string[]]$SearchRoots,
        [string[]]$Exts,
        [switch]$Simulate,
        [string]$ReportPath
    )

    # Allow XmlPath to be either a file or a directory containing the XML
    try {
        $argPath = (Resolve-Path -LiteralPath $XmlPath -ErrorAction Stop).Path
    } catch {
        throw ("XML/project path not found: {0}" -f $XmlPath)
    }
    if (Test-Path -LiteralPath $argPath -PathType Container) {
        $dir = $argPath
        $candidate = Join-Path $dir 'project.xml'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            $resolvedXml = $candidate
        } else {
            $xmls = Get-ChildItem -LiteralPath $dir -Filter *.xml -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
            if ($xmls.Count -eq 1) {
                $resolvedXml = $xmls[0]
            } elseif ($xmls.Count -gt 1) {
                throw ("Multiple XML files found in '{0}'. Please specify the XML file explicitly." -f $dir)
            } else {
                throw ("No XML file found in '{0}'. Please provide the project XML path." -f $dir)
            }
        }
    } else {
        $resolvedXml = $argPath
    }
    $projectDir  = Split-Path -Path $resolvedXml -Parent
    Write-Info ("Bento Project Packer - Loading XML: {0}" -f $resolvedXml)
    [xml]$doc = Get-Content -LiteralPath $resolvedXml -Raw

    # Select all params elements with a filename attribute
    $paramNodes = $doc.SelectNodes("//params[@filename]")
    if (-not $paramNodes -or $paramNodes.Count -eq 0) {
        Write-Warn "No <params filename=...> entries found. Nothing to do."
        return
    }

    $trackInfoMap = @{}
    $trackNodes = $doc.SelectNodes("//track")
    $trackOrder = 1
    foreach ($track in $trackNodes) {
        $trackParams = $track.SelectSingleNode('./params')
        $rawName = $null
        if ($trackParams) {
            $rawName = $trackParams.GetAttribute('cellname')
            if (-not $rawName) { $rawName = $trackParams.GetAttribute('name') }
        }
        if ([string]::IsNullOrWhiteSpace($rawName)) { $rawName = "Track $trackOrder" }
        $safeName = Get-SafeFolderName -Name $rawName
        if (-not $safeName) { $safeName = "Track_$trackOrder" }
        $folderName = ("{0:D2}-{1}" -f $trackOrder, $safeName)
        $trackDir = Join-Path $projectDir $folderName
        $trackId = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($track)
        $trackInfoMap[$trackId] = [pscustomobject]@{
            Id = $trackId
            Name = $rawName
            FolderName = $folderName
            Directory = $trackDir
            Order = $trackOrder
        }
        $trackOrder++
    }
    $normalizedProjectDir = [System.IO.Path]::GetFullPath($projectDir)
    if (-not $normalizedProjectDir.EndsWith('\')) { $normalizedProjectDir = $normalizedProjectDir + '\' }

    # Determine search strategy
    $esPath = $null
    $es = Get-Command es -ErrorAction SilentlyContinue
    if (-not $es) { $es = Get-Command es.exe -ErrorAction SilentlyContinue }
    # Ignore WindowsApps alias (stub) which often doesn't execute real Everything
    if ($es -and ($es.Source -like '*WindowsApps*')) { $es = $null }
    # Try common install paths if not in PATH
    if (-not $es) {
        $commonEs = @(
            'C:\\Program Files\\Everything\\es.exe',
            'C:\\Program Files (x86)\\Everything\\es.exe',
            'C:\\Program Files\\Voidtools\\Everything\\es.exe',
            'C:\\Program Files (x86)\\Voidtools\\Everything\\es.exe'
        ) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
        if ($commonEs) {
            $esPath = $commonEs
        }
    }
    if (-not $esPath -and $es) { $esPath = $es.Source }

    if ($esPath) {
        Write-Info ("Using Everything CLI: {0}" -f $esPath)
    } else {
        Write-Info "Tip: Install Everything (Voidtools) and add 'es.exe' to PATH for quicker results."
    }

    # Build search indexes once if not using Everything
    $nameIndex = $null
    $dirIndex = $null
    if (-not $esPath) {
        $nameIndex = @{}
        $dirIndex  = @{}
        foreach ($root in $SearchRoots) {
            if (-not (Test-Path -LiteralPath $root)) { continue }
            $ni = Build-NameIndex -Root $root -Exts $Exts
            foreach ($k in $ni.Keys) { if (-not $nameIndex.ContainsKey($k)) { $nameIndex[$k] = @() }; $nameIndex[$k] += $ni[$k] }
            $di = Build-DirIndex  -Root $root
            foreach ($k in $di.Keys) { if (-not $dirIndex.ContainsKey($k)) { $dirIndex[$k] = @() }; $dirIndex[$k] += $di[$k] }
        }
    }

    $sourceToDest = @{}   # maps chosen source full path + track -> final relative destination path
    $stats = [pscustomobject]@{ Updated = 0; Skipped = 0; Missing = 0; Ambiguous = 0; Copied = 0 }

    # Track planned destination names to simulate collisions in DryRun and avoid overwrites in Apply
    $plannedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    try {
        Get-ChildItem -LiteralPath $projectDir -File -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $fullPath = [System.IO.Path]::GetFullPath($_.FullName)
            $relative = $fullPath
            if ($fullPath.StartsWith($normalizedProjectDir, [System.StringComparison]::OrdinalIgnoreCase)) {
                $relative = $fullPath.Substring($normalizedProjectDir.Length)
            }
            $relative = $relative.Replace('/', '\')
            if ($relative) { [void]$plannedNames.Add($relative) }
        }
    } catch {}
    $localFilesToDelete = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    # Prepare reporting
    $projectDir  = Split-Path -Path $resolvedXml -Parent
    $finalReportPath = New-ReportPath -ProjectDir $projectDir -ReportPath $ReportPath
    $report = New-Object System.Collections.Generic.List[string]
    $report.Add("Bento Project Packer Run Report") | Out-Null
    $report.Add(("Date: {0}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))) | Out-Null
    $report.Add(("Project: {0}" -f $resolvedXml)) | Out-Null
    $report.Add(("Project Folder: {0}" -f $projectDir)) | Out-Null
    $report.Add(("Search Roots: {0}" -f ($SearchRoots -join '; '))) | Out-Null
    if ($Exts -and $Exts.Count -gt 0) { $report.Add(("Extensions: {0}" -f ($Exts -join ', '))) | Out-Null }
    $report.Add("") | Out-Null
    $report.Add("Decisions") | Out-Null
    $report.Add("---------") | Out-Null

    foreach ($node in $paramNodes) {
        $orig = $node.GetAttribute('filename')
        if ([string]::IsNullOrWhiteSpace($orig)) { $stats.Skipped++; continue }

        $normalized = $orig.Trim()
        $normalized = $normalized.Replace('/', '\\')
        $leaf = Split-Path -Path $normalized -Leaf
        $trackNode = $node.SelectSingleNode('ancestor::track')
        $trackId = 0
        $trackInfo = $null
        if ($trackNode) {
            $trackId = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($trackNode)
            if ($trackInfoMap.ContainsKey($trackId)) { $trackInfo = $trackInfoMap[$trackId] }
        }

        $isAlreadyLocal = $normalized -match '^(?:\.\\|\./)'
        $localRelativePath = $null
        $localFullPath = $null
        if ($isAlreadyLocal) {
            $localRelativePath = $normalized.Substring(2)
            $localRelativePath = $localRelativePath.TrimStart('\', '/')
            if (-not $localRelativePath) { $localRelativePath = $leaf }
            $localFullPath = Join-Path $projectDir $localRelativePath
            if (-not (Test-Path -LiteralPath $localFullPath -PathType Leaf)) {
                $msg = "Missing local reference: $localFullPath"
                Write-Warn $msg
                if ($Simulate) { Write-Info $msg }
                $report.Add(("MISSING-LOCAL | attr='{0}' | path='{1}'" -f $orig, $localFullPath)) | Out-Null
                $localFullPath = $null
                $localRelativePath = $null
            }
        }

        # Prefer candidates using Everything if available; else use indexes (or reuse existing local files)
        $tailDirs = @(Get-TailDirNames -OriginalPath $normalized)
        $candidates = @()
        $existingRelativeNormalized = $null
        if ($localFullPath) {
            $candidates = @($localFullPath)
            if ($localRelativePath) { $existingRelativeNormalized = $localRelativePath.Replace('/', '\') }
        } elseif ($esPath) {
            foreach ($root in $SearchRoots) {
                if (-not (Test-Path -LiteralPath $root)) { continue }
                $c = Find-CandidatesWithEverything -EsPath $esPath -SearchRoot $root -TailDirNames $tailDirs -LeafFileName $leaf
                if ($c) { $candidates += $c }
            }
            if (-not $candidates -or $candidates.Count -eq 0) {
                # Fallback to filesystem indexes if Everything found nothing (or es.exe stub)
                if (-not $nameIndex -or -not $dirIndex) {
                    $nameIndex = @{}
                    $dirIndex  = @{}
                    foreach ($root in $SearchRoots) {
                        if (-not (Test-Path -LiteralPath $root)) { continue }
                        $ni = Build-NameIndex -Root $root -Exts $Exts
                        foreach ($k in $ni.Keys) { if (-not $nameIndex.ContainsKey($k)) { $nameIndex[$k] = @() }; $nameIndex[$k] += $ni[$k] }
                        $di = Build-DirIndex  -Root $root
                        foreach ($k in $di.Keys) { if (-not $dirIndex.ContainsKey($k)) { $dirIndex[$k] = @() }; $dirIndex[$k] += $di[$k] }
                    }
                }
                if ($tailDirs) {
                    $candidates = Find-CandidatesByFolder -DirIndex $dirIndex -TailDirNames $tailDirs -LeafFileName $leaf
                }
                if (-not $candidates -or $candidates.Count -eq 0) {
                    $key = $leaf.ToLowerInvariant()
                    if (-not $nameIndex.ContainsKey($key)) {
                        $msg = "Missing: '$leaf' not found under '$SearchRoot' (Everything+scan)"
                        Write-Warn $msg
                        if ($Simulate) { Write-Info $msg }
                        $report.Add(("MISSING | orig='{0}' | leaf='{1}' | via=Everything+scan" -f $orig, $leaf)) | Out-Null
                        $stats.Missing++
                        continue
                    }
                    $candidates = $nameIndex[$key]
                }
            }
        } else {
            if ($tailDirs) {
                $candidates = Find-CandidatesByFolder -DirIndex $dirIndex -TailDirNames $tailDirs -LeafFileName $leaf
            }
            # Fallback to global name index if folder-guided search fails
            if (-not $candidates -or $candidates.Count -eq 0) {
                $key = $leaf.ToLowerInvariant()
                if (-not $nameIndex.ContainsKey($key)) {
                    $msg = "Missing: '$leaf' not found under '$SearchRoot'"
                    Write-Warn $msg
                    if ($Simulate) { Write-Info $msg }
                    $report.Add(("MISSING | orig='{0}' | leaf='{1}'" -f $orig, $leaf)) | Out-Null
                    $stats.Missing++
                    continue
                }
                $candidates = $nameIndex[$key]
            }
        }
        $chosen = Choose-CandidatePath -Candidates $candidates -OriginalValue $normalized
        if ($candidates.Count -gt 1) {
            $stats.Ambiguous++
            $ambMsg = ("Ambiguous: {0} candidates for '{1}'. Chosen: {2}" -f $candidates.Count, $leaf, $chosen)
            Write-Warn $ambMsg
            if ($Simulate) { Write-Info $ambMsg }
            $report.Add(("AMBIGUOUS({0}) | leaf='{1}' | chosen='{2}'" -f $candidates.Count, $leaf, $chosen)) | Out-Null
        }

        $sourceKey = ("{0}|{1}" -f $trackId, $chosen.ToLowerInvariant())
        if ($sourceToDest.ContainsKey($sourceKey)) {
            $finalRelativePath = $sourceToDest[$sourceKey]
        } else {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
            $ext  = [System.IO.Path]::GetExtension($leaf)
            $candidate = "{0}{1}" -f $base, $ext
            $targetFolderRelative = $null
            if ($trackInfo) { $targetFolderRelative = $trackInfo.FolderName }
            $relativeCandidate = if ($targetFolderRelative) { Join-Path $targetFolderRelative $candidate } else { $candidate }
            $relativeCandidate = $relativeCandidate.Replace('/', '\')
            $targetFolderFull = $projectDir
            if ($trackInfo) { $targetFolderFull = $trackInfo.Directory }
            $destPath = $null
            $needsCopy = $true
            if ($existingRelativeNormalized -and ($existingRelativeNormalized -ieq $relativeCandidate)) {
                $needsCopy = $false
                $finalRelativePath = $existingRelativeNormalized
                $destPath = Join-Path $projectDir $finalRelativePath
            } else {
                $n = 1
                while ($plannedNames.Contains($relativeCandidate)) {
                    $candidate = "{0} ({1}){2}" -f $base, $n, $ext
                    if ($targetFolderRelative) {
                        $relativeCandidate = Join-Path $targetFolderRelative $candidate
                    } else {
                        $relativeCandidate = $candidate
                    }
                    $relativeCandidate = $relativeCandidate.Replace('/', '\')
                    $n++
                }
                $destPath = Join-Path $targetFolderFull $candidate
                while (Test-Path -LiteralPath $destPath) {
                    $candidate = "{0} ({1}){2}" -f $base, $n, $ext
                    if ($targetFolderRelative) {
                        $relativeCandidate = Join-Path $targetFolderRelative $candidate
                    } else {
                        $relativeCandidate = $candidate
                    }
                    $relativeCandidate = $relativeCandidate.Replace('/', '\')
                    $destPath = Join-Path $targetFolderFull $candidate
                    $n++
                }
                $finalRelativePath = $relativeCandidate
            }

            if ($finalRelativePath -and -not $plannedNames.Contains($finalRelativePath)) {
                [void]$plannedNames.Add($finalRelativePath)
            }

            if ($needsCopy) {
                if (-not (Test-Path -LiteralPath $targetFolderFull)) {
                    if ($Simulate) {
                        Write-Info ("WhatIf: Create folder '{0}'" -f $targetFolderFull)
                    } else {
                        New-Item -ItemType Directory -Path $targetFolderFull -Force | Out-Null
                    }
                }

                if (-not (Test-Path -LiteralPath $destPath)) {
                    $leafName = Split-Path -Path $finalRelativePath -Leaf
                    if ($Simulate) {
                        Write-Info ("WhatIf: Copy '{0}' -> '{1}' as '{2}'" -f $chosen, $targetFolderFull, $leafName)
                    } else {
                        Write-Info ("Copy '{0}' -> '{1}' as '{2}'" -f $chosen, $targetFolderFull, $leafName)
                        Copy-Item -LiteralPath $chosen -Destination $destPath
                    }
                    $report.Add(("COPY | src='{0}' | dest='{1}'" -f $chosen, $destPath)) | Out-Null
                    $stats.Copied++
                } else {
                    $report.Add(("COPY-SKIP | already exists | dest='{0}'" -f $destPath)) | Out-Null
                }

                if ($localFullPath -and $existingRelativeNormalized -and ($finalRelativePath -ne $existingRelativeNormalized)) {
                    $added = $localFilesToDelete.Add($localFullPath)
                    if ($added) {
                        if ($Simulate) {
                            Write-Info ("WhatIf: Would delete old local '{0}' after packaging" -f $localFullPath)
                        }
                        $report.Add(("REPACK-LOCAL | from='{0}' | to='{1}'" -f $localFullPath, $destPath)) | Out-Null
                    }
                }
            } else {
                $report.Add(("REUSE | dest='{0}'" -f (Join-Path $projectDir $finalRelativePath))) | Out-Null
            }
            $sourceToDest[$sourceKey] = $finalRelativePath
        }

        $newValue = ".\" + ($finalRelativePath -replace '/', '\')
        if ($orig -ne $newValue) {
            $node.SetAttribute('filename', $newValue)
            $stats.Updated++
            if ($Simulate) { Write-Info ("Would set attr filename='{0}'" -f $newValue) }
            $report.Add(("SETATTR | filename='{0}'" -f $newValue)) | Out-Null
        } else {
            $stats.Skipped++
            if ($Simulate) { Write-Info ("Attr already correct: '{0}'" -f $newValue) }
            $report.Add(("SETATTR-SKIP | already .\\ | filename='{0}'" -f $newValue)) | Out-Null
        }
    }

    if ($localFilesToDelete.Count -gt 0) {
        foreach ($oldPath in $localFilesToDelete) {
            if ($Simulate) {
                Write-Info ("WhatIf: Would delete '{0}' after repack" -f $oldPath)
                $report.Add(("DELETE-LOCAL-WHATIF | path='{0}'" -f $oldPath)) | Out-Null
            } else {
                try {
                    if (Test-Path -LiteralPath $oldPath) {
                        Remove-Item -LiteralPath $oldPath -Force
                        $report.Add(("DELETE-LOCAL | path='{0}'" -f $oldPath)) | Out-Null
                    }
                } catch {
                    $report.Add(("DELETE-LOCAL-FAIL | path='{0}' | err='{1}'" -f $oldPath, $_)) | Out-Null
                    Write-Warn ("Failed to delete '{0}': {1}" -f $oldPath, $_)
                }
            }
        }
    }

    # Summary + report write
    $report.Add("") | Out-Null
    $report.Add("Summary") | Out-Null
    $report.Add("-------") | Out-Null
    $report.Add(("Updated: {0}" -f $stats.Updated)) | Out-Null
    $report.Add(("Copied: {0}" -f $stats.Copied)) | Out-Null
    $report.Add(("Missing: {0}" -f $stats.Missing)) | Out-Null
    $report.Add(("Ambiguous: {0}" -f $stats.Ambiguous)) | Out-Null
    $report.Add(("Skipped: {0}" -f $stats.Skipped)) | Out-Null
    try {
        $reportDir = Split-Path -Path $finalReportPath -Parent
        if ($reportDir -and -not (Test-Path -LiteralPath $reportDir)) { New-Item -ItemType Directory -Path $reportDir -Force | Out-Null }
        Set-Content -LiteralPath $finalReportPath -Value $report -Encoding UTF8 -Force
        Write-Info ("Report: {0}" -f $finalReportPath)
        # Also save a copy in the current working directory
        $reportFileName = Split-Path -Path $finalReportPath -Leaf
        $cwd = (Get-Location).Path
        $cwdReportPath = Join-Path $cwd $reportFileName
        if ($cwdReportPath -ne $finalReportPath) {
            $cwdDir = Split-Path -Path $cwdReportPath -Parent
            if ($cwdDir -and -not (Test-Path -LiteralPath $cwdDir)) { New-Item -ItemType Directory -Path $cwdDir -Force | Out-Null }
            Set-Content -LiteralPath $cwdReportPath -Value $report -Encoding UTF8 -Force
            Write-Info ("Report (copy): {0}" -f $cwdReportPath)
        }
    } catch {
        Write-Warn ("Failed to write report: {0}" -f $_)
    }

    if ($Simulate) {
        Write-Info ("WhatIf: Would update {0} entries. Copies: {1}, Missing: {2}, Ambiguous: {3}" -f $stats.Updated, $stats.Copied, $stats.Missing, $stats.Ambiguous)
        return
    }

    # Always create a simple pre-change backup beside the XML
    try {
        $simpleBackup = "$resolvedXml.bak"
        Copy-Item -LiteralPath $resolvedXml -Destination $simpleBackup -Force
        Write-Info ("Backup (pre-change): {0}" -f $simpleBackup)
    } catch {
        Write-Warn ("Failed to create pre-change backup: {0}" -f $_)
    }

    Write-Info "Saving XML..."
    $doc.Save($resolvedXml)
    Write-Info ("Done. Updated: {0}, Copied: {1}, Missing: {2}, Ambiguous: {3}, Skipped: {4}" -f $stats.Updated, $stats.Copied, $stats.Missing, $stats.Ambiguous, $stats.Skipped)
}

# End of function library
