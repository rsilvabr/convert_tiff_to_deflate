# ── SETTINGS ──────────────────────────────────────────────────────
$Workers    = 8
$DryRun     = $false
$Recurse    = $true
$SafeMode   = $true      # true  → skip multi-page TIFFs (scanner IR, Photoshop layers)
                          # false → compress all TIFFs including multi-page ones
$SkipLzwAsCompressed = $false  # true = treat LZW as already compressed (skip ZIP re-compression)
                          # false = re-compress LZW to ZIP (default)
$OutputDir  = ""          # ""          = overwrite original in place
                          # "tiff_zip"  = create subfolder per group
                          # "F:\ZIP"    = absolute output path
$StagingDir = ""          # ""          = disabled
                          # "E:\staging" = write here, move to final destination after each group
$Overwrite  = $false
$MagickTimeout = 30       # seconds timeout for magick identify (prevents hang on corrupted files)
# ──────────────────────────────────────────────────────────────────

# ── Cleanup on interrupt ─────────────────────────────────────────
$script:cleanupDirs = @()
if ($StagingDir) { $script:cleanupDirs += $StagingDir }

trap {
    Write-Log "Interrupted! Cleaning up staging files..." "WARN"
    foreach ($dir in $script:cleanupDirs) {
        if (Test-Path -LiteralPath $dir) {
            Remove-Item -Path "$dir\*" -Force -ErrorAction SilentlyContinue
        }
    }
    break
}

# ── Logging ───────────────────────────────────────────────────────
$scriptName = "compress_tiff_zip"
$logDir     = Join-Path $PWD.Path "Logs\$scriptName"
[System.IO.Directory]::CreateDirectory($logDir) | Out-Null
$logFile    = Join-Path $logDir "$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$msg, [string]$level = "INFO")
    $line = "$(Get-Date -Format 'HH:mm:ss') | $level | $msg"
    Write-Host $line
    [System.IO.File]::AppendAllText($logFile, $line + [System.Environment]::NewLine)
}

$script:counterTotal   = 0
$script:okTotal        = 0
$script:skipTotal      = 0
$script:multiTotal     = 0
$script:errTotal       = 0
$script:total          = 0
$script:multiPagePaths = [System.Collections.Concurrent.ConcurrentBag[string]]::new()

function Process-Results {
    param($lines)
    foreach ($line in $lines) {
        $script:counterTotal++
        $lvl = "INFO"
        if     ($line -match '^OK')    { $script:okTotal++ }
        elseif ($line -match '^SKIP')  { $script:skipTotal++ }
        elseif ($line -match '^MULTI') { $script:multiTotal++; $lvl = "WARN" }
        elseif ($line -match '^ERROR') { $script:errTotal++; $lvl = "ERROR" }
        elseif ($line -match '^WARN')  { $lvl = "WARN" }
        Write-Log "[$($script:counterTotal)/$($script:total)] $line" $lvl
    }
}

# ── Collect files ─────────────────────────────────────────────────
$root = $PWD.Path

$files = Get-ChildItem -LiteralPath $root -File -Recurse:$Recurse |
         Where-Object { $_.Extension -match '^\.(tif|tiff)$' }

$script:total = $files.Count

Write-Log "Log: $logFile"

if ($script:total -eq 0) {
    Write-Log "No TIFF files found in: $root" "WARN"
    Write-Log "Make sure you are in the right folder and that .tif or .tiff files exist"
} else {
    $modeLabel = if ($SafeMode) { "SAFE (multi-page TIFFs will be skipped)" } else { "STANDARD (all TIFFs will be compressed)" }
    Write-Log "TIFFs: $($script:total) | Workers: $Workers | OutputDir: $(if ($OutputDir) { $OutputDir } else { '(overwrite in place)' }) | Staging: $(if ($StagingDir) { $StagingDir } else { 'disabled' }) | DryRun: $DryRun"
    Write-Log "Mode: $modeLabel"

    $groups = $files | Group-Object { $_.DirectoryName }

    foreach ($group in $groups) {
        $groupDir   = $group.Name
        $groupFiles = $group.Group

        if ($groups.Count -gt 1) {
            Write-Log ""
            Write-Log "── Group: $groupDir ($($groupFiles.Count) file(s))"
        }

        $finalDir = if ($OutputDir) {
            if ([System.IO.Path]::IsPathRooted($OutputDir)) { $OutputDir }
            else { Join-Path $groupDir $OutputDir }
        } else { $groupDir }

        $writeDir = if ($StagingDir -and -not $DryRun) { $StagingDir } else { $finalDir }

        if ($StagingDir -and -not $DryRun) { [System.IO.Directory]::CreateDirectory($StagingDir) | Out-Null }
        if ($OutputDir)                    { [System.IO.Directory]::CreateDirectory($finalDir)   | Out-Null }

        $safeL        = $SafeMode
        $multiPageBag = $script:multiPagePaths
        $skipLzwL     = $SkipLzwAsCompressed

        # Build staging map to track UUID -> original filename
        $script:stagingMap = @{}

        $job = $groupFiles | ForEach-Object -Parallel {
            $src       = $_.FullName
            $name      = $_.Name
            $writeDirL = $using:writeDir
            $finalDirL = $using:finalDir
            $dryL      = $using:DryRun
            $overL     = $using:Overwrite
            $safeMode  = $using:safeL
            $bagL      = $using:multiPageBag
            $skipLzw   = $using:skipLzwL

            $stagingName = "$([guid]::NewGuid().ToString('N'))_$name"
            $writeDst = Join-Path $writeDirL $stagingName
            $finalDst = Join-Path $finalDirL $name

            # Check current compression (uses -@ to handle brackets in path names)
            $argComp = [System.IO.Path]::GetTempFileName()
            [System.IO.File]::WriteAllText($argComp, "-s`n-s`n-s`n-Compression`n$src`n")
            $comp = exiftool -@ $argComp 2>$null
            $exifExit = $LASTEXITCODE
            Remove-Item $argComp -Force
            if ($exifExit -ne 0 -or -not $comp) {
                return @{ Result = "ERROR (exiftool check) | $name | cannot detect compression"; StagingName = $null; OriginalName = $name }
            }
            if ($comp -match $(if ($skipLzw) { 'Deflate|ZIP|Adobe|LZW' } else { 'Deflate|ZIP|Adobe' })) { 
                return @{ Result = "SKIP ($comp) | $name"; StagingName = $null; OriginalName = $name }
            }

            # Check if output already exists
            if ((Test-Path -LiteralPath $finalDst) -and -not $overL -and ($finalDst -ne $src)) {
                return @{ Result = "SKIP (exists) | $name"; StagingName = $null; OriginalName = $name }
            }

            # Safe mode: detect and skip multi-page TIFFs before touching them.
            if ($safeMode) {
                $magickTimeoutSec = 30  # Fixed timeout for magick identify
                $pageCountJob = Start-Job { magick identify $using:src 2>$null }
                $pageCountJob | Wait-Job -Timeout $magickTimeoutSec | Out-Null
                if ($pageCountJob.State -eq 'Running') {
                    Stop-Job $pageCountJob
                    Remove-Job $pageCountJob
                    return @{ Result = "ERROR (magick timeout) | $name | possibly corrupted"; StagingName = $null; OriginalName = $name }
                }
                $pageCount = ($pageCountJob | Receive-Job | Measure-Object -Line).Lines
                Remove-Job $pageCountJob
                if ($pageCount -gt 1) {
                    $bagL.Add($src) | Out-Null
                    return @{ Result = "MULTI ($pageCount IFDs — skipped) | $name"; StagingName = $null; OriginalName = $name }
                }
            }

            if ($dryL) { return @{ Result = "DRY ($comp → ZIP) | $name"; StagingName = $null; OriginalName = $name } }

            # Compress with ImageMagick
            $out = magick -quiet $src -compress zip $writeDst 2>&1
            if ($LASTEXITCODE -ne 0) { return @{ Result = "ERROR (magick) | $name | $out"; StagingName = $null; OriginalName = $name } }

            # Verify EXIF was preserved (magick usually keeps it, but check)
            $argExif = [System.IO.Path]::GetTempFileName()
            [System.IO.File]::WriteAllText($argExif, "-s`n-s`n-s`n-EXIF:Make`n$writeDst`n")
            $hasExif = exiftool -@ $argExif 2>$null
            Remove-Item $argExif -Force

            if (-not $hasExif) {
                # Fallback: copy EXIF from original to the compressed file
                $argCopy = [System.IO.Path]::GetTempFileName()
                [System.IO.File]::WriteAllText($argCopy, "-q`n-q`n-overwrite_original`n-tagsfromfile`n$src`n-all:all`n-unsafe`n$writeDst`n")
                exiftool -@ $argCopy | Out-Null
                Remove-Item $argCopy -Force
                if ($LASTEXITCODE -ne 0) { return @{ Result = "WARN (exiftool failed, ZIP ok) | $name"; StagingName = $stagingName; OriginalName = $name } }
            }

            return @{ Result = "OK ($comp → ZIP) | $name"; StagingName = $stagingName; OriginalName = $name }

        } -ThrottleLimit $Workers -AsJob

        while ($job.State -eq 'Running') {
            $results = Receive-Job $job
            foreach ($r in $results) {
                if ($r.StagingName) { $script:stagingMap[$r.OriginalName] = $r.StagingName }
                Process-Results @($r.Result)
            }
            Start-Sleep -Milliseconds 300
        }
        $finalResults = Receive-Job $job
        foreach ($r in $finalResults) {
            if ($r.StagingName) { $script:stagingMap[$r.OriginalName] = $r.StagingName }
            Process-Results @($r.Result)
        }
        Remove-Job $job

        # Move from staging to final destination (with integrity check)
        if ($StagingDir -and -not $DryRun) {
            $moved = 0
            foreach ($f in $groupFiles) {
                # Use the UUID-mapped staging name, not the original name
                $originalName = $f.Name
                if (-not $script:stagingMap.ContainsKey($originalName)) { continue }
                
                $stagingName = $script:stagingMap[$originalName]
                $stagePath = Join-Path $StagingDir $stagingName
                $destPath  = Join-Path $finalDir   $originalName
                
                if ((Test-Path -LiteralPath $stagePath) -and $stagePath -ne $destPath) {
                    $stageSize = (Get-Item -LiteralPath $stagePath).Length
                    Move-Item -Force -LiteralPath $stagePath -Destination $destPath
                    # Verify move succeeded
                    if ((Test-Path -LiteralPath $destPath) -and ((Get-Item -LiteralPath $destPath).Length -eq $stageSize)) {
                        $moved++
                    } else {
                        Write-Log "ERROR (move failed) | $originalName" "ERROR"
                    }
                }
            }
            if ($moved -gt 0) { Write-Log "  → Moved $moved file(s) → $finalDir" }
        }
    }

    Write-Log ""
    Write-Log ("─" * 50)
    if ($SafeMode) {
        Write-Log "Done: $($script:okTotal) OK | $($script:skipTotal) skipped | $($script:multiTotal) multi-page (not touched) | $($script:errTotal) errors | $($script:counterTotal)/$($script:total) processed"
    } else {
        Write-Log "Done: $($script:okTotal) OK | $($script:skipTotal) skipped | $($script:errTotal) errors | $($script:counterTotal)/$($script:total) processed"
    }

    if ($script:multiTotal -gt 0) {
        Write-Log ""
        Write-Log "── Multi-page TIFFs found (not compressed — review manually):"
        foreach ($p in ($script:multiPagePaths | Sort-Object)) {
            Write-Log "   $p" "WARN"
        }
    }
}

Write-Log "Log: $logFile"
