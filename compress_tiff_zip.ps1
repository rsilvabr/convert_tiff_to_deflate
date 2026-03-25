# ── SETTINGS ──────────────────────────────────────────────────────
$Workers    = 8
$DryRun     = $false
$Recurse    = $true
$SafeMode   = $true      # true  → skip multi-page TIFFs (scanner IR, Photoshop layers)
                          # false → compress all TIFFs including multi-page ones
$OutputDir  = ""          # ""          = overwrite original in place
                          # "tiff_zip"  = create subfolder per group
                          # "F:\ZIP"    = absolute output path
$StagingDir = ""          # ""          = disabled
                          # "E:\staging" = write here, move to final destination after each group
$Overwrite  = $false
# ──────────────────────────────────────────────────────────────────

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

        $job = $groupFiles | ForEach-Object -Parallel {
            $src       = $_.FullName
            $name      = $_.Name
            $writeDirL = $using:writeDir
            $finalDirL = $using:finalDir
            $dryL      = $using:DryRun
            $overL     = $using:Overwrite
            $safeMode  = $using:safeL
            $bagL      = $using:multiPageBag

            $writeDst = Join-Path $writeDirL $name
            $finalDst = Join-Path $finalDirL $name

            # Check current compression (uses -@ to handle brackets in path names)
            $argComp = [System.IO.Path]::GetTempFileName()
            [System.IO.File]::WriteAllText($argComp, "-s`n-s`n-s`n-Compression`n$src`n")
            $comp = exiftool -@ $argComp 2>$null
            Remove-Item $argComp -Force
            if ($comp -match 'Deflate|ZIP|Adobe') { return "SKIP ($comp) | $name" }

            # Check if output already exists
            if ((Test-Path -LiteralPath $finalDst) -and -not $overL -and ($finalDst -ne $src)) {
                return "SKIP (exists) | $name"
            }

            # Safe mode: detect and skip multi-page TIFFs before touching them.
            # Multi-page TIFFs include: scanner RGB+IR files (SilverFast), Photoshop layered files.
            # Compressing them with external tools breaks internal byte-offset pointers
            # and proprietary tags — e.g. SilverFast loses its IR dust-removal channel.
            if ($safeMode) {
                $pageCount = (magick identify $src 2>$null | Measure-Object -Line).Lines
                if ($pageCount -gt 1) {
                    $bagL.Add($src) | Out-Null
                    return "MULTI ($pageCount IFDs — skipped) | $name"
                }
            }

            if ($dryL) { return "DRY ($comp → ZIP) | $name" }

            # Compress with ImageMagick
            $out = magick -quiet $src -compress zip $writeDst 2>&1
            if ($LASTEXITCODE -ne 0) { return "ERROR (magick) | $name | $out" }

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
                if ($LASTEXITCODE -ne 0) { return "WARN (exiftool failed, ZIP ok) | $name" }
            }

            return "OK ($comp → ZIP) | $name"

        } -ThrottleLimit $Workers -AsJob

        while ($job.State -eq 'Running') {
            Process-Results (Receive-Job $job)
            Start-Sleep -Milliseconds 300
        }
        Process-Results (Receive-Job $job)
        Remove-Job $job

        # Move from staging to final destination
        if ($StagingDir -and -not $DryRun) {
            $moved = 0
            foreach ($f in $groupFiles) {
                $stagePath = Join-Path $StagingDir $f.Name
                $destPath  = Join-Path $finalDir   $f.Name
                if ((Test-Path -LiteralPath $stagePath) -and $stagePath -ne $destPath) {
                    Move-Item -Force -LiteralPath $stagePath -Destination $destPath
                    $moved++
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
