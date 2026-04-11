# compress_tiff_zip.ps1

Batch TIFF → TIFF ZIP/Deflate compressor. Re-compresses uncompressed TIFFs in place
(or to a separate folder), preserving all EXIF metadata and ICC profiles.

Includes a **safe mode** (`$SafeMode = $true`, default) that automatically skips
multi-page TIFFs to protect scanner IR channels, Photoshop layers, and other
multi-IFD files that external tools cannot safely recompress.

---

## Requirements

```
PowerShell 7+          https://github.com/PowerShell/PowerShell/releases
ImageMagick 7 (magick) https://imagemagick.org
exiftool               https://exiftool.org
```

Both `magick.exe` and `exiftool.exe` must be on your PATH.

---

## Disclaimer

These tools were made for my personal workflow (with the help of Claude). Use at your own risk — I am not responsible for any issues you may encounter.
If you choose to use it and find any errors/bugs, please let me know.

---

## Quick start

```powershell
# Navigate to the folder with your TIFFs, then run:
.\compress_tiff_zip.ps1
```

Alternatively, copy the entire script, open a new PowerShell 7 terminal in the folder
containing your TIFFs, paste and press Enter. No file installation needed.

```powershell
# To preview without changing anything, set $DryRun = $true at the top.
# Recursive mode is $true by default — all subfolders are processed.
# Set $Recurse = $false to process only the current folder.
```

---

## Key settings

Edit at the top of the script:

```powershell
$SafeMode   = $true
# true  → skip multi-page TIFFs (recommended — see Safe Mode section below)
# false → compress all TIFFs including multi-page ones

$SkipLzwAsCompressed = $false
# true = treat LZW as already compressed (skip ZIP re-compression)
# false (default) = convert LZW → ZIP

$Recurse    = $true
# true  → process TIFFs in all subfolders recursively
# false → only process TIFFs in the current folder

$Workers    = 8
# Number of parallel threads.

$OutputDir  = ""
# ""           → overwrite original TIFF in place (default)
# "tiff_zip"   → create a subfolder named "tiff_zip" inside each source folder
# "F:\Archive" → absolute path — all compressed TIFFs go here

$StagingDir = ""
# ""           → disabled, write directly to OutputDir
# "E:\staging" → write here first, move to final destination after each group.
#                Useful to separate read I/O (source HDD) from write I/O.

$DryRun     = $false
# true  → show what would be compressed without changing any files

$Overwrite  = $false
# false → skip if output file already exists
# true  → always overwrite
```

---

## Safe mode

`$SafeMode = $true` (default) detects TIFFs with more than one IFD before compressing
and skips them — logged as `MULTI` but never touched.

**Why this matters:**

Some TIFFs contain multiple internal "pages" (IFDs) that external tools cannot safely
recompress:

**SilverFast scanner TIFFs (Epson V700, etc.)** store 3 IFDs: the main RGB image,
a thumbnail, and an infrared channel used for automatic dust and scratch removal.
The IR channel is referenced by a proprietary byte-offset pointer (tag `0x89ab`).
After recompression, the data moves to a different byte position — the pointer becomes
invalid and SilverFast can no longer find the IR channel. Dust removal stops working.

**Photoshop layered TIFFs** — untested, but may use proprietary Adobe structures across
multiple IFDs that external tools do not preserve correctly.

Multi-page files are listed at the end of the log with their full paths, so you can
review and decide what to do with each one manually.

Use `$SafeMode = $false` only if you are certain your collection contains no multi-page
TIFFs — for example, pure camera exports from Capture One or NX Studio.

---

## Using with Capture One

In my experience, the script works fine on TIFFs that are actively being used as
input in Capture One — colors, edits, and adjustments are all preserved, and I never
needed to close Capture One before running it. Other software may behave differently,
so test with a few files before processing a large library.

---

## Output modes

| `$OutputDir` | Behavior |
|--------------|----------|
| `""` (default) | Overwrites the original TIFF in place |
| `"tiff_zip"` | Creates a `tiff_zip/` subfolder inside each source folder |
| `"F:\Archive"` | All files go to this absolute path |

---

## Log

```
<script_folder>/Logs/compress_tiff_zip/YYYYMMDD_HHMMSS.log
```

Each file produces one line:
```
OK (Uncompressed → ZIP) | photo.tif
SKIP (Deflate)           | photo.tif    ← already compressed
SKIP (exists)            | photo.tif    ← output already exists
MULTI (3 IFDs — skipped) | scan.tif     ← safe mode, not touched
ERROR (magick) | photo.tif              ← compression failed
```

Final summary (safe mode on):
```
Done: 45 OK | 3 skipped | 2 multi-page (not touched) | 0 errors | 50/50 processed

── Multi-page TIFFs found (not compressed — review manually):
   F:\scans\scan_RGBI.tif
   F:\archive\photo_layers.tif
```

---

## Notes

- Already-compressed TIFFs (Deflate/ZIP/Adobe) are automatically skipped (LZW if `$SkipLzwAsCompressed = $true`)
- EXIF is verified after compression — if ImageMagick dropped it, exiftool restores it from the original
- All exiftool calls use `-@` argument files to handle folder names with brackets like `[FINAL]`
- Tested with PowerShell 7.6, Windows 11
