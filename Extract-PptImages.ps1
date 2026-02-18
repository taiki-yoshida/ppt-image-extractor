<#
.SYNOPSIS
    Extracts all images from a PowerPoint (.pptx) file, named by slide and sequence.

.DESCRIPTION
    Extracts images from a .pptx file and saves them with the naming convention
    slideNN_MM.ext where NN is the slide number and MM is the image sequence
    within that slide (based on document order).

.PARAMETER Path
    Path to the .pptx file.

.PARAMETER OutputDir
    Optional output directory. Defaults to a folder named after the .pptx file
    in the current directory.

.EXAMPLE
    .\Extract-PptImages.ps1 -Path "presentation.pptx"
    .\Extract-PptImages.ps1 -Path "C:\docs\deck.pptx" -OutputDir "C:\output\images"
#>

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [Parameter(Position = 1)]
    [string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Resolve full path
$Path = Resolve-Path $Path -ErrorAction Stop | Select-Object -ExpandProperty Path

if (-not (Test-Path $Path)) {
    Write-Error "File not found: $Path"
    return
}

if ([System.IO.Path]::GetExtension($Path) -notin '.pptx', '.pptm') {
    Write-Error "File must be a .pptx or .pptm file."
    return
}

# Default output directory: same name as pptx without extension
if (-not $OutputDir) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $OutputDir = Join-Path (Get-Location) $baseName
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# Extract to temp directory
Add-Type -AssemblyName System.IO.Compression.FileSystem
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("pptx_extract_" + [System.Guid]::NewGuid().ToString("N"))
try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($Path, $tempDir)

    $slidesDir = Join-Path $tempDir "ppt\slides"
    $relsDir = Join-Path $tempDir "ppt\slides\_rels"
    $mediaDir = Join-Path $tempDir "ppt\media"

    if (-not (Test-Path $slidesDir)) {
        Write-Error "Invalid pptx structure: ppt/slides not found."
        return
    }

    # Get slide files sorted by slide number
    $slideFiles = Get-ChildItem -Path $slidesDir -Filter "slide*.xml" |
        Where-Object { $_.Name -match '^slide(\d+)\.xml$' } |
        Sort-Object { [int]($_.Name -replace '[^\d]', '') }

    $totalImages = 0
    $padWidth = [Math]::Max(2, $slideFiles.Count.ToString().Length)

    foreach ($slideFile in $slideFiles) {
        $slideNum = [int]($slideFile.Name -replace '[^\d]', '')
        $slideNumPadded = $slideNum.ToString().PadLeft($padWidth, '0')

        # Parse the relationship file for this slide
        $relsFile = Join-Path $relsDir "$($slideFile.Name).rels"
        if (-not (Test-Path $relsFile)) {
            continue
        }

        [xml]$relsXml = Get-Content -Path $relsFile -Raw
        # Build a lookup: rId -> media file path
        $imageRels = @{}
        foreach ($rel in $relsXml.Relationships.Relationship) {
            if ($rel.Type -like "*/image") {
                # Target is relative, e.g. "../media/image1.png"
                $mediaPath = Join-Path (Join-Path $tempDir "ppt\slides") $rel.Target
                $mediaPath = [System.IO.Path]::GetFullPath($mediaPath)
                $imageRels[$rel.Id] = $mediaPath
            }
        }

        if ($imageRels.Count -eq 0) {
            continue
        }

        # Parse the slide XML to find image references in document order
        [xml]$slideXml = Get-Content -Path $slideFile.FullName -Raw
        $nsManager = New-Object System.Xml.XmlNamespaceManager($slideXml.NameTable)
        $nsManager.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
        $nsManager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

        # Find all blip elements which reference images, in document order
        $blipNodes = $slideXml.SelectNodes("//a:blip[@r:embed]", $nsManager)

        $seenEmbeds = @{}
        $imageIndex = 0

        foreach ($blip in $blipNodes) {
            $embedId = $blip.GetAttribute("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

            # Skip duplicate references to the same image on the same slide
            if ($seenEmbeds.ContainsKey($embedId)) {
                continue
            }
            $seenEmbeds[$embedId] = $true

            if ($imageRels.ContainsKey($embedId)) {
                $imageIndex++
                $sourcePath = $imageRels[$embedId]

                if (Test-Path $sourcePath) {
                    $ext = [System.IO.Path]::GetExtension($sourcePath)
                    $newName = "slide${slideNumPadded}_$($imageIndex.ToString().PadLeft(2, '0'))${ext}"
                    $destPath = Join-Path $OutputDir $newName

                    Copy-Item -Path $sourcePath -Destination $destPath -Force
                    $totalImages++
                }
            }
        }

        if ($imageIndex -gt 0) {
            Write-Host "Slide ${slideNum}: ${imageIndex} image(s)"
        }
    }

    Write-Host ""
    Write-Host "Done. Extracted $totalImages image(s) to: $OutputDir"
}
finally {
    # Clean up temp directory
    if (Test-Path $tempDir) {
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
