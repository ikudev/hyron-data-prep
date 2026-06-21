<#
.SYNOPSIS
    Recursively converts Word, Excel, and PowerPoint files in a folder to PDF using default print/layout settings.
.DESCRIPTION
    This script traverses a folder recursively, identifies Office documents (Word, Excel, PowerPoint),
    and converts them to PDF using native COM automation APIs. It runs silently in the background
    without user prompts.
.PARAMETER InputFolder
    The directory to search recursively for documents to convert.
.PARAMETER OutputFolder
    The directory where the output PDF files will be stored. Maintains the source directory structure.
    If not specified, defaults to the InputFolder.
.PARAMETER LogPath
    Optional path to a log file where conversion logs will be appended.
.EXAMPLE
    .\Convert-DocumentsToPdf.ps1 -InputFolder "C:\Docs" -OutputFolder "C:\PDFs" -LogPath "C:\Docs\conversion.log"
#>

param (
    [Parameter(Mandatory=$true, HelpMessage="Path to the directory containing documents to convert.")]
    [string]$InputFolder,

    [Parameter(Mandatory=$false, HelpMessage="Path to the output directory for PDF files.")]
    [string]$OutputFolder,

    [Parameter(Mandatory=$false, HelpMessage="Path to a text log file.")]
    [string]$LogPath
)

# Convert relative paths to absolute paths
$InputFolder = [System.IO.Path]::GetFullPath($InputFolder)
if ($OutputFolder) {
    $OutputFolder = [System.IO.Path]::GetFullPath($OutputFolder)
} else {
    $OutputFolder = $InputFolder
}

# Logger helper
function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMsg = "[$timestamp] [$Level] $Message"
    Write-Host $logMsg
    if ($LogPath) {
        try {
            $logMsg | Out-File -FilePath $LogPath -Append -Encoding utf8 -ErrorAction SilentlyContinue
        } catch {}
    }
}

# Helper to resolve relative subdirectory structure
function Get-RelativePath {
    param (
        [string]$BasePath,
        [string]$TargetPath
    )
    $baseClean = $BasePath.TrimEnd([System.IO.Path]::DirectorySeparatorChar)
    $targetClean = $TargetPath.TrimEnd([System.IO.Path]::DirectorySeparatorChar)

    if ($targetClean.StartsWith($baseClean, [System.StringComparison]::OrdinalIgnoreCase)) {
        $relative = $targetClean.Substring($baseClean.Length)
        if ($relative.StartsWith([System.IO.Path]::DirectorySeparatorChar)) {
            $relative = $relative.Substring(1)
        }
        return $relative
    }
    return ""
}

# Helper to safely release COM references
function Release-Ref ($object) {
    if ($null -ne $object) {
        try {
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) | Out-Null
        } catch {}
    }
}

# Force garbage collection to free up memory/processes
function Trigger-GC {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# Verification of input folder
if (-not (Test-Path $InputFolder -PathType Container)) {
    Write-Log "Input folder does not exist: $InputFolder" -Level "ERROR"
    exit 1
}

Write-Log "Starting PDF Conversion."
Write-Log "Input Folder: $InputFolder"
Write-Log "Output Folder: $OutputFolder"
if ($LogPath) { Write-Log "Logging to: $LogPath" }

# Scan for files
Write-Log "Scanning files..."
$files = Get-ChildItem -Path $InputFolder -Recurse -File

$wordFiles = $files | Where-Object { $_.Extension -match '^\.docx?$' }
$excelFiles = $files | Where-Object { $_.Extension -match '^\.xlsx?m?$' }
$pptFiles = $files | Where-Object { $_.Extension -match '^\.pptx?$' }

$totalFiles = ($wordFiles ? $wordFiles.Count : 0) + ($excelFiles ? $excelFiles.Count : 0) + ($pptFiles ? $pptFiles.Count : 0)
Write-Log "Found $totalFiles supported files ($($wordFiles ? $wordFiles.Count : 0) Word, $($excelFiles ? $excelFiles.Count : 0) Excel, $($pptFiles ? $pptFiles.Count : 0) PowerPoint)."

if ($totalFiles -eq 0) {
    Write-Log "No compatible files found."
    exit 0
}

$successCount = 0
$failCount = 0

# --- Process Word Files ---
if ($wordFiles) {
    Write-Log "Initializing Word.Application..."
    $word = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0 # wdAlertsNone

        foreach ($file in $wordFiles) {
            $sourcePath = $file.FullName
            $relativeDir = Get-RelativePath -BasePath $InputFolder -TargetPath $file.DirectoryName
            $targetDir = Join-Path $OutputFolder $relativeDir
            if (-not (Test-Path $targetDir)) {
                New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
            }
            $targetName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".pdf"
            $pdfPath = Join-Path $targetDir $targetName

            Write-Log "Word -> PDF: $($file.Name)"
            $doc = $null
            try {
                # Open(FileName, ConfirmConversions, ReadOnly)
                $doc = $word.Documents.Open($sourcePath, $false, $true)
                
                # ExportAsFixedFormat(OutputFileName, ExportFormat=17 [wdExportFormatPDF])
                $doc.ExportAsFixedFormat($pdfPath, 17)
                $doc.Close($false) # Close(SaveChanges=$false)
                
                Write-Log "  [SUCCESS] Created: $targetName"
                $successCount++
            } catch {
                Write-Log "  [ERROR] Failed to convert $($file.Name). Details: $_" -Level "ERROR"
                $failCount++
            } finally {
                Release-Ref $doc
            }
        }
    } catch {
        Write-Log "Failed to execute Word automation: $_" -Level "ERROR"
    } finally {
        if ($null -ne $word) {
            $word.Quit()
            Release-Ref $word
            Trigger-GC
        }
    }
}

# --- Process Excel Files ---
if ($excelFiles) {
    Write-Log "Initializing Excel.Application..."
    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        foreach ($file in $excelFiles) {
            $sourcePath = $file.FullName
            $relativeDir = Get-RelativePath -BasePath $InputFolder -TargetPath $file.DirectoryName
            $targetDir = Join-Path $OutputFolder $relativeDir
            if (-not (Test-Path $targetDir)) {
                New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
            }
            $targetName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".pdf"
            $pdfPath = Join-Path $targetDir $targetName

            Write-Log "Excel -> PDF: $($file.Name)"
            $wb = $null
            try {
                # Open(Filename, UpdateLinks=0, ReadOnly=$true)
                $wb = $excel.Workbooks.Open($sourcePath, 0, $true)
                
                # Apply print settings to each sheet: Horizontal/Landscape, Fit all columns on one page
                foreach ($sheet in $wb.Worksheets) {
                    try {
                        $sheet.PageSetup.Orientation = 2         # 2 = xlLandscape
                        $sheet.PageSetup.Zoom = $false           # Must be $false for FitToPages to work
                        $sheet.PageSetup.FitToPagesWide = 1      # Fit columns to one page
                        $sheet.PageSetup.FitToPagesTall = $false # Scale height automatically
                    } catch {
                        Write-Log "  [WARNING] Could not apply page setup to sheet $($sheet.Name). Details: $_" -Level "WARNING"
                    }
                }
                
                # ExportAsFixedFormat(Type=0 [xlTypePDF], Filename)
                # By exporting the workbook object directly, the whole book is printed by default
                $wb.ExportAsFixedFormat(0, $pdfPath)
                $wb.Close($false)
                
                Write-Log "  [SUCCESS] Created: $targetName"
                $successCount++
            } catch {
                Write-Log "  [ERROR] Failed to convert $($file.Name). Details: $_" -Level "ERROR"
                $failCount++
            } finally {
                Release-Ref $wb
            }
        }
    } catch {
        Write-Log "Failed to execute Excel automation: $_" -Level "ERROR"
    } finally {
        if ($null -ne $excel) {
            $excel.Quit()
            Release-Ref $excel
            Trigger-GC
        }
    }
}

# --- Process PowerPoint Files ---
if ($pptFiles) {
    Write-Log "Initializing PowerPoint.Application..."
    $ppt = $null
    try {
        $ppt = New-Object -ComObject PowerPoint.Application
        # Do not set Visible = $false as it throws an error in some environments.
        # Instead, we open each presentation with WithWindow = msoFalse (0) to hide it.

        foreach ($file in $pptFiles) {
            $sourcePath = $file.FullName
            $relativeDir = Get-RelativePath -BasePath $InputFolder -TargetPath $file.DirectoryName
            $targetDir = Join-Path $OutputFolder $relativeDir
            if (-not (Test-Path $targetDir)) {
                New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
            }
            $targetName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".pdf"
            $pdfPath = Join-Path $targetDir $targetName

            Write-Log "PowerPoint -> PDF: $($file.Name)"
            $pres = $null
            try {
                # Open(FileName, ReadOnly=msoTrue[-1], Untitled=msoFalse[0], WithWindow=msoFalse[0])
                $pres = $ppt.Presentations.Open($sourcePath, -1, 0, 0)
                
                # SaveAs(FileName, FileFormat=32 [ppSaveAsPDF])
                $pres.SaveAs($pdfPath, 32)
                $pres.Close()
                
                Write-Log "  [SUCCESS] Created: $targetName"
                $successCount++
            } catch {
                Write-Log "  [ERROR] Failed to convert $($file.Name). Details: $_" -Level "ERROR"
                $failCount++
            } finally {
                Release-Ref $pres
            }
        }
    } catch {
        Write-Log "Failed to execute PowerPoint automation: $_" -Level "ERROR"
    } finally {
        if ($null -ne $ppt) {
            $ppt.Quit()
            Release-Ref $ppt
            Trigger-GC
        }
    }
}

# Final memory cleanup
Trigger-GC

Write-Log "PDF Conversion completed."
Write-Log "Summary: Successful: $successCount, Failed: $failCount"
