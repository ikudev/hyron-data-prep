<#
.SYNOPSIS
    Recursively converts Word, Excel, and PowerPoint files in a folder to PDF using default print/layout settings.
.DESCRIPTION
    This script traverses a folder recursively, identifies Office documents (Word, Excel, PowerPoint),
    and converts them to PDF using native COM automation APIs. It runs silently in the background
    without user prompts. Can run conversions in parallel for Word and Excel.
.PARAMETER InputFolder
    The directory to search recursively for documents to convert.
.PARAMETER OutputFolder
    The directory where the output PDF files will be stored. Maintains the source directory structure.
    If not specified, defaults to the InputFolder.
.PARAMETER LogPath
    Optional path to a log file where conversion logs will be appended.
.PARAMETER NoParallel
    Optional switch to disable parallel processing and run sequentially.
.PARAMETER ThrottleLimit
    Optional parameter to control the number of parallel workers. Defaults to 4.
.EXAMPLE
    .\Office2Pdf.ps1 -InputFolder "C:\Docs" -OutputFolder "C:\PDFs" -LogPath "C:\Docs\conversion.log"
#>

param (
    [Parameter(Mandatory=$true, HelpMessage="Path to the directory containing documents to convert.")]
    [string]$InputFolder,

    [Parameter(Mandatory=$false, HelpMessage="Path to the output directory for PDF files.")]
    [string]$OutputFolder,

    [Parameter(Mandatory=$false, HelpMessage="Path to a text log file.")]
    [string]$LogPath,

    [Parameter(Mandatory=$false, HelpMessage="Disable parallel processing and run sequentially.")]
    [switch]$NoParallel,

    [Parameter(Mandatory=$false, HelpMessage="Number of parallel processes to run.")]
    [int]$ThrottleLimit = 4
)

# Convert relative paths to absolute paths
$InputFolder = [System.IO.Path]::GetFullPath($InputFolder)
if ($LogPath) {
    $LogPath = [System.IO.Path]::GetFullPath($LogPath)
}

# Determine if we should run in parallel
$Parallel = -not $NoParallel

# Define sync lock for writing to log file from parallel runspaces
$LogLock = New-Object Object

# Relaunch under pwsh if running on Windows PowerShell 5.1 and Parallel is active
if ($Parallel -and $PSVersionTable.PSVersion.Major -lt 7) {
    if (Get-Command pwsh -ErrorAction SilentlyContinue) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "[$timestamp] [WARNING] Parallel mode requires PowerShell 7+. Relaunching script under pwsh..."
        
        $arguments = @()
        foreach ($key in $MyInvocation.BoundParameters.Keys) {
            $val = $MyInvocation.BoundParameters[$key]
            if ($val -is [switch]) {
                if ($val) { $arguments += "-$key" }
            } else {
                $arguments += "-$key"
                $arguments += $val
            }
        }
        
        & pwsh -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath $arguments
        exit $LASTEXITCODE
    } else {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "[$timestamp] [WARNING] Parallel mode was enabled, but PowerShell 7+ (pwsh) is not installed. Falling back to sequential execution."
        $Parallel = $false
    }
}

# Prompt user for OutputFolder if not specified in bound parameters
if (-not $MyInvocation.BoundParameters.ContainsKey('OutputFolder') -or [string]::IsNullOrWhiteSpace($OutputFolder)) {
    $userInput = Read-Host -Prompt "Enter the output folder path (Press Enter to use '$InputFolder')"
    if ([string]::IsNullOrWhiteSpace($userInput)) {
        $OutputFolder = $InputFolder
    } else {
        $OutputFolder = $userInput
    }
}

# Convert OutputFolder to absolute path
$OutputFolder = [System.IO.Path]::GetFullPath($OutputFolder)

# Logger helper for sequential mode
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

# Helper to resolve relative subdirectory structure (sequential)
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

# Helper to safely release COM references (sequential)
function Release-Ref ($object) {
    if ($null -ne $object) {
        try {
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) | Out-Null
        } catch {}
    }
}

# Force garbage collection to free up memory/processes (sequential)
function Trigger-GC {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# We define these function bodies as strings to be able to dynamically instantiate them inside parallel runspaces
$WriteLogCode = {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [string]$LogPath,
        [object]$LogLock
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMsg = "[$timestamp] [$Level] $Message"
    Write-Host $logMsg
    if ($LogPath) {
        [System.Threading.Monitor]::Enter($LogLock)
        try {
            $logMsg | Out-File -FilePath $LogPath -Append -Encoding utf8
        } catch {
            # Ignore file lock errors/retry silently
        } finally {
            [System.Threading.Monitor]::Exit($LogLock)
        }
    }
}.ToString()

$ReleaseRefCode = {
    param ($object)
    if ($null -ne $object) {
        try {
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) | Out-Null
        } catch {}
    }
}.ToString()

$GetRelativePathCode = {
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
}.ToString()

$TriggerGCCode = {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}.ToString()

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

$wordFiles = $files | Where-Object { $_.Extension -match '^\.docx?$' -and $_.Name -notlike '~$*' }
$excelFiles = $files | Where-Object { $_.Extension -match '^\.xlsx?m?$' -and $_.Name -notlike '~$*' }
$pptFiles = $files | Where-Object { $_.Extension -match '^\.pptx?$' -and $_.Name -notlike '~$*' }

$wordCount = if ($null -ne $wordFiles) { $wordFiles.Count } else { 0 }
$excelCount = if ($null -ne $excelFiles) { $excelFiles.Count } else { 0 }
$pptCount = if ($null -ne $pptFiles) { $pptFiles.Count } else { 0 }
$totalFiles = $wordCount + $excelCount + $pptCount

Write-Log "Found $totalFiles supported files ($wordCount Word, $excelCount Excel, $pptCount PowerPoint)."

if ($totalFiles -eq 0) {
    Write-Log "No compatible files found."
    exit 0
}

$successCount = 0
$failCount = 0

# --- Process Word Files ---
if ($wordFiles) {
    if ($Parallel) {
        Write-Log "Converting Word files in parallel (ThrottleLimit: $ThrottleLimit)..."
        $chunkCount = [Math]::Min($ThrottleLimit, $wordFiles.Count)
        $chunks = @()
        for ($i = 0; $i -lt $chunkCount; $i++) {
            $chunks += ,(New-Object System.Collections.Generic.List[Object])
        }
        for ($i = 0; $i -lt $wordFiles.Count; $i++) {
            $chunks[$i % $chunkCount].Add($wordFiles[$i])
        }

        $results = $chunks | ForEach-Object -ThrottleLimit $chunkCount -Parallel {
            $WriteLog = [scriptblock]::Create($using:WriteLogCode)
            $ReleaseRef = [scriptblock]::Create($using:ReleaseRefCode)
            $GetRelativePath = [scriptblock]::Create($using:GetRelativePathCode)
            $TriggerGC = [scriptblock]::Create($using:TriggerGCCode)

            $chunk = $_
            $word = $null
            $success = 0
            $fail = 0
            try {
                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                $word.DisplayAlerts = 0 # wdAlertsNone

                foreach ($file in $chunk) {
                    $sourcePath = $file.FullName
                    $relativeDir = & $GetRelativePath -BasePath $using:InputFolder -TargetPath $file.DirectoryName
                    $targetDir = Join-Path $using:OutputFolder $relativeDir
                    $targetDir = [System.IO.Path]::GetFullPath($targetDir)
                    if (-not (Test-Path $targetDir)) {
                        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                    }
                    $targetName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".pdf"
                    $pdfPath = Join-Path $targetDir $targetName
                    $pdfPath = [System.IO.Path]::GetFullPath($pdfPath)

                    & $WriteLog "Word -> PDF: $($file.Name)" "INFO" $using:LogPath $using:LogLock
                    $doc = $null
                    try {
                        # Open(FileName, ConfirmConversions, ReadOnly)
                        $doc = $word.Documents.Open($sourcePath, $false, $true)
                        
                        # ExportAsFixedFormat(OutputFileName, ExportFormat=17 [wdExportFormatPDF])
                        $doc.ExportAsFixedFormat($pdfPath, 17)
                        $doc.Close($false) # Close(SaveChanges=$false)
                        
                        & $WriteLog "  [SUCCESS] Created: $targetName" "INFO" $using:LogPath $using:LogLock
                        $success++
                    } catch {
                        & $WriteLog "  [ERROR] Failed to convert $($file.Name). Details: $_" "ERROR" $using:LogPath $using:LogLock
                        $fail++
                    } finally {
                        & $ReleaseRef $doc
                    }
                }
            } catch {
                & $WriteLog "Failed to execute Word automation: $_" "ERROR" $using:LogPath $using:LogLock
            } finally {
                if ($null -ne $word) {
                    $word.Quit()
                    & $ReleaseRef $word
                    & $TriggerGC
                }
            }
            [PSCustomObject]@{ SuccessCount = $success; FailCount = $fail }
        }

        foreach ($res in $results) {
            $successCount += $res.SuccessCount
            $failCount += $res.FailCount
        }
    } else {
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
                $targetDir = [System.IO.Path]::GetFullPath($targetDir)
                if (-not (Test-Path $targetDir)) {
                    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                }
                $targetName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".pdf"
                $pdfPath = Join-Path $targetDir $targetName
                $pdfPath = [System.IO.Path]::GetFullPath($pdfPath)

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
}

# --- Process Excel Files ---
if ($excelFiles) {
    if ($Parallel) {
        Write-Log "Converting Excel files in parallel (ThrottleLimit: $ThrottleLimit)..."
        $chunkCount = [Math]::Min($ThrottleLimit, $excelFiles.Count)
        $chunks = @()
        for ($i = 0; $i -lt $chunkCount; $i++) {
            $chunks += ,(New-Object System.Collections.Generic.List[Object])
        }
        for ($i = 0; $i -lt $excelFiles.Count; $i++) {
            $chunks[$i % $chunkCount].Add($excelFiles[$i])
        }

        $results = $chunks | ForEach-Object -ThrottleLimit $chunkCount -Parallel {
            $WriteLog = [scriptblock]::Create($using:WriteLogCode)
            $ReleaseRef = [scriptblock]::Create($using:ReleaseRefCode)
            $GetRelativePath = [scriptblock]::Create($using:GetRelativePathCode)
            $TriggerGC = [scriptblock]::Create($using:TriggerGCCode)

            $chunk = $_
            $excel = $null
            $success = 0
            $fail = 0
            try {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $excel.DisplayAlerts = $false

                foreach ($file in $chunk) {
                    $sourcePath = $file.FullName
                    $relativeDir = & $GetRelativePath -BasePath $using:InputFolder -TargetPath $file.DirectoryName
                    $targetDir = Join-Path $using:OutputFolder $relativeDir
                    $targetDir = [System.IO.Path]::GetFullPath($targetDir)
                    if (-not (Test-Path $targetDir)) {
                        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                    }
                    $targetName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".pdf"
                    $pdfPath = Join-Path $targetDir $targetName
                    $pdfPath = [System.IO.Path]::GetFullPath($pdfPath)

                    & $WriteLog "Excel -> PDF: $($file.Name)" "INFO" $using:LogPath $using:LogLock
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
                                & $WriteLog "  [WARNING] Could not apply page setup to sheet $($sheet.Name). Details: $_" "WARNING" $using:LogPath $using:LogLock
                            }
                        }
                        
                        # ExportAsFixedFormat(Type=0 [xlTypePDF], Filename, Quality=0, IncludeDocProperties=$true, IgnorePrintAreas=$true)
                        $wb.ExportAsFixedFormat(0, $pdfPath, 0, $true, $true)
                        $wb.Close($false)
                        
                        & $WriteLog "  [SUCCESS] Created: $targetName" "INFO" $using:LogPath $using:LogLock
                        $success++
                    } catch {
                        & $WriteLog "  [ERROR] Failed to convert $($file.Name). Details: $_" "ERROR" $using:LogPath $using:LogLock
                        $fail++
                    } finally {
                        & $ReleaseRef $wb
                    }
                }
            } catch {
                & $WriteLog "Failed to execute Excel automation: $_" "ERROR" $using:LogPath $using:LogLock
            } finally {
                if ($null -ne $excel) {
                    $excel.Quit()
                    & $ReleaseRef $excel
                    & $TriggerGC
                }
            }
            [PSCustomObject]@{ SuccessCount = $success; FailCount = $fail }
        }

        foreach ($res in $results) {
            $successCount += $res.SuccessCount
            $failCount += $res.FailCount
        }
    } else {
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
                $targetDir = [System.IO.Path]::GetFullPath($targetDir)
                if (-not (Test-Path $targetDir)) {
                    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                }
                $targetName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".pdf"
                $pdfPath = Join-Path $targetDir $targetName
                $pdfPath = [System.IO.Path]::GetFullPath($pdfPath)

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
                    
                    # ExportAsFixedFormat(Type=0 [xlTypePDF], Filename, Quality=0, IncludeDocProperties=$true, IgnorePrintAreas=$true)
                    $wb.ExportAsFixedFormat(0, $pdfPath, 0, $true, $true)
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
}

# --- Process PowerPoint Files ---
if ($pptFiles) {
    if ($Parallel) {
        Write-Log "Converting PowerPoint files sequentially (Parallel not supported for PowerPoint COM)..."
    }
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
            $targetDir = [System.IO.Path]::GetFullPath($targetDir)
            if (-not (Test-Path $targetDir)) {
                New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
            }
            $targetName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".pdf"
            $pdfPath = Join-Path $targetDir $targetName
            $pdfPath = [System.IO.Path]::GetFullPath($pdfPath)

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
