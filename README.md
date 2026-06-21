# hyron-data-prep

A PowerShell utility to recursively convert Microsoft Office documents (Word, Excel, PowerPoint) to PDF using native COM automation.

## Features

- **Recursive Processing:** Scans an input folder and maintains the source folder hierarchy in the output folder.
- **Silent Background Execution:** Converts documents without user prompts or alerts.
- **Customized Excel Print Settings:**
  - **Horizontal Orientation:** Pages are printed in landscape.
  - **Column Fitting:** Fits all columns on one page (`FitToPagesWide = 1`), with height scaling automatically.
  - **Whole Book:** Automatically exports all sheets within the workbook.
- **Memory & Process Management:** Safely releases COM objects and triggers garbage collection to prevent hung background processes.

## Requirements

- Windows OS
- PowerShell
- Microsoft Office (Word, Excel, PowerPoint) installed locally

## Usage

1. Open PowerShell.
2. Bypass execution policy for the current session (if required):
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```
3. Run the script:
   ```powershell
   & .\Office2Pdf.ps1 -InputFolder "C:\path\to\source" -OutputFolder "C:\path\to\destination"
   ```

### Parameters

- `-InputFolder` (Mandatory): The folder containing documents to convert.
- `-OutputFolder` (Optional): The directory to save the output PDFs (defaults to the `InputFolder`).
- `-LogPath` (Optional): Path to a log file where the results will be written.
