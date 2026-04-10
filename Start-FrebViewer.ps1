#Requires -Version 5.1
<#
.SYNOPSIS
    FREB (Failed Request Event Buffering) log viewer — WinForms GUI.

.DESCRIPTION
    Opens a Windows Forms window that lets you browse IIS FREB XML log files
    and display them rendered through their freb.xsl stylesheet.
    The script must be launched via FrebViewer.bat so that PowerShell starts
    in STA (Single-Threaded Apartment) mode, which is required for WinForms.

.PARAMETER FolderPath
    Optional. Pre-loads all FREB XML logs from this directory on startup.

.EXAMPLE
    .\Start-FrebViewer.ps1
    .\Start-FrebViewer.ps1 -FolderPath "C:\inetpub\logs\FailedReqLogFiles\W3SVC2"
#>
param(
    [string] $FolderPath
)

Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# Assemblies
# ---------------------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# ---------------------------------------------------------------------------
# Service modules
# ---------------------------------------------------------------------------
$servicesDir = Join-Path $PSScriptRoot 'Services'
. (Join-Path $servicesDir 'FrebFileService.ps1')
. (Join-Path $servicesDir 'FrebXmlService.ps1')

# ---------------------------------------------------------------------------
# Script-level constants & state
# ---------------------------------------------------------------------------
$script:XslFileName      = 'freb.xsl'
$script:CurrentDirectory = $null
$script:LoadedFiles      = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

# ---------------------------------------------------------------------------
# UI — Main window
# ---------------------------------------------------------------------------
$mainForm                  = [System.Windows.Forms.Form]::new()
$mainForm.Text             = 'FREB Log Viewer'
$mainForm.Width            = 520
$mainForm.Height           = 720
$mainForm.StartPosition    = 'CenterScreen'
$mainForm.BackColor        = [System.Drawing.Color]::White
$mainForm.MinimumSize      = [System.Drawing.Size]::new(460, 400)

# Toolbar panel
$toolbarPanel              = [System.Windows.Forms.Panel]::new()
$toolbarPanel.Dock         = 'Top'
$toolbarPanel.Height       = 70
$toolbarPanel.BackColor    = [System.Drawing.Color]::FromArgb(240, 240, 240)
$toolbarPanel.BorderStyle  = 'FixedSingle'
$mainForm.Controls.Add($toolbarPanel)

# "Open File..." button
$btnOpenFile               = [System.Windows.Forms.Button]::new()
$btnOpenFile.Text          = 'Open File...'
$btnOpenFile.Location      = [System.Drawing.Point]::new(10, 8)
$btnOpenFile.Size          = [System.Drawing.Size]::new(155, 26)
$btnOpenFile.BackColor     = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnOpenFile.ForeColor     = [System.Drawing.Color]::White
$toolbarPanel.Controls.Add($btnOpenFile)

# Current-directory label
$lblCurrentDirectory          = [System.Windows.Forms.Label]::new()
$lblCurrentDirectory.Text     = 'No folder loaded'
$lblCurrentDirectory.Location = [System.Drawing.Point]::new(10, 40)
$lblCurrentDirectory.Size     = [System.Drawing.Size]::new(480, 22)
$lblCurrentDirectory.TextAlign= 'MiddleLeft'
$lblCurrentDirectory.Font     = [System.Drawing.Font]::new('Segoe UI', 8)
$toolbarPanel.Controls.Add($lblCurrentDirectory)

# Log-file list
$fileListBox           = [System.Windows.Forms.ListBox]::new()
$fileListBox.Location  = [System.Drawing.Point]::new(10, 80)
$fileListBox.Size      = [System.Drawing.Size]::new(480, 570)
$fileListBox.Font      = [System.Drawing.Font]::new('Courier New', 8)
$fileListBox.Anchor    = 'Top,Bottom,Left,Right'
$mainForm.Controls.Add($fileListBox)

# Status bar
$statusBar      = [System.Windows.Forms.StatusBar]::new()
$statusBar.Text = 'Ready'
$mainForm.Controls.Add($statusBar)

# ---------------------------------------------------------------------------
# Functions
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Opens a file picker, loads the chosen file's folder, and auto-selects the file.
#>
function Select-FrebXmlFile {
    $fileDialog                 = [System.Windows.Forms.OpenFileDialog]::new()
    $fileDialog.Title           = 'Select a FREB XML log file'
    $fileDialog.Filter          = 'XML Files (*.xml)|*.xml|All Files (*.*)|*.*'
    $fileDialog.CheckFileExists = $true
    $fileDialog.Multiselect     = $false

    if ($fileDialog.ShowDialog() -ne 'OK') { return }

    $chosenFile   = $fileDialog.FileName
    $chosenFolder = Split-Path -Parent $chosenFile
    $chosenName   = Split-Path -Leaf   $chosenFile

    $script:CurrentDirectory          = $chosenFolder
    $lblCurrentDirectory.Text         = $chosenFolder
    Update-FileList

    # Auto-select the file the user picked
    $names = $script:LoadedFiles | ForEach-Object { $_.Name }
    $idx   = [Array]::IndexOf($names, $chosenName)
    if ($idx -ge 0) { $fileListBox.SelectedIndex = $idx }

    $statusBar.Text = "Loaded folder: $($script:LoadedFiles.Count) file(s)"
}

<#
.SYNOPSIS
    Repopulates the list box from $script:CurrentDirectory.
#>
function Update-FileList {
    $fileListBox.Items.Clear()
    $script:LoadedFiles.Clear()

    $files = Get-FrebLogFiles -Directory $script:CurrentDirectory

    if ($files.Count -eq 0) {
        $statusBar.Text = 'No FREB XML files found in the selected folder'
        return
    }

    foreach ($file in $files) {
        $script:LoadedFiles.Add($file)
        $dateStr = $file.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
        [void] $fileListBox.Items.Add("$dateStr  $($file.Name)")
    }

    $statusBar.Text = "Loaded $($files.Count) FREB XML log(s)"
}

<#
.SYNOPSIS
    Opens the selected log entry in a maximised viewer window.

.PARAMETER Index
    Zero-based index into $script:LoadedFiles.
#>
function Open-FrebLogEntry {
    param(
        [int] $Index
    )

    if ($Index -lt 0 -or $Index -ge $script:LoadedFiles.Count) { return }

    $logFile = $script:LoadedFiles[$Index]
    $xmlPath = $logFile.FullName
    $xslPath = Find-FrebXslFile -XmlDirectory $script:CurrentDirectory

    $statusBar.Text = "Opening: $($logFile.Name)"

    # --- Build HTML content --------------------------------------------------
    if ($null -ne $xslPath) {
        Add-Type -AssemblyName System.Web

        $htmlContent = Convert-XmlWithXsl -XmlPath $xmlPath -XslPath $xslPath
    }
    else {
        # No stylesheet found: show syntax-highlighted raw XML
        Add-Type -AssemblyName System.Web
        $raw         = Get-Content -LiteralPath $xmlPath -Raw
        $encoded     = [System.Web.HttpUtility]::HtmlEncode($raw)
        $htmlContent = "<html><body style='font-family:Courier New;font-size:12px'><pre>$encoded</pre></body></html>"

        $statusBar.Text = "Opening (no freb.xsl found - showing raw XML): $($logFile.Name)"
    }

    # --- Viewer window -------------------------------------------------------
    $viewForm              = [System.Windows.Forms.Form]::new()
    $viewForm.Text         = $logFile.Name
    $viewForm.WindowState  = 'Maximized'
    $viewForm.StartPosition= 'CenterScreen'

    $browser                       = [System.Windows.Forms.WebBrowser]::new()
    $browser.Dock                  = 'Fill'
    $browser.ScriptErrorsSuppressed= $true
    $browser.DocumentText          = $htmlContent
    $viewForm.Controls.Add($browser)

    $viewForm.Show()
    $statusBar.Text = "Opened: $($logFile.Name)"
}

# ---------------------------------------------------------------------------
# Event handlers
# ---------------------------------------------------------------------------
$btnOpenFile.Add_Click({ Select-FrebXmlFile })

$fileListBox.Add_SelectedIndexChanged({
    Open-FrebLogEntry -Index $fileListBox.SelectedIndex
})

$mainForm.Add_Load({
    if ($FolderPath -and (Test-Path -LiteralPath $FolderPath -PathType Container)) {
        $script:CurrentDirectory  = $FolderPath
        $lblCurrentDirectory.Text = $FolderPath
        Update-FileList
    }
})

# ---------------------------------------------------------------------------
# Application entry point
# ---------------------------------------------------------------------------
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($mainForm)
