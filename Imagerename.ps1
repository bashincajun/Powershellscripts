Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Supported image extensions
$ImageExtensions = @('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', '.webp', '.heic', '.heif')
$VideoExtensions = @('.mp4', '.avi', '.mov', '.mkv', '.wmv', '.flv', '.webm', '.m4v', '.mpg', '.mpeg')

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Image Renamer with Log Export'
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false

# === Controls ===

# Folder
$labelFolder = New-Object System.Windows.Forms.Label
$labelFolder.Location = New-Object System.Drawing.Point(20, 20)
$labelFolder.Size = New-Object System.Drawing.Size(100, 20)
$labelFolder.Text = 'Source Folder:'
$form.Controls.Add($labelFolder)

$textBoxFolder = New-Object System.Windows.Forms.TextBox
$textBoxFolder.Location = New-Object System.Drawing.Point(130, 20)
$textBoxFolder.Size = New-Object System.Drawing.Size(340, 20)
$textBoxFolder.ReadOnly = $true
$form.Controls.Add($textBoxFolder)

$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Location = New-Object System.Drawing.Point(480, 18)
$buttonBrowse.Size = New-Object System.Drawing.Size(80, 24)
$buttonBrowse.Text = 'Browse...'
$form.Controls.Add($buttonBrowse)

# Base Name
$labelName = New-Object System.Windows.Forms.Label
$labelName.Location = New-Object System.Drawing.Point(20, 60)
$labelName.Size = New-Object System.Drawing.Size(100, 20)
$labelName.Text = 'Base Name:'
$form.Controls.Add($labelName)

$textBoxName = New-Object System.Windows.Forms.TextBox
$textBoxName.Location = New-Object System.Drawing.Point(130, 60)
$textBoxName.Size = New-Object System.Drawing.Size(430, 20)
$form.Controls.Add($textBoxName)

# Zero Padding
$checkPad = New-Object System.Windows.Forms.CheckBox
$checkPad.Location = New-Object System.Drawing.Point(130, 90)
$checkPad.Size = New-Object System.Drawing.Size(300, 20)
$checkPad.Text = 'Pad numbers to 3 digits (001, 002...)'
$checkPad.Checked = $true
$form.Controls.Add($checkPad)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 130)
$progressBar.Size = New-Object System.Drawing.Size(540, 25)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$progressBar.Style = 'Continuous'
$form.Controls.Add($progressBar)

# Progress Label
$labelProgress = New-Object System.Windows.Forms.Label
$labelProgress.Location = New-Object System.Drawing.Point(20, 165)
$labelProgress.Size = New-Object System.Drawing.Size(540, 20)
$labelProgress.Text = 'Ready.'
$form.Controls.Add($labelProgress)

# Start Button
$buttonRename = New-Object System.Windows.Forms.Button
$buttonRename.Location = New-Object System.Drawing.Point(180, 200)
$buttonRename.Size = New-Object System.Drawing.Size(140, 35)
$buttonRename.Text = 'Start Renaming'
$buttonRename.BackColor = [System.Drawing.Color]::LightGreen
$buttonRename.Enabled = $false
$form.Controls.Add($buttonRename)

# Save Log Button
$buttonSaveLog = New-Object System.Windows.Forms.Button
$buttonSaveLog.Location = New-Object System.Drawing.Point(330, 200)
$buttonSaveLog.Size = New-Object System.Drawing.Size(100, 35)
$buttonSaveLog.Text = 'Save Log'
$buttonSaveLog.BackColor = [System.Drawing.Color]::LightBlue
$buttonSaveLog.Visible = $false
$form.Controls.Add($buttonSaveLog)

# Log Label
$labelLog = New-Object System.Windows.Forms.Label
$labelLog.Location = New-Object System.Drawing.Point(20, 250)
$labelLog.Size = New-Object System.Drawing.Size(100, 20)
$labelLog.Text = 'Log Output:'
$form.Controls.Add($labelLog)

# === RICH TEXT BOX (Supports colors!) ===
$richTextLog = New-Object System.Windows.Forms.RichTextBox
$richTextLog.Location = New-Object System.Drawing.Point(20, 275)
$richTextLog.Size = New-Object System.Drawing.Size(540, 140)
$richTextLog.ReadOnly = $true
$richTextLog.BackColor = [System.Drawing.Color]::White
$richTextLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$richTextLog.ScrollBars = 'ForcedBoth'
$form.Controls.Add($richTextLog)

# === In-Memory Log for Export ===
$global:LogEntries = [System.Collections.Generic.List[string]]::new()

# === Write-Log with Color ===
function Write-Log {
    param(
        [string]$Message,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::Black
    )
    $timestamp = Get-Date -Format "HH:mm:ss"
    $line = "[$timestamp] $Message"
    
    # Save to memory
    $global:LogEntries.Add($line)
    
    # Append to RichTextBox with color
    $richTextLog.SelectionStart = $richTextLog.TextLength
    $richTextLog.SelectionLength = 0
    $richTextLog.SelectionColor = $Color
    $richTextLog.AppendText("$line`r`n")
    $richTextLog.SelectionColor = $richTextLog.ForeColor  # Reset
    $richTextLog.ScrollToCaret()
}

# === Browse Folder ===
$buttonBrowse.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = 'Select folder with images'
    if ($folderBrowser.ShowDialog() -eq 'OK') {
        $textBoxFolder.Text = $folderBrowser.SelectedPath
        $buttonRename.Enabled = $true
        Write-Log "Folder selected: $($folderBrowser.SelectedPath)" ([System.Drawing.Color]::Blue)
    }
})

# === Save Log to File ===
$buttonSaveLog.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $saveDialog.Title = "Save Rename Log"
    $saveDialog.FileName = "ImageRename_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    
    if ($saveDialog.ShowDialog() -eq 'OK') {
        try {
            $global:LogEntries | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show(
                "Log saved to:`n$($saveDialog.FileName)",
                "Log Saved",
                'OK',
                'Information'
            )
            Write-Log "Log exported to: $($saveDialog.FileName)" ([System.Drawing.Color]::Purple)
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to save log:`n$($_.Exception.Message)",
                "Save Error",
                'OK',
                'Error'
            )
        }
    }
})

# === Rename Action ===
$buttonRename.Add_Click({
    # Reset
    $global:LogEntries.Clear()
    $richTextLog.Clear()
    $buttonSaveLog.Visible = $false

    $folderPath = $textBoxFolder.Text.Trim()
    $baseName = $textBoxName.Text.Trim()
    $pad = $checkPad.Checked

    Write-Log "Starting rename process..." ([System.Drawing.Color]::Blue)

    if (-not (Test-Path $folderPath)) {
        [System.Windows.Forms.MessageBox]::Show("Folder not found!", "Error", 'OK', 'Error')
        return
    }
    if ([string]::IsNullOrWhiteSpace($baseName)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a base name.", "Input Required", 'OK', 'Warning')
        return
    }

    $files = Get-ChildItem -Path $folderPath -File | Where-Object {
        $ext = $_.Extension.ToLower()
        ($ImageExtensions -contains $ext) -and ($VideoExtensions -notcontains $ext)
    } | Sort-Object Name

    if ($files.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No image files found.", "No Files", 'OK', 'Information')
        Write-Log "No image files found." ([System.Drawing.Color]::Orange)
        $buttonSaveLog.Visible = $true
        return
    }

    $progressBar.Maximum = $files.Count
    $progressBar.Value = 0
    $labelProgress.Text = "Processing 0 of $($files.Count)..."
    $renamedCount = 0
    $errorCount = 0
    $counter = 1

    foreach ($file in $files) {
        $number = if ($pad) { $counter.ToString("D3") } else { $counter.ToString() }
        $newName = "$baseName`_$number$($file.Extension)"
        $newPath = Join-Path $folderPath $newName

        $uniqueName = $newName
        $i = 1
        while (Test-Path $newPath) {
            $uniqueName = "$baseName`_$number`_$i$($file.Extension)"
            $newPath = Join-Path $folderPath $uniqueName
            $i++
        }

        try {
            Rename-Item -Path $file.FullName -NewName $uniqueName -ErrorAction Stop
            $renamedCount++
            Write-Log "Renamed: $($file.Name) → $uniqueName" ([System.Drawing.Color]::Green)
        } catch {
            $errorCount++
            Write-Log "Failed: $($file.Name) - $($_.Exception.Message)" ([System.Drawing.Color]::Red)
        }

        $progressBar.Value = $counter
        $labelProgress.Text = "Processed $counter of $($files.Count) ($renamedCount renamed, $errorCount errors)"
        $form.Refresh()
        $counter++
    }

    # Final result
    $finalMsg = "$renamedCount file(s) renamed."
    if ($errorCount -gt 0) { $finalMsg += "`n$errorCount error(s) occurred." }
    $finalMsg += "`n`nYou can now save the log."

    $icon = if ($errorCount -gt 0) { 'Warning' } else { 'Information' }

    $labelProgress.Text = "Complete! $renamedCount renamed, $errorCount error(s)"
    Write-Log "Complete! $renamedCount renamed, $errorCount error(s)" ([System.Drawing.Color]::Purple)
    $buttonSaveLog.Visible = $true

    [System.Windows.Forms.MessageBox]::Show($finalMsg, "Renaming Complete", 'OK', $icon)
})

# Enable rename button
$textBoxFolder.Add_TextChanged({
    $buttonRename.Enabled = (-not [string]::IsNullOrWhiteSpace($textBoxFolder.Text))
})

# Show Form
$form.ShowDialog() | Out-Null