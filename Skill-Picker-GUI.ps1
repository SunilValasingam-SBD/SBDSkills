<#
.SYNOPSIS
    Graphical picker/launcher for all skills under this folder.

.DESCRIPTION
    Reads skills-manifest.json and shows a WinForms window:
      - left list of all skills
      - right panel with description, mode, expected inputs and results
      - dynamically generated input fields (with native Browse dialogs for file/folder params)
      - Run button that executes "script" skills and streams live output into a log box
      - "assisted" skills (content-authoring, run by the Copilot model) get a
        "Copy Chat Instruction" button instead of a Run button

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Skill-Picker-GUI.ps1
#>

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$root = $PSScriptRoot
$manifestPath = Join-Path $root 'skills-manifest.json'
if (-not (Test-Path $manifestPath)) {
    [System.Windows.Forms.MessageBox]::Show("Manifest not found: $manifestPath", "Skill Picker", "OK", "Error") | Out-Null
    exit 1
}
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$manifest = Get-Content $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$skills = @($manifest.skills)

# ---------------------------------------------------------------- Form shell
$form = New-Object System.Windows.Forms.Form
$form.Text = "Skill Picker"
$form.Size = New-Object System.Drawing.Size(1180, 760)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)

$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'
$split.SplitterDistance = 300
$form.Controls.Add($split)

# ---------------------------------------------------------------- Left: skill list
$listLabel = New-Object System.Windows.Forms.Label
$listLabel.Text = "Skills"
$listLabel.Dock = 'Top'
$listLabel.Font = New-Object System.Drawing.Font($listLabel.Font, [System.Drawing.FontStyle]::Bold)
$listLabel.Height = 24
$split.Panel1.Controls.Add($listLabel)

$skillList = New-Object System.Windows.Forms.ListBox
$skillList.Dock = 'Fill'
$skillList.Font = New-Object System.Drawing.Font("Segoe UI", 10)
foreach ($s in $skills) {
    $tag = if ($s.mode -eq 'assisted') { '[assisted]' } else { '[script]' }
    [void]$skillList.Items.Add("$($s.name)  $tag")
}
$split.Panel1.Controls.Add($skillList)
$skillList.BringToFront()

# ---------------------------------------------------------------- Right: details + inputs + log
$rightPanel = New-Object System.Windows.Forms.TableLayoutPanel
$rightPanel.Dock = 'Fill'
$rightPanel.ColumnCount = 1
$rightPanel.RowCount = 5
[void]$rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute), 70))
[void]$rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute), 190))
[void]$rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute), 40))
[void]$rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent), 100))
$split.Panel2.Controls.Add($rightPanel)

# Title + description + results
$detailBox = New-Object System.Windows.Forms.Label
$detailBox.AutoSize = $false
$detailBox.Dock = 'Top'
$detailBox.Height = 160
$detailBox.Padding = New-Object System.Windows.Forms.Padding(8)
$detailBox.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$rightPanel.Controls.Add($detailBox, 0, 0)

# Results summary
$resultsBox = New-Object System.Windows.Forms.Label
$resultsBox.AutoSize = $false
$resultsBox.Dock = 'Fill'
$resultsBox.Padding = New-Object System.Windows.Forms.Padding(8, 0, 8, 0)
$resultsBox.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Italic)
$resultsBox.ForeColor = [System.Drawing.Color]::DarkSlateGray
$rightPanel.Controls.Add($resultsBox, 0, 1)

# Dynamic inputs area (scrollable)
$inputsPanel = New-Object System.Windows.Forms.Panel
$inputsPanel.Dock = 'Fill'
$inputsPanel.AutoScroll = $true
$inputsPanel.BorderStyle = 'FixedSingle'
$rightPanel.Controls.Add($inputsPanel, 0, 2)

# Action buttons row
$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = 'Fill'
$buttonPanel.FlowDirection = 'LeftToRight'
$rightPanel.Controls.Add($buttonPanel, 0, 3)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = "Run Skill"
$runButton.Width = 120
$runButton.Height = 30
$runButton.Enabled = $false
$buttonPanel.Controls.Add($runButton)

$copyButton = New-Object System.Windows.Forms.Button
$copyButton.Text = "Copy Chat Instruction"
$copyButton.Width = 160
$copyButton.Height = 30
$copyButton.Visible = $false
$buttonPanel.Controls.Add($copyButton)

$openFolderButton = New-Object System.Windows.Forms.Button
$openFolderButton.Text = "Open Skill Folder"
$openFolderButton.Width = 140
$openFolderButton.Height = 30
$openFolderButton.Enabled = $false
$buttonPanel.Controls.Add($openFolderButton)

# Log output
$logLabel = New-Object System.Windows.Forms.Label
$logLabel.Text = "Output"
$logLabel.Dock = 'Top'
$logLabel.Font = New-Object System.Drawing.Font($logLabel.Font, [System.Drawing.FontStyle]::Bold)
$logLabel.Height = 20

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ReadOnly = $true
$logBox.ScrollBars = 'Vertical'
$logBox.Dock = 'Fill'
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logBox.BackColor = [System.Drawing.Color]::Black
$logBox.ForeColor = [System.Drawing.Color]::LightGreen

$logWrap = New-Object System.Windows.Forms.Panel
$logWrap.Dock = 'Fill'
$logWrap.Controls.Add($logBox)
$logWrap.Controls.Add($logLabel)
$rightPanel.Controls.Add($logWrap, 0, 4)

# ---------------------------------------------------------------- State
$script:selectedSkill = $null
$script:paramControls = @{}   # flag -> control (TextBox or CheckBox)

function Append-Log {
    param([string]$text)
    if ($logBox.InvokeRequired) {
        $logBox.Invoke([Action]{ Append-Log $text })
        return
    }
    $logBox.AppendText($text + [Environment]::NewLine)
}

function Guess-DialogKind {
    param($param)
    $prompt = "$($param.prompt)".ToLower()
    $flag = "$($param.flag)".ToLower()
    if ($flag -match 'directory' -or $prompt -match 'directory') { return 'folder' }
    if ($prompt -match '^output' -or $prompt -match 'output .*path' -or $flag -match 'output') { return 'save' }
    return 'open'
}

function Show-SkillDetails {
    param($s)

    $script:selectedSkill = $s
    $modeTxt = if ($s.mode -eq 'assisted') { "Assisted (run by the Copilot model, not a standalone script)" } else { "Script (runs directly: $($s.runner))" }
    $detailBox.Text = "$($s.name)`r`nMode: $modeTxt`r`n`r`n$($s.description)" + $(if ($s.note) { "`r`n`r`nNote: $($s.note)" } else { "" })
    $resultsBox.Text = "Results: $($s.outputs)"

    $inputsPanel.Controls.Clear()
    $script:paramControls = @{}
    $y = 8
    foreach ($p in $s.params) {
        $req = if ($p.required) { "required" } else { "optional" }
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = "$($p.flag)  ($($p.type), $req) - $($p.prompt)"
        $lbl.Left = 8; $lbl.Top = $y; $lbl.Width = 560
        $lbl.AutoSize = $false
        $lbl.Height = 18
        $inputsPanel.Controls.Add($lbl)
        $y += 20

        if ($p.type -eq 'switch') {
            $chk = New-Object System.Windows.Forms.CheckBox
            $chk.Left = 8; $chk.Top = $y; $chk.Width = 300
            $chk.Text = "Enable"
            $inputsPanel.Controls.Add($chk)
            $script:paramControls[$p.flag] = $chk
            $y += 28
        } else {
            $txt = New-Object System.Windows.Forms.TextBox
            $txt.Left = 8; $txt.Top = $y; $txt.Width = 460
            $inputsPanel.Controls.Add($txt)
            $script:paramControls[$p.flag] = $txt

            if ($p.type -eq 'path' -or $p.type -eq 'path-list') {
                $kind = Guess-DialogKind $p
                $browse = New-Object System.Windows.Forms.Button
                $browse.Left = 476; $browse.Top = $y - 1; $browse.Width = 80
                $browse.Text = "Browse..."
                $multiSelect = ($p.type -eq 'path-list')
                $browse.Add_Click({
                    param($btnSender, $btnArgs)
                    $targetTxt = $txt
                    switch ($kind) {
                        'folder' {
                            $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
                            if ($dlg.ShowDialog() -eq 'OK') { $targetTxt.Text = $dlg.SelectedPath }
                        }
                        'save' {
                            $dlg = New-Object System.Windows.Forms.SaveFileDialog
                            $dlg.Filter = "All files|*.*"
                            if ($dlg.ShowDialog() -eq 'OK') { $targetTxt.Text = $dlg.FileName }
                        }
                        default {
                            $dlg = New-Object System.Windows.Forms.OpenFileDialog
                            $dlg.Filter = "Data files|*.csv;*.xlsx;*.xls;*.docx;*.pptx;*.pdf|All files|*.*"
                            $dlg.Multiselect = $multiSelect
                            if ($dlg.ShowDialog() -eq 'OK') {
                                if ($multiSelect) { $targetTxt.Text = ($dlg.FileNames -join ';') }
                                else { $targetTxt.Text = $dlg.FileName }
                            }
                        }
                    }
                }.GetNewClosure())
                $inputsPanel.Controls.Add($browse)
            }
            $y += 28
        }
    }

    $runButton.Visible = ($s.mode -ne 'assisted')
    $runButton.Enabled = ($s.mode -ne 'assisted')
    $copyButton.Visible = ($s.mode -eq 'assisted')
    $openFolderButton.Enabled = $true
    $logBox.Clear()
}

$skillList.Add_SelectedIndexChanged({
    if ($skillList.SelectedIndex -ge 0) {
        Show-SkillDetails $skills[$skillList.SelectedIndex]
    }
})

$openFolderButton.Add_Click({
    if ($script:selectedSkill) {
        $dir = Join-Path $root $script:selectedSkill.folder
        if (Test-Path $dir) { Start-Process explorer.exe $dir }
    }
})

function Get-ParamValues {
    $values = @{}
    foreach ($p in $script:selectedSkill.params) {
        $ctrl = $script:paramControls[$p.flag]
        if ($p.type -eq 'switch') {
            $values[$p.flag] = $ctrl.Checked
        } else {
            $values[$p.flag] = $ctrl.Text
        }
    }
    return $values
}

function Validate-Required {
    foreach ($p in $script:selectedSkill.params) {
        if ($p.required) {
            $ctrl = $script:paramControls[$p.flag]
            if ($p.type -ne 'switch' -and [string]::IsNullOrWhiteSpace($ctrl.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Please provide a value for: $($p.prompt)", "Missing input", "OK", "Warning") | Out-Null
                $ctrl.Focus()
                return $false
            }
        }
    }
    return $true
}

$copyButton.Add_Click({
    if (-not $script:selectedSkill) { return }
    $values = Get-ParamValues
    $lines = @("Use the $($script:selectedSkill.id) skill.")
    foreach ($p in $script:selectedSkill.params) {
        $v = $values[$p.flag]
        if (-not [string]::IsNullOrWhiteSpace("$v")) { $lines += "- $($p.flag): $v" }
    }
    $text = $lines -join "`r`n"
    [System.Windows.Forms.Clipboard]::SetText($text)
    Append-Log "Copied to clipboard - paste into a Copilot CLI chat:"
    Append-Log $text
})

$runButton.Add_Click({
    if (-not $script:selectedSkill) { return }
    if (-not (Validate-Required)) { return }
    $s = $script:selectedSkill
    $skillDir = Join-Path $root $s.folder
    $entryPath = Join-Path $skillDir $s.entry
    if (-not (Test-Path $entryPath)) {
        [System.Windows.Forms.MessageBox]::Show("Entry script not found: $entryPath", "Error", "OK", "Error") | Out-Null
        return
    }

    $values = Get-ParamValues
    $argList = New-Object System.Collections.Generic.List[string]
    $argHash = @{}
    foreach ($p in $s.params) {
        $v = $values[$p.flag]
        $paramName = $p.flag.TrimStart('-')
        switch ($p.type) {
            'switch' {
                if ($v) { $argList.Add($p.flag); $argHash[$paramName] = $true }
            }
            'repeat-csv' {
                if (-not [string]::IsNullOrWhiteSpace("$v")) {
                    $items = @($v -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    foreach ($item in $items) { $argList.Add($p.flag); $argList.Add($item) }
                    $argHash[$paramName] = $items
                }
            }
            default {
                if (-not [string]::IsNullOrWhiteSpace("$v")) {
                    $argList.Add($p.flag); $argList.Add("$v")
                    $argHash[$paramName] = "$v"
                }
            }
        }
    }

    $logBox.Clear()
    $runButton.Enabled = $false
    Append-Log "Running $($s.name)..."

    # Build the underlying command line (python / powershell) the same way the
    # terminal launcher does, but execute it via Process + async stdout/stderr
    # events so the GUI log updates live instead of blocking until exit.
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.WorkingDirectory = $skillDir
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    if ($s.runner -eq 'python') {
        $psi.FileName = 'py'
        $psi.ArgumentList.Add($entryPath)
        foreach ($a in $argList) { $psi.ArgumentList.Add($a) }
    } else {
        $psi.FileName = 'powershell.exe'
        $psi.ArgumentList.Add('-ExecutionPolicy'); $psi.ArgumentList.Add('Bypass')
        $psi.ArgumentList.Add('-File'); $psi.ArgumentList.Add($entryPath)
        foreach ($key in $argHash.Keys) {
            $val = $argHash[$key]
            if ($val -is [bool]) {
                if ($val) { $psi.ArgumentList.Add("-$key") }
            } else {
                $psi.ArgumentList.Add("-$key")
                $psi.ArgumentList.Add("$val")
            }
        }
    }

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.EnableRaisingEvents = $true

    $outHandler = {
        if ($EventArgs.Data) { Append-Log $EventArgs.Data }
    }
    $errHandler = {
        if ($EventArgs.Data) { Append-Log "ERR: $($EventArgs.Data)" }
    }
    Register-ObjectEvent -InputObject $proc -EventName OutputDataReceived -Action $outHandler | Out-Null
    Register-ObjectEvent -InputObject $proc -EventName ErrorDataReceived -Action $errHandler | Out-Null
    Register-ObjectEvent -InputObject $proc -EventName Exited -Action {
        $form.Invoke([Action]{
            Append-Log "--- finished (exit code $($proc.ExitCode)) ---"
            $runButton.Enabled = $true
        })
    } | Out-Null

    try {
        [void]$proc.Start()
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()
    } catch {
        Append-Log "Failed to start: $($_.Exception.Message)"
        $runButton.Enabled = $true
    }
})

if ($skills.Count -gt 0) {
    $skillList.SelectedIndex = 0
}

[void]$form.ShowDialog()
