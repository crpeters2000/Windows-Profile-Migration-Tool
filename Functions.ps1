# Functions.ps1 - Helper Functions Module for ProfileMigration
# Version: v2.13.0
# Description: Reusable utility functions extracted from ProfileMigration.ps1

function Log-Message {
    param(
        [string]$Message,
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$ts [$Level] $Message"
    
    # Store entry in global array
    $global:LogEntries += [PSCustomObject]@{
        Timestamp     = $ts
        Level         = $Level
        Message       = $Message
        FormattedLine = $line
    }
    
    # Filter by log level
    $messageLevel = $LogLevels[$Level]
    $configuredLevel = $LogLevels[$Config.LogLevel]
    if ($messageLevel -lt $configuredLevel) {
        # Still write to file, just don't display to UI
        if ($global:CurrentLogFile) { $line | Out-File $global:CurrentLogFile -Append -Encoding UTF8 }
        # Capture for conversion log if active
        if ($global:ConversionLogPath) { $line | Out-File $global:ConversionLogPath -Append -Encoding UTF8 }
        return
    }

    # UI Update
    if ($global:LogBox) {
        $global:LogBox.AppendText("$line`r`n")
        $global:LogBox.SelectionStart = $global:LogBox.TextLength
        $global:LogBox.ScrollToCaret()
    }
    
    # Update conversion form status label if active
    if ($global:ConversionStatusLabel -and $Level -eq 'INFO') {
        try {
            $global:ConversionStatusLabel.Text = $Message
            [System.Windows.Forms.Application]::DoEvents()
        }
        catch {
            # Silently ignore if form is closed or label is disposed
        }
    }
    
    # Console Output (for pre-GUI phase)
    $color = switch ($Level) {
        'DEBUG' { 'Gray' }
        'INFO' { 'White' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'White' }
    }
    Write-Host $line -ForegroundColor $color

    # File Update
    if ($global:CurrentLogFile) { $line | Out-File $global:CurrentLogFile -Append -Encoding UTF8 }
    
    # Capture for conversion log if active
    if ($global:ConversionLogPath) { $line | Out-File $global:ConversionLogPath -Append -Encoding UTF8 }
}


function Log-Debug { param($Message) Log-Message -Message $Message -Level 'DEBUG' }


function Log-Info { param($Message) Log-Message -Message $Message -Level 'INFO' }


function Log-Warning { param($Message) Log-Message -Message $Message -Level 'WARN' }


function Log-Error { param($Message) Log-Message -Message $Message -Level 'ERROR' }


function Refresh-LogDisplay {
    # Refresh the log display based on current filter level
    if (-not $global:LogBox) { return }
    
    $configuredLevel = $LogLevels[$Config.LogLevel]
    $global:LogBox.Clear()
    
    foreach ($entry in $global:LogEntries) {
        $messageLevel = $LogLevels[$entry.Level]
        if ($messageLevel -ge $configuredLevel) {
            $global:LogBox.AppendText("$($entry.FormattedLine)`r`n")
        }
    }
    
    # Scroll to bottom
    $global:LogBox.SelectionStart = $global:LogBox.TextLength
    $global:LogBox.ScrollToCaret()
}


function Show-ModernDialog {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [string]$Title = "",
        [ValidateSet('Info', 'Error', 'Warning', 'Success', 'Question')][string]$Type = 'Info',
        [ValidateSet('OK', 'OKCancel', 'YesNo')][string]$Buttons = 'OK'
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $form = New-Object System.Windows.Forms.Form
    $form.Text = if ($Title) { $Title } else { 'Message' }
    $form.Size = New-Object System.Drawing.Size(440, 320)  # Increased height for long messages/paths
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true
    
    # Apply theme to dialog
    $currentTheme = if ($global:CurrentTheme) { $global:CurrentTheme } else { "Light" }
    $theme = $Themes[$currentTheme]
    $form.BackColor = $theme.FormBackColor
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Header panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size = New-Object System.Drawing.Size(440, 48)
    # Use theme-aware header colors based on message type
    switch ($Type) {
        'Error' { 
            $headerPanel.BackColor = if ($currentTheme -eq "Dark") { [System.Drawing.Color]::FromArgb(139, 69, 19) } else { [System.Drawing.Color]::FromArgb(232, 17, 35) }
        }
        'Warning' { 
            $headerPanel.BackColor = if ($currentTheme -eq "Dark") { [System.Drawing.Color]::FromArgb(184, 134, 11) } else { [System.Drawing.Color]::FromArgb(255, 185, 0) }
        }
        'Success' { 
            $headerPanel.BackColor = if ($currentTheme -eq "Dark") { [System.Drawing.Color]::FromArgb(34, 139, 34) } else { [System.Drawing.Color]::FromArgb(16, 124, 16) }
        }
        'Question' { 
            $headerPanel.BackColor = $theme.HeaderTextColor  # Use theme header color for questions
        }
        default { 
            $headerPanel.BackColor = $theme.HeaderTextColor  # Use theme header color for default
        }
    }
    $form.Controls.Add($headerPanel)
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 12)
    $lblTitle.Size = New-Object System.Drawing.Size(400, 24)
    $lblTitle.Text = if ($Title) { $Title } else { 'Message' }
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::White
    $headerPanel.Controls.Add($lblTitle)
    # Message textbox for scrollable long messages
    $txtMsg = New-Object System.Windows.Forms.TextBox
    $txtMsg.Location = New-Object System.Drawing.Point(24, 60)
    $txtMsg.Size = New-Object System.Drawing.Size(390, 140)
    $txtMsg.Text = $Message
    $txtMsg.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $txtMsg.ForeColor = $theme.LabelTextColor
    $txtMsg.BackColor = $theme.FormBackColor
    $txtMsg.Multiline = $true
    $txtMsg.ReadOnly = $true
    $txtMsg.ScrollBars = "Vertical"
    $txtMsg.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $form.Controls.Add($txtMsg)
    # Buttons
    $result = $null
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Size = New-Object System.Drawing.Size(100, 32)
    $btnOK.Location = New-Object System.Drawing.Point(290, 230)  # Move button lower for taller dialog
    $btnOK.BackColor = $theme.ButtonPrimaryBackColor
    $btnOK.ForeColor = $theme.ButtonPrimaryForeColor
    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOK.FlatAppearance.BorderSize = 0
    $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOK.Add_MouseEnter({ $this.BackColor = $theme.ButtonPrimaryHoverBackColor })
    $btnOK.Add_MouseLeave({ $this.BackColor = $theme.ButtonPrimaryBackColor })
    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnOK
    $btnCancel = $null
    $btnYes = $null
    $btnNo = $null
    if ($Buttons -eq 'OK') {
        $form.Controls.Add($btnOK)
    }
    elseif ($Buttons -eq 'OKCancel') {
        $btnOK.Location = New-Object System.Drawing.Point(170, 230)
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = "Cancel"
        $btnCancel.Size = New-Object System.Drawing.Size(100, 32)
        $btnCancel.Location = New-Object System.Drawing.Point(290, 230)
        $btnCancel.BackColor = $theme.ButtonSecondaryBackColor
        $btnCancel.ForeColor = $theme.ButtonSecondaryForeColor
        $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnCancel.FlatAppearance.BorderSize = 0
        $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $btnCancel.Add_MouseEnter({ $this.BackColor = $theme.ButtonSecondaryHoverBackColor })
        $btnCancel.Add_MouseLeave({ $this.BackColor = $theme.ButtonSecondaryBackColor })
        $form.Controls.Add($btnOK)
        $form.Controls.Add($btnCancel)
        $form.CancelButton = $btnCancel
    }
    elseif ($Buttons -eq 'YesNo') {
        $btnYes = New-Object System.Windows.Forms.Button
        $btnYes.Text = "Yes"
        $btnYes.Size = New-Object System.Drawing.Size(100, 32)
        $btnYes.Location = New-Object System.Drawing.Point(170, 230)
        $btnYes.BackColor = $theme.ButtonPrimaryBackColor
        $btnYes.ForeColor = $theme.ButtonPrimaryForeColor
        $btnYes.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnYes.FlatAppearance.BorderSize = 0
        $btnYes.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $btnYes.DialogResult = [System.Windows.Forms.DialogResult]::Yes
        $btnYes.Add_MouseEnter({ $this.BackColor = $theme.ButtonPrimaryHoverBackColor })
        $btnYes.Add_MouseLeave({ $this.BackColor = $theme.ButtonPrimaryBackColor })
        $btnNo = New-Object System.Windows.Forms.Button
        $btnNo.Text = "No"
        $btnNo.Size = New-Object System.Drawing.Size(100, 32)
        $btnNo.Location = New-Object System.Drawing.Point(290, 230)
        $btnNo.BackColor = $theme.ButtonSecondaryBackColor
        $btnNo.ForeColor = $theme.ButtonSecondaryForeColor
        $btnNo.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnNo.FlatAppearance.BorderSize = 0
        $btnNo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $btnNo.DialogResult = [System.Windows.Forms.DialogResult]::No
        $btnNo.Add_MouseEnter({ $this.BackColor = $theme.ButtonSecondaryHoverBackColor })
        $btnNo.Add_MouseLeave({ $this.BackColor = $theme.ButtonSecondaryBackColor })
        $form.Controls.Add($btnYes)
        $form.Controls.Add($btnNo)
        $form.AcceptButton = $btnYes
        $form.CancelButton = $btnNo
    }
    $dialogResult = $form.ShowDialog()
    switch ($dialogResult) {
        'OK' { return 'OK' }
        'Cancel' { return 'Cancel' }
        'Yes' { return 'Yes' }
        'No' { return 'No' }
        default { return 'OK' }
    }
}


function Show-InputDialog {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Message,
        [string]$DefaultValue = "",
        [string]$ExampleText = ""
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    # Apply theme
    $currentTheme = if ($global:CurrentTheme) { $global:CurrentTheme } else { "Light" }
    $theme = $Themes[$currentTheme]
    
    $inputForm = New-Object System.Windows.Forms.Form
    $inputForm.Text = $Title
    $inputForm.Size = New-Object System.Drawing.Size(500, 280)
    $inputForm.StartPosition = "CenterScreen"
    $inputForm.FormBorderStyle = "FixedDialog"
    $inputForm.MaximizeBox = $false
    $inputForm.MinimizeBox = $false
    $inputForm.TopMost = $true
    $inputForm.BackColor = $theme.FormBackColor
    $inputForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Header panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size = New-Object System.Drawing.Size(500, 70)
    $headerPanel.BackColor = $theme.HeaderBackColor
    $inputForm.Controls.Add($headerPanel)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(460, 25)
    $lblTitle.Text = $Title
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $theme.HeaderTextColor
    $headerPanel.Controls.Add($lblTitle)

    $lblMessage = New-Object System.Windows.Forms.Label
    $lblMessage.Location = New-Object System.Drawing.Point(22, 43)
    $lblMessage.Size = New-Object System.Drawing.Size(460, 20)
    $lblMessage.Text = $Message
    $lblMessage.ForeColor = $theme.SubHeaderTextColor
    $headerPanel.Controls.Add($lblMessage)

    # Main content card
    $contentCard = New-Object System.Windows.Forms.Panel
    $contentCard.Location = New-Object System.Drawing.Point(15, 85)
    $contentCard.Size = New-Object System.Drawing.Size(460, 80)
    $contentCard.BackColor = $theme.PanelBackColor
    $inputForm.Controls.Add($contentCard)

    $lblInput = New-Object System.Windows.Forms.Label
    $lblInput.Location = New-Object System.Drawing.Point(20, 25)
    $lblInput.Size = New-Object System.Drawing.Size(420, 20)
    $lblInput.Text = $Message
    $lblInput.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblInput.ForeColor = $theme.LabelTextColor
    $contentCard.Controls.Add($lblInput)

    $txtInput = New-Object System.Windows.Forms.TextBox
    $txtInput.Location = New-Object System.Drawing.Point(20, 50)
    $txtInput.Size = New-Object System.Drawing.Size(420, 25)
    $txtInput.Text = $DefaultValue
    $txtInput.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtInput.BackColor = $theme.TextBoxBackColor
    $txtInput.ForeColor = $theme.TextBoxForeColor
    $contentCard.Controls.Add($txtInput)

    # Example text (if provided)
    if ($ExampleText) {
        $lblExample = New-Object System.Windows.Forms.Label
        $lblExample.Location = New-Object System.Drawing.Point(20, 78)
        $lblExample.Size = New-Object System.Drawing.Size(420, 20)
        $lblExample.Text = $ExampleText
        $lblExample.Font = New-Object System.Drawing.Font("Segoe UI", 8)
        $lblExample.ForeColor = $theme.SubHeaderTextColor
        $contentCard.Controls.Add($lblExample)
    }

    # Action buttons
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(250, 180)
    $btnOK.Size = New-Object System.Drawing.Size(110, 35)
    $btnOK.Text = "OK"
    $btnOK.BackColor = $theme.ButtonPrimaryBackColor
    $btnOK.ForeColor = $theme.ButtonPrimaryForeColor
    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOK.FlatAppearance.BorderSize = 0
    $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOK.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnOK.Add_MouseEnter({ $this.BackColor = $theme.ButtonPrimaryHoverBackColor })
    $btnOK.Add_MouseLeave({ $this.BackColor = $theme.ButtonPrimaryBackColor })
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $inputForm.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(370, 180)
    $btnCancel.Size = New-Object System.Drawing.Size(105, 35)
    $btnCancel.Text = "Cancel"
    $btnCancel.BackColor = $theme.ButtonSecondaryBackColor
    $btnCancel.ForeColor = $theme.ButtonSecondaryForeColor
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.FlatAppearance.BorderSize = 0
    $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCancel.Add_MouseEnter({ $this.BackColor = $theme.ButtonSecondaryHoverBackColor })
    $btnCancel.Add_MouseLeave({ $this.BackColor = $theme.ButtonSecondaryBackColor })
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $inputForm.Controls.Add($btnCancel)

    $inputForm.AcceptButton = $btnOK
    $inputForm.CancelButton = $btnCancel
    $txtInput.SelectAll()
    $txtInput.Focus()

    $result = $inputForm.ShowDialog()
    $inputValue = $txtInput.Text.Trim()
    $inputForm.Dispose()
    
    # Return hashtable with result and value
    return @{
        Result = $result
        Value  = $inputValue
    }
}


function Show-LogViewer {
    param(
        [Parameter(Mandatory = $true)][string]$LogPath,
        [string]$Title = "Log Viewer"
    )

    try {
        # Initial read
        $rawContent = if (Test-Path $LogPath) { Get-Content $LogPath -Encoding UTF8 -ErrorAction SilentlyContinue } else { "Log file not found: $LogPath" }
        if ($rawContent -is [string]) { $rawContent = @($rawContent) } # Ensure array

        # Apply theme to dialog
        $currentTheme = if ($global:CurrentTheme) { $global:CurrentTheme } else { "Light" }
        $theme = $Themes[$currentTheme]

        $form = New-Object System.Windows.Forms.Form
        $form.Text = $Title
        $form.Size = New-Object System.Drawing.Size(1000, 800)
        $form.StartPosition = 'CenterScreen'
        $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $form.BackColor = $theme.FormBackColor
        $form.MinimumSize = New-Object System.Drawing.Size(800, 600)

        # Bottom button panel (create first for proper docking order)
        $panel = New-Object System.Windows.Forms.Panel
        $panel.Dock = 'Bottom'
        $panel.Height = 60
        $panel.BackColor = $theme.PanelBackColor
        $form.Controls.Add($panel)

        # Header panel
        $headerPanel = New-Object System.Windows.Forms.Panel
        $headerPanel.Height = 100
        $headerPanel.BackColor = $theme.HeaderBackColor
        $headerPanel.Dock = 'Top'
        $form.Controls.Add($headerPanel)

        $lblHeader = New-Object System.Windows.Forms.Label
        $lblHeader.Location = New-Object System.Drawing.Point(20, 15)
        $lblHeader.Size = New-Object System.Drawing.Size(400, 30)
        $lblHeader.Text = "Migration Log Viewer"
        $lblHeader.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
        $lblHeader.ForeColor = $theme.HeaderTextColor
        $headerPanel.Controls.Add($lblHeader)

        # --- CONTROLS ---
        
        # Filter Level
        $lblFilter = New-Object System.Windows.Forms.Label
        $lblFilter.Text = "Filter Level:"
        $lblFilter.Location = New-Object System.Drawing.Point(25, 55)
        $lblFilter.AutoSize = $true
        $lblFilter.ForeColor = $theme.HeaderTextColor
        $headerPanel.Controls.Add($lblFilter)

        $cmbFilter = New-Object System.Windows.Forms.ComboBox
        $cmbFilter.Location = New-Object System.Drawing.Point(100, 52)
        $cmbFilter.Width = 100
        $cmbFilter.DropDownStyle = 'DropDownList'
        $cmbFilter.Items.AddRange(@("All", "INFO", "WARN", "ERROR", "DEBUG"))
        $cmbFilter.SelectedIndex = 0
        $headerPanel.Controls.Add($cmbFilter)

        # Search Box
        $lblSearch = New-Object System.Windows.Forms.Label
        $lblSearch.Text = "Search:"
        $lblSearch.Location = New-Object System.Drawing.Point(220, 55)
        $lblSearch.AutoSize = $true
        $lblSearch.ForeColor = $theme.HeaderTextColor
        $headerPanel.Controls.Add($lblSearch)

        $txtSearch = New-Object System.Windows.Forms.TextBox
        $txtSearch.Location = New-Object System.Drawing.Point(270, 52)
        $txtSearch.Width = 250
        $headerPanel.Controls.Add($txtSearch)
        
        # Refresh Button
        $btnRefresh = New-Object System.Windows.Forms.Button
        $btnRefresh.Text = "Refresh Log"
        $btnRefresh.Location = New-Object System.Drawing.Point(540, 50)
        $btnRefresh.Size = New-Object System.Drawing.Size(90, 26)
        $btnRefresh.BackColor = $theme.ButtonSecondaryBackColor
        $btnRefresh.ForeColor = $theme.ButtonSecondaryForeColor
        $btnRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnRefresh.FlatAppearance.BorderSize = 0
        $headerPanel.Controls.Add($btnRefresh)

        # Log content panel - use positioning instead of Dock=Fill to avoid overlap
        $contentPanel = New-Object System.Windows.Forms.Panel
        $contentPanel.Location = New-Object System.Drawing.Point(0, 100)  # Start after header (height=100)
        $contentPanel.Size = New-Object System.Drawing.Size(1000, 580)    # Form height 800 - header 100 - bottom 60 - padding 60 = 580
        $contentPanel.Anchor = 'Top,Left,Right,Bottom'  # Resize with form
        $contentPanel.Padding = New-Object System.Windows.Forms.Padding(15)
        $contentPanel.BackColor = $theme.FormBackColor
        $form.Controls.Add($contentPanel)

        $txtLog = New-Object System.Windows.Forms.TextBox
        $txtLog.Multiline = $true
        $txtLog.ReadOnly = $true
        $txtLog.ScrollBars = 'Both'
        $txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
        $txtLog.Dock = 'Fill'
        $txtLog.BackColor = $theme.LogBoxBackColor
        $txtLog.ForeColor = $theme.LogBoxForeColor
        $txtLog.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $contentPanel.Controls.Add($txtLog)

        # Bottom panel buttons (panel already created above)
        $btnOpen = New-Object System.Windows.Forms.Button
        $btnOpen.Text = 'Open in Notepad'
        $btnOpen.Width = 150
        $btnOpen.Height = 35
        $btnOpen.Location = New-Object System.Drawing.Point(15, 12)
        $btnOpen.BackColor = $theme.ButtonPrimaryBackColor
        $btnOpen.ForeColor = $theme.ButtonPrimaryForeColor
        $btnOpen.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnOpen.FlatAppearance.BorderSize = 0
        $btnOpen.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
        $btnOpen.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnOpen.Add_MouseEnter({ $this.BackColor = $theme.ButtonPrimaryHoverBackColor })
        $btnOpen.Add_MouseLeave({ $this.BackColor = $theme.ButtonPrimaryBackColor })
        $btnOpen.Add_Click({ Start-Process notepad.exe -ArgumentList "`"$LogPath`"" })
        $panel.Controls.Add($btnOpen)

        $btnClose = New-Object System.Windows.Forms.Button
        $btnClose.Text = 'Close'
        $btnClose.Width = 110
        $btnClose.Height = 35
        $btnClose.Location = New-Object System.Drawing.Point(180, 12)
        $btnClose.BackColor = $theme.ButtonSecondaryBackColor
        $btnClose.ForeColor = $theme.ButtonSecondaryForeColor
        $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnClose.FlatAppearance.BorderSize = 0
        $btnClose.Add_MouseEnter({ $this.BackColor = $theme.ButtonSecondaryHoverBackColor })
        $btnClose.Add_MouseLeave({ $this.BackColor = $theme.ButtonSecondaryBackColor })
        $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $btnClose.Add_Click({ $form.Close() })
        $panel.Controls.Add($btnClose)

        $form.CancelButton = $btnClose

        # --- LOGIC ---
        
        $UpdateLogView = {
            # Use .Text property for reliable string value
            $filter = $cmbFilter.Text
            if ([string]::IsNullOrWhiteSpace($filter)) {
                $filter = "All"
            }
            
            $search = $txtSearch.Text
            
            $filteredLines = $script:rawContent
            
            # Apply Level Filter
            if ($filter -and $filter -ne "All") {
                # $levelTag = "[$filter]" -> Unused
                
                # Match lines with timestamp format: YYYY-MM-DD HH:MM:SS [LEVEL]
                # This prevents matching header lines that just contain the level name as text
                $filteredLines = @($filteredLines | Where-Object { 
                        $_ -match "^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\s+\[$filter\]"
                    })
            }
            
            # Apply Search
            if (-not [string]::IsNullOrWhiteSpace($search)) {
                $filteredLines = $filteredLines | Where-Object { $_ -like "*$search*" }
            }
            
            $txtLog.Text = $filteredLines -join "`r`n"
            $txtLog.Refresh()  # Force UI refresh
            $txtLog.Select($txtLog.Text.Length, 0)
            $txtLog.ScrollToCaret()
        }
        
        # Wire Events
        $cmbFilter.Add_SelectedIndexChanged($UpdateLogView)
        $txtSearch.Add_TextChanged($UpdateLogView)
        
        $btnRefresh.Add_Click({
                if (Test-Path $LogPath) {
                    $script:rawContent = Get-Content $LogPath -Encoding UTF8 -ErrorAction SilentlyContinue
                    if ($script:rawContent -is [string]) { $script:rawContent = @($script:rawContent) }
                    & $UpdateLogView
                }
            })
        
        # Initial Load
        & $UpdateLogView
        
        $form.ShowDialog() | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error opening log viewer: $_", "Error", "OK", "Error")
    }
}


function Test-InternetConnectivity {
    $endpoints = @(
        "login.microsoftonline.com",
        "enterprise.registration.windows.net"
    )
    
    foreach ($endpoint in $endpoints) {
        try {
            # Try simple ping first (fastest)
            if (Test-Connection -ComputerName $endpoint -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                return $true
            }
            
            # If ping fails (firewall?), try TCP port 443
            if (Test-NetConnection -ComputerName $endpoint -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue) {
                return $true
            }
        }
        catch {
            # Continue to next endpoint
        }
    }
    
    return $false
}


function Get-WindowsTheme {
    try {
        # Check Windows registry for theme preference
        $themeKey = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue
        if ($themeKey -and $themeKey.AppsUseLightTheme -eq 0) {
            return "Dark"
        }
        else {
            return "Light"
        }
    }
    catch {
        return "Light"  # Default to light theme if detection fails
    }
}


function Apply-Theme {
    param([string]$ThemeName)
    
    if (-not $Themes.ContainsKey($ThemeName)) {
        $ThemeName = "Light"
    }
    
    $global:CurrentTheme = $ThemeName
    
    # Robust lookup to handle potential key issues
    $theme = $null
    if ($global:Themes) {
        foreach ($key in $global:Themes.Keys) {
            if ($key -eq $ThemeName) {
                $theme = $global:Themes[$key]
                break
            }
        }
        
        # Fallback if specific lookup fails but keys exist
        if (-not $theme -and $global:Themes.Count -gt 0) {
            foreach ($key in $global:Themes.Keys) {
                if ($key -ieq $ThemeName) {
                    $theme = $global:Themes[$key]
                    break
                }
            }
        }
    }
    
    # Apply to main form
    if ($global:Form) {
        $global:Form.BackColor = $theme.FormBackColor
    }
    
    # Apply to header panel
    $headerPanel.BackColor = $theme.HeaderBackColor
    $lblHeader.ForeColor = $theme.HeaderTextColor
    $lblSubheader.ForeColor = $theme.SubHeaderTextColor
    
    # Apply to main card panel
    $cardPanel.BackColor = $theme.PanelBackColor
    
    # Apply to domain card
    $domainCard.BackColor = $theme.PanelBackColor
    
    # Apply to progress card
    $progressCard.BackColor = $theme.PanelBackColor
    
    # Apply to buttons
    $btnSettings.BackColor = $theme.PanelBackColor
    $btnSettings.ForeColor = $theme.HeaderTextColor
    $btnSettings.FlatAppearance.BorderColor = $theme.BorderColor
    
    $global:BrowseButton.BackColor = $theme.ButtonSecondaryBackColor
    $global:BrowseButton.ForeColor = $theme.ButtonSecondaryForeColor
    
    $global:ExportButton.BackColor = $theme.ButtonPrimaryBackColor
    $global:ExportButton.ForeColor = $theme.ButtonPrimaryForeColor
    
    $global:ImportButton.BackColor = $theme.ButtonSuccessBackColor
    $global:ImportButton.ForeColor = $theme.ButtonSuccessForeColor
    
    $global:CancelButton.BackColor = $theme.ButtonDangerBackColor
    $global:CancelButton.ForeColor = $theme.ButtonDangerForeColor
    
    $btnSetTargetUser.BackColor = $theme.ButtonPrimaryBackColor
    $btnSetTargetUser.ForeColor = $theme.ButtonPrimaryForeColor
    
    $btnRefreshProfiles.BackColor = $theme.ButtonPrimaryBackColor
    $btnRefreshProfiles.ForeColor = $theme.ButtonPrimaryForeColor
    
    $global:ViewLogButton.BackColor = $theme.ButtonSecondaryBackColor
    $global:ViewLogButton.ForeColor = $theme.ButtonSecondaryForeColor
    
    $global:DomainJoinButton.BackColor = $theme.ButtonPrimaryBackColor
    $global:DomainJoinButton.ForeColor = $theme.ButtonPrimaryForeColor
    
    $global:DomainRetryButton.BackColor = $theme.ButtonSecondaryBackColor
    $global:DomainRetryButton.ForeColor = $theme.ButtonSecondaryForeColor
    
    # Apply to text boxes
    $global:UserComboBox.BackColor = $theme.TextBoxBackColor
    $global:UserComboBox.ForeColor = $theme.TextBoxForeColor
    
    $global:ComputerNameTextBox.BackColor = $theme.TextBoxBackColor
    $global:ComputerNameTextBox.ForeColor = $theme.TextBoxForeColor
    
    $global:DomainNameTextBox.BackColor = $theme.TextBoxBackColor
    $global:DomainNameTextBox.ForeColor = $theme.TextBoxForeColor
    
    $global:DelayTextBox.BackColor = $theme.TextBoxBackColor
    $global:DelayTextBox.ForeColor = $theme.TextBoxForeColor
    
    # Apply to combo boxes
    $global:RestartComboBox.BackColor = $theme.TextBoxBackColor
    $global:RestartComboBox.ForeColor = $theme.TextBoxForeColor
    
    $global:LogLevelComboBox.BackColor = $theme.TextBoxBackColor
    $global:LogLevelComboBox.ForeColor = $theme.TextBoxForeColor
    
    # Apply to labels
    $lblU.ForeColor = $theme.LabelTextColor
    $lblC.ForeColor = $theme.LabelTextColor
    $lblD.ForeColor = $theme.LabelTextColor
    $lblRestart.ForeColor = $theme.LabelTextColor
    $lblDelay.ForeColor = $theme.LabelTextColor
    $lblLogLevel.ForeColor = $theme.SubHeaderTextColor
    $global:FileLabel.ForeColor = $theme.SubHeaderTextColor
    $lblUserTypeStatus.ForeColor = $theme.HeaderTextColor
    $lblDomainHeader.ForeColor = $theme.HeaderTextColor
    $lblLogHeader.ForeColor = $theme.HeaderTextColor
    
    # Apply to log box
    $global:LogBox.BackColor = $theme.LogBoxBackColor
    $global:LogBox.ForeColor = $theme.LogBoxForeColor
    
    # Apply to progress bar
    $global:ProgressBar.ForeColor = $theme.ProgressBarForeColor
    
    # Apply to status text
    $global:StatusText.ForeColor = $theme.LabelTextColor
    
    # Apply to checkboxes
    $global:DomainCheckBox.ForeColor = $theme.LabelTextColor
    $global:DebugCheckBox.ForeColor = $theme.LabelTextColor

    
    # Update theme toggle button text and colors
    if ($global:ThemeToggleButton) {
        $global:ThemeToggleButton.Text = if ($ThemeName -eq "Light") { "L" } else { "D" }
        $global:ThemeToggleButton.BackColor = $theme.HeaderBackColor
        $global:ThemeToggleButton.ForeColor = $theme.HeaderTextColor
        $global:ThemeToggleButton.FlatAppearance.BorderColor = $theme.BorderColor
    }
    
    # Update settings button colors
    if ($btnSettings) {
        $btnSettings.BackColor = $theme.HeaderBackColor
        $btnSettings.ForeColor = $theme.HeaderTextColor
        $btnSettings.FlatAppearance.BorderColor = $theme.BorderColor
    }
    
    # Force redraw
    if ($global:Form) {
        $global:Form.Refresh()
    }
}


function Toggle-Theme {
    $newTheme = if ($global:CurrentTheme -eq "Light") { "Dark" } else { "Light" }
    Apply-Theme -ThemeName $newTheme
}


function Test-PathWithRetry {
    <#
    .SYNOPSIS
    Tests if a path exists with retry logic
    
    .DESCRIPTION
    Retries path existence check with delays for timing-sensitive operations
    
    .PARAMETER Path
    The path to test
    
    .PARAMETER MaxAttempts
    Maximum number of retry attempts (default: 3)
    
    .PARAMETER DelayMs
    Delay in milliseconds between attempts (default: 500)
    
    .OUTPUTS
    Returns $true if path exists, $false otherwise
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$MaxAttempts = 3,
        [int]$DelayMs = 500
    )
    
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            if (Test-Path $Path -ErrorAction Stop) {
                if ($attempt -gt 1) {
                    Log-Message "Path found on attempt $attempt - $Path" 'DEBUG'
                }
                return $true
            }
        }
        catch {
            Log-Message "Attempt $attempt to access '$Path' failed - $_" 'DEBUG'
        }
        
        if ($attempt -lt $MaxAttempts) {
            Start-Sleep -Milliseconds $DelayMs
        }
    }
    
    return $false
}


function Get-ProfileInfo {
    param(
        [string]$ProfilePath,
        [switch]$SkipSizeCalculation
    )
    
    try {
        # Check if NTUSER.DAT exists (indicates valid profile)
        $hasHive = Test-Path (Join-Path $ProfilePath "NTUSER.DAT")
        
        # Skip expensive size calculation during initial load
        if ($SkipSizeCalculation) {
            return @{
                Path      = $ProfilePath
                SizeMB    = -1  # -1 indicates not calculated
                ItemCount = 0
                HasHive   = $hasHive
                IsValid   = $hasHive
            }
        }
        
        # Get-FolderSize handles recursive calculation robustly
        $sizeBytes = Get-FolderSize -Path $ProfilePath
        
        # Get rough item count (files + folders in root)
        $itemCount = (Get-ChildItem -Path $ProfilePath -Force -ErrorAction SilentlyContinue).Count
        
        $sizeMB = [math]::Round($sizeBytes / 1MB, 1)
        
        return @{
            Path      = $ProfilePath
            SizeMB    = $sizeMB
            ItemCount = $itemCount
            HasHive   = $hasHive
            IsValid   = $hasHive
        }
    }
    catch {
        return @{
            Path      = $ProfilePath
            SizeMB    = 0
            ItemCount = 0
            HasHive   = $false
            IsValid   = $false
        }
    }
}


function Get-ProfileDisplayEntries {
    param([switch]$CalculateSizes)
    
    $base = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    $computer = $env:COMPUTERNAME
    $profiles = Get-LocalProfiles
    $result = @()
    foreach ($p in $profiles) {
        # Get profile info - skip size calculation during initial load for speed
        $profileInfo = Get-ProfileInfo -ProfilePath $p.Path -SkipSizeCalculation:(-not $CalculateSizes)
        
        $sidKey = Get-ChildItem $base -ErrorAction SilentlyContinue | Where-Object {
            (Get-ItemProperty $_.PSPath -Name ProfileImagePath -EA SilentlyContinue).ProfileImagePath -like "*\$($p.Username)"
        } | Select-Object -First 1
        $display = $p.Username
        if ($sidKey) {

            $sid = $sidKey.PSChildName
            # Convert-SIDToAccountName centralizes translation and error handling
            $parts = Convert-SIDToAccountName -SID $sid -ReturnParts
            
            if ($parts) {
                # Normalize display name: always prefix with computer name if local
                if ($parts.Domain -ieq $computer) {
                    $display = "$computer\$($parts.User)"
                }
                else {
                    $display = "$($parts.Domain)\$($parts.User)"
                }
            }
        }
        else {
            # no registry mapping; assume local
            if ($p.Username -match '\\') {
                $display = $p.Username
            }
            else {
                $display = "$computer\$($p.Username)"
            }
        }
        
        # Add size estimate to display name only if calculated
        if ($CalculateSizes -and $profileInfo.SizeMB -ge 0) {
            $sizeDisplay = if ($profileInfo.SizeMB -gt 1024) {
                "$([math]::Round($profileInfo.SizeMB/1024, 1)) GB"
            }
            else {
                "$($profileInfo.SizeMB) MB"
            }
            $displayWithSize = "$display - [$sizeDisplay]"
        }
        else {
            $displayWithSize = $display
        }
        
        $result += [pscustomobject]@{ 
            DisplayName = $displayWithSize
            Username    = $p.Username
            Path        = $p.Path
            SizeMB      = $profileInfo.SizeMB
            IsValid     = $profileInfo.IsValid
        }
    }
    return $result | Sort-Object -Property Username
}


function Get-ProfileType {
    param(
        [Parameter(Mandatory = $false)][string]$UserSID,
        [Parameter(Mandatory = $false)][string]$Username
    )
    
    try {
        # If username provided, get SID first
        if ($Username -and -not $UserSID) {
            try {
                $UserSID = Get-LocalUserSID -UserName $Username
            }
            catch {
                Log-Warning "Could not resolve SID for username: $Username"
                return "Unknown"
            }
        }
        
        if (-not $UserSID) {
            Log-Warning "No SID or Username provided to Get-ProfileType"
            return "Unknown"
        }
        
        # Check if AzureAD SID (S-1-12-1-...)
        if (Test-IsAzureADSID -SID $UserSID) {
            return "AzureAD"
        }
        
        # Check if it's a domain or local account by attempting SID translation
        try {
            $sidObj = New-Object System.Security.Principal.SecurityIdentifier($UserSID)
            $ntAccount = $sidObj.Translate([System.Security.Principal.NTAccount])
            $accountName = $ntAccount.Value
            
            # Check if domain part matches computer name (local) or is different (domain)
            if ($accountName -match '^(.+?)\\(.+)$') {
                $domain = $matches[1]
                $computer = $env:COMPUTERNAME
                
                if ($domain -ieq $computer) {
                    return "Local"
                }
                else {
                    return "Domain"
                }
            }
            else {
                # No domain separator, assume local
                return "Local"
            }
        }
        catch {
            Log-Warning "Could not translate SID to account name: $_"
            return "Unknown"
        }
    }
    catch {
        Log-Error "Error in Get-ProfileType: $_"
        return "Unknown"
    }
}


function Test-ValidProfilePath {
    <#
    .SYNOPSIS
    Validates that a path is a valid user profile folder
    
    .DESCRIPTION
    Checks if path exists and optionally verifies NTUSER.DAT presence
    
    .PARAMETER Path
    The profile path to validate
    
    .PARAMETER RequireNTUSER
    If specified, also checks for NTUSER.DAT file
    
    .PARAMETER ThrowOnError
    If specified, throws exception instead of returning $false
    
    .OUTPUTS
    Returns $true if valid, $false otherwise (or throws if ThrowOnError)
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$RequireNTUSER,
        [switch]$ThrowOnError
    )
    
    if (-not (Test-PathWithRetry -Path $Path)) {
        $msg = "Profile path not found: $Path"
        if ($ThrowOnError) { throw $msg }
        Log-Message "WARNING: $msg"
        return $false
    }
    
    if ($RequireNTUSER) {
        $ntuserPath = Join-Path $Path "NTUSER.DAT"
        if (-not (Test-PathWithRetry -Path $ntuserPath)) {
            $msg = "NTUSER.DAT not found in profile: $Path"
            if ($ThrowOnError) { throw $msg }
            Log-Message "WARNING: $msg"
            return $false
        }
    }
    
    return $true
}


function Test-ProfilePathWriteable {
    param([string]$Path)
    
    try {
        $timeout = New-TimeSpan -Seconds $Config.ProfileValidationTimeoutSec
        $sw = [Diagnostics.Stopwatch]::StartNew()

        # If path doesn't exist, try to create it (this is expected for imports to new machines)
        $createdDir = $false
        if (-not (Test-Path $Path)) {
            try {
                New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
                $createdDir = $true
                Log-Message "Created missing target directory for validation: $Path"
            }
            catch {
                throw "Could not create target directory: $_"
            }
        }

        # Handle network paths differently for test file creation
        if ($Path -match '^\\\\') {
            # Network path - try to create temp file
            $testFile = Join-Path $Path ("test_" + [Guid]::NewGuid().ToString())
            
            # Check permissions
            $null | Out-File -FilePath $testFile -Force -ErrorAction Stop

            # Verify it exists
            if (-not (Test-Path $testFile -PathType Leaf)) {
                throw "Test file created but cannot be verified"
            }

            # Clean up test file
            Remove-Item $testFile -Force -ErrorAction Stop
        }
        else {
            # Local path
            # Create test file inside target path
            $testFile = Join-Path $Path (".PROFILE_TEST_$(Get-Random)")
            $null | Out-File -FilePath $testFile -Force -ErrorAction Stop

            # Verify it exists
            if (-not (Test-Path $testFile -PathType Leaf)) {
                throw "Test file created but cannot be verified"
            }

            # Clean up test file
            Remove-Item $testFile -Force -ErrorAction Stop
        }

        $sw.Stop()
        Log-Message "Profile path validation successful ($($sw.ElapsedMilliseconds)ms)"
        return $true

    }
    catch {
        Log-Message "CRITICAL: Profile path not writable: $_"
        return $false
    }
    finally {
        # If we created the directory only for validation and it's still empty, leave it (import will use it).
        # We avoid deleting it to reduce surprising side-effects; creation confirms write access.
    }
}


function Test-ProfileMounted {
    param([string]$UserSID)
    
    try {
        $mounted = $false
        
        # Method 1: Check using WMI Win32_UserProfile (Most reliable for loaded status)
        try {
            $profile = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID = '$UserSID'" -ErrorAction Stop
            if ($profile -and $profile.Loaded) {
                Log-Message "WARNING: User profile SID $UserSID is detected as LOADED via WMI"
                $mounted = $true
            }
        }
        catch {
            # Fallback if WMI fails
            Log-Debug "WMI profile check failed: $_"
        }
        
        # Method 2: Check registry HKU if WMI didn't find it (Double check)
        if (-not $mounted) {
            if (Test-Path "Registry::HKU\$UserSID") {
                Log-Message "WARNING: User profile SID $UserSID is mounted in HKU (Registry check)"
                $mounted = $true
            }
        }
        
        return $mounted
    }
    catch {
        Log-Message "Could not verify if profile is mounted: $_"
        return $false
    }
}


function Test-UserLoggedOut {
    param([Parameter(Mandatory = $true)][string]$Username)
    
    try {
        # Strip domain prefix if present
        $shortUsername = $Username
        if ($Username -match '\\(.+)$') {
            $shortUsername = $matches[1]
        }
        
        # Method 1: Check using quser command
        try {
            $quserOutput = quser 2>&1 | Out-String
            # Parse quser output line by line for exact username match
            $lines = $quserOutput -split "`n"
            foreach ($line in $lines) {
                # quser format: USERNAME  SESSIONNAME  ID  STATE  IDLE TIME  LOGON TIME
                if ($line -match "^\s*(\S+)\s+") {
                    $loggedUser = $matches[1]
                    if ($loggedUser -eq $shortUsername) {
                        Log-Debug "User $Username is currently logged in (detected via quser)"
                        return $false
                    }
                }
            }
        }
        catch {
            # quser fails if no users logged in, which is fine
        }
        
        # Method 2: Check using WMI
        try {
            $loggedInUsers = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop | Select-Object -ExpandProperty UserName
            if ($loggedInUsers) {
                foreach ($user in $loggedInUsers) {
                    # Extract just the username from DOMAIN\Username format
                    $wmiUsername = $user
                    if ($user -match '\\(.+)$') {
                        $wmiUsername = $matches[1]
                    }
                    if ($wmiUsername -eq $shortUsername) {
                        Log-Debug "User $Username is currently logged in (detected via WMI)"
                        return $false
                    }
                }
            }
        }
        catch {
            Log-Debug "WMI check failed: $_"
        }
        
        # Method 3: Check for loaded registry hive
        try {
            $profileList = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
            $profiles = Get-ChildItem $profileList -ErrorAction SilentlyContinue
            
            foreach ($profile in $profiles) {
                $sid = $profile.PSChildName
                try {
                    $sidObj = New-Object System.Security.Principal.SecurityIdentifier($sid)
                    $ntAccount = $sidObj.Translate([System.Security.Principal.NTAccount])
                    
                    # Extract username from DOMAIN\Username or COMPUTER\Username format
                    $accountUsername = $ntAccount.Value
                    if ($ntAccount.Value -match '\\(.+)$') {
                        $accountUsername = $matches[1]
                    }
                    
                    if ($accountUsername -eq $shortUsername) {
                        # Check if hive is loaded (indicates user is logged in)
                        $hiveLoaded = Test-Path "Registry::HKU\$sid"
                        if ($hiveLoaded) {
                            Log-Debug "User $Username has loaded registry hive (may be logged in)"
                            return $false
                        }
                    }
                }
                catch {
                    # SID translation can fail, continue
                }
            }
        }
        catch {
            Log-Debug "Registry hive check failed: $_"
        }
        
        # If all checks pass, user is logged out
        Log-Debug "User $Username is logged out"
        return $true
    }
    catch {
        Log-Warning "Error checking if user is logged out: $_"
        # Err on the side of caution - assume user is logged in if we can't determine
        return $false
    }
}


function Test-ProfileConversionPreconditions {
    param(
        [Parameter(Mandatory = $true)][string]$SourceUsername,
        [Parameter(Mandatory = $true)][ValidateSet('Local', 'Domain', 'AzureAD')][string]$TargetType,
        [string]$SourceSID
    )
    
    $errors = @()
    $warnings = @()
    
    try {
        # Get source SID if not provided
        if (-not $SourceSID) {
            try {
                $SourceSID = Get-LocalUserSID -UserName $SourceUsername
            }
            catch {
                $errors += "Could not resolve SID for source user: $SourceUsername"
                return @{ Success = $false; Errors = $errors; Warnings = $warnings }
            }
        }
        
        # Check 1: User must be logged out
        Log-Info "Checking if user is logged out..."
        if (-not (Test-UserLoggedOut -Username $SourceUsername)) {
            $errors += "User '$SourceUsername' is currently logged in. Please log out before converting the profile."
        }
        
        # Check 2: Administrator privileges
        Log-Info "Checking administrator privileges..."
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdmin) {
            $errors += "Administrator privileges required for profile conversion"
        }
        
        # Check 3: Profile exists
        Log-Info "Checking if profile exists..."
        $profileList = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
        $profileKey = Get-ItemProperty -Path "$profileList\$SourceSID" -ErrorAction SilentlyContinue
        if (-not $profileKey) {
            $errors += "Profile not found in registry for SID: $SourceSID"
        }
        else {
            $profilePath = $profileKey.ProfileImagePath
            if (-not (Test-Path $profilePath)) {
                $errors += "Profile folder not found: $profilePath"
            }
            else {
                # Check for NTUSER.DAT
                $hiveFile = Join-Path $profilePath "NTUSER.DAT"
                if (-not (Test-Path $hiveFile)) {
                    $errors += "NTUSER.DAT not found in profile folder"
                }
            }
        }
        
        # Check 4: Target type specific requirements
        switch ($TargetType) {
            'Domain' {
                Log-Info "Checking domain connectivity..."
                # Check if computer is domain joined
                try {
                    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
                    if (-not $computerSystem.PartOfDomain) {
                        $errors += "Computer is not joined to a domain. Cannot convert to domain profile."
                    }
                    else {
                        Log-Info "Computer is domain-joined: $($computerSystem.Domain)"
                    }
                }
                catch {
                    $warnings += "Could not verify domain membership: $_"
                }
            }
            'AzureAD' {
                Log-Info "Checking AzureAD join status..."
                if (-not (Test-IsAzureADJoined)) {
                    $errors += "Computer is not joined to AzureAD/Entra ID. Cannot convert to AzureAD profile."
                }
                else {
                    Log-Info "Computer is AzureAD-joined"
                }
            }
            'Local' {
                # No special requirements for converting to local
                Log-Info "Target type is Local - no special requirements"
            }
        }
        
        # Check 5: Disk space (estimate 2x profile size needed for backup + conversion)
        Log-Info "Checking disk space..."
        try {
            if ($profileKey) {
                $profilePath = $profileKey.ProfileImagePath
                $drive = Split-Path $profilePath -Qualifier
                $driveInfo = Get-PSDrive -Name $drive.TrimEnd(':') -ErrorAction Stop
                
                # Use fast estimation instead of slow recursive scan
                # Estimate based on typical profile size (5-10 GB)
                $profileSize = 10GB  # Conservative estimate
                
                $requiredSpace = $profileSize * 2  # 2x for backup + conversion
                $freeSpace = $driveInfo.Free
                
                if ($freeSpace -lt $requiredSpace) {
                    $requiredGB = [Math]::Round($requiredSpace / 1GB, 2)
                    $freeGB = [Math]::Round($freeSpace / 1GB, 2)
                    $warnings += "Low disk space: $freeGB GB free, $requiredGB GB recommended"
                }
                else {
                    Log-Info "Disk space check passed: $([Math]::Round($freeSpace / 1GB, 2)) GB free"
                }
            }
        }
        catch {
            $warnings += "Could not verify disk space: $_"
        }
        
        # Return results
        $success = ($errors.Count -eq 0)
        return @{
            Success  = $success
            Errors   = $errors
            Warnings = $warnings
        }
    }
    catch {
        Log-Error "Error in Test-ProfileConversionPreconditions: $_"
        return @{
            Success  = $false
            Errors   = @("Unexpected error during precondition check: $_")
            Warnings = $warnings
        }
    }
}


function Get-LocalProfiles {
    Get-ChildItem "C:\Users" -Directory | Where-Object {
        $_.Name -notmatch "^Public$|^Default$|^Administrator$|^All Users$|^Default User$"
    } | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{ Username = $_.Name; Path = $_.FullName }
    }
}


function Get-LocalUserSID {
    param([Parameter(Mandatory = $true)][string]$UserName)
    
    # Handle AzureAD usernames (AzureAD\username format)
    if ($UserName -match '^AzureAD\\(.+)$') {
        $azureUser = $matches[1]
        Log-Message "Detected AzureAD user format: $azureUser"
        
        # Check if system is AzureAD joined
        if (-not (Test-IsAzureADJoined)) {
            Log-Message "WARNING: System is not AzureAD joined"
            throw "AzureAD_NOT_JOINED:$azureUser"
        }
        
        # Try to find the AzureAD account by username
        # AzureAD accounts are stored in ProfileList with S-1-12-1 SIDs
        try {
            $profileList = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
            $profiles = Get-ChildItem $profileList -ErrorAction Stop
            
            foreach ($profile in $profiles) {
                $sid = $profile.PSChildName
                if (Test-IsAzureADSID -SID $sid) {
                    $accountParts = Convert-SIDToAccountName -SID $sid -ReturnParts
                    if ($accountParts -and $accountParts.User -eq $azureUser) {
                        Log-Message "Found AzureAD account: $($accountParts.FullName) with SID: $sid"
                        return $sid
                    }
                }
            }
            
            # If not found locally, try Microsoft Graph API
            Log-Message "AzureAD user '$azureUser' not found on this system"
            Log-Message "Attempting to retrieve SID from Microsoft Graph..."
            
            try {
                # Validate that username is in email format (UPN required for Graph API)
                if ($azureUser -notmatch '@') {
                    throw "AzureAD username must be in email format (e.g., user@domain.com) for Microsoft Graph lookup. You provided: '$azureUser'"
                }
                
                $graphSID = Get-AzureADUserSID -UserPrincipalName $azureUser
                
                # Ensure SID is a string (not an object or array)
                $graphSIDString = [string]$graphSID
                
                # Debug logging
                Log-Message "Graph SID retrieved: '$graphSIDString' (Type: $($graphSIDString.GetType().Name), Length: $($graphSIDString.Length))"
                $isValidAzureADSID = Test-IsAzureADSID -SID $graphSIDString
                Log-Message "AzureAD SID validation result: $isValidAzureADSID"
                
                if ($graphSIDString -and $isValidAzureADSID) {
                    Log-Message "Successfully retrieved AzureAD SID from Microsoft Graph: $graphSIDString"
                    return $graphSIDString
                }
                else {
                    throw "Microsoft Graph did not return a valid AzureAD SID"
                }
            }
            catch {
                Log-Message "Failed to retrieve SID from Microsoft Graph: $_"
                throw "AzureAD user '$azureUser' has not logged into this device yet and could not be found via Microsoft Graph. Error: $_"
            }
        }
        catch {
            throw $_
        }
    }
    
    # Do NOT convert NetBIOS domain to FQDN for ACL/ownership operations.
    # Always use NetBIOS format for icacls/takeown.
    # If $UserName is in DOMAIN\user format, leave as-is.
    # If $UserName is in AzureAD format, handle as above.
    # No changes needed here.
    
    $ntAccount = New-Object System.Security.Principal.NTAccount($UserName)
    $sid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier])
    return $sid.Value
}


function Convert-SIDToAccountName {
    <#
    .SYNOPSIS
    Converts a SID to an account name
    
    .DESCRIPTION
    Translates SID to NTAccount with error handling and optional parsing
    
    .PARAMETER SID
    The SID string to convert
    
    .PARAMETER ReturnParts
    If specified, returns hashtable with Domain and User properties
    
    .OUTPUTS
    Returns account name string, or hashtable if ReturnParts is specified
    Returns $null on error
    #>
    param(
        [Parameter(Mandatory = $true)][string]$SID,
        [switch]$ReturnParts
    )
    
    try {
        $sidObj = New-Object System.Security.Principal.SecurityIdentifier($SID)
        $ntAccount = $sidObj.Translate([System.Security.Principal.NTAccount])
        $accountName = $ntAccount.Value
        
        if ($ReturnParts) {
            if ($accountName -match '^(.+?)\\(.+)$') {
                return @{
                    Domain   = $matches[1]
                    User     = $matches[2]
                    FullName = $accountName
                }
            }
            else {
                return @{
                    Domain   = $null
                    User     = $accountName
                    FullName = $accountName
                }
            }
        }
        
        return $accountName
    }
    catch {
        Log-Message "WARNING: Could not translate SID '$SID': $_"
        return $null
    }
}


function Test-IsAzureADSID {
    param([Parameter(Mandatory = $true)][string]$SID)
    # AzureAD/Entra ID SIDs start with S-1-12-1
    return $SID -match '^S-1-12-1-'
}


function Test-IsAzureADJoined {
    try {
        $dsreg = dsregcmd /status
        $azureAdJoined = $dsreg | Select-String "AzureAdJoined\s*:\s*YES" -Quiet
        return $azureAdJoined
    }
    catch {
        Log-Message "Could not determine AzureAD join status: $_"
        return $false
    }
}


function Get-AzureADUserSID {
    param([string]$UserPrincipalName)
    
    try {
        # Check if Microsoft.Graph.Users module is available
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
            Log-Info "Microsoft.Graph.Users module not found. Installing..."
            
            # Notify user about the NuGet prompt
            $message = "The Microsoft Graph PowerShell module needs to be installed.`r`n`r`n"
            $message += "You may see a prompt in the PowerShell window asking to install the NuGet provider.`r`n`r`n"
            $message += "Please press 'Y' (Yes) when prompted to continue.`r`n`r`n"
            $message += "This is a one-time installation."
            
            $null = Show-ModernDialog -Message $message -Title "Module Installation Required" -Type Info -Buttons OK
            
            # Install-Module with -Force should automatically bootstrap NuGet without prompting
            Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
            Log-Info "Microsoft.Graph.Users module installed successfully"
        }
        
        # Import the module
        Import-Module Microsoft.Graph.Users -ErrorAction Stop
        
        # Connect to Microsoft Graph (will prompt for authentication)
        Log-Info "Connecting to Microsoft Graph..."
        $null = Connect-MgGraph -Scopes "User.Read.All" -NoWelcome -ErrorAction Stop
        
        # Get the user's ObjectId
        Log-Info "Retrieving user information for: $UserPrincipalName"
        $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        
        if ($user -and $user.Id) {
            # Convert ObjectId to SID
            $sid = Convert-EntraObjectIdToSid -ObjectId $user.Id
            Log-Info "Converted ObjectId $($user.Id) to SID: $sid"
            
            # Store SID before disconnecting
            $sidString = [string]$sid
            
            # Disconnect from Graph
            $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
            
            # Validate the SID before returning
            if ([string]::IsNullOrWhiteSpace($sidString)) {
                throw "SID conversion resulted in empty value"
            }
            
            return $sidString
        }
        else {
            throw "User not found in AzureAD"
        }
    }
    catch {
        Log-Error "Failed to get AzureAD user SID: $_"
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        return $null
    }
}


function Convert-EntraObjectIdToSid {
    param([String]$ObjectId)
    
    # Parse the GUID into a byte array
    $bytes = [Guid]::Parse($ObjectId).ToByteArray()
    
    # Cloud SIDs use a specific prefix (S-1-12-1) followed by 4 UInt32 values
    $array = New-Object 'UInt32[]' 4
    [Buffer]::BlockCopy($bytes, 0, $array, 0, 16)
    
    # Join the segments to create the full SID string
    $sid = "S-1-12-1-" + ($array -join "-")
    return [string]$sid
}


function Update-ConversionProgress {
    <#
    .SYNOPSIS
    Updates conversion progress bar and status label
    
    .DESCRIPTION
    Safely updates global progress controls if they exist
    
    .PARAMETER PercentComplete
    Progress percentage (0-100)
    
    .PARAMETER StatusMessage
    Status message to display
    #>
    param(
        [Parameter(Mandatory = $true)][int]$PercentComplete,
        [string]$StatusMessage
    )
    
    if ($global:ConversionProgressBar) {
        $global:ConversionProgressBar.Value = [Math]::Min(100, [Math]::Max(0, $PercentComplete))
    }
    
    if ($StatusMessage -and $global:ConversionStatusLabel) {
        $global:ConversionStatusLabel.Text = $StatusMessage
    }
    
    if ($global:StatusText) {
        if ($StatusMessage) {
            $global:StatusText.Text = $StatusMessage
        }
    }
    
    [System.Windows.Forms.Application]::DoEvents()
}


function Update-ProfileListRegistry {
    param(
        [Parameter(Mandatory = $true)][string]$OldSID,
        [Parameter(Mandatory = $true)][string]$NewSID,
        [Parameter(Mandatory = $true)][string]$NewProfilePath
    )
    
    try {
        Log-Info "Updating ProfileList registry for SID conversion..."
        $profileListBase = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
        
        # Get old profile key
        $oldKeyPath = "$profileListBase\$OldSID"
        $oldKey = Get-ItemProperty -Path $oldKeyPath -ErrorAction SilentlyContinue
        
        if (-not $oldKey) {
            throw "Source profile registry key not found: $oldKeyPath"
        }
        
        Log-Info "Source registry key found: $oldKeyPath"
        
        # Create new key for new SID
        $newKeyPath = "$profileListBase\$NewSID"
        
        # Check if new key already exists
        if (Test-Path $newKeyPath) {
            Log-Warning "Target registry key already exists: $newKeyPath"
            $response = Show-ModernDialog -Message "A profile registry entry already exists for the target user.`r`n`r`nDo you want to overwrite it?" -Title "Profile Exists" -Type Warning -Buttons YesNo
            if ($response -ne "Yes") {
                throw "User cancelled: Target profile registry key already exists"
            }
            Log-Info "Removing existing target registry key..."
            Remove-Item -Path $newKeyPath -Recurse -Force
        }
        
        Log-Info "Creating new registry key: $newKeyPath"
        New-Item -Path $newKeyPath -Force | Out-Null
        
        # Copy all values from old key to new key
        $oldKey.PSObject.Properties | ForEach-Object {
            $propName = $_.Name
            $propValue = $_.Value
            
            # Skip PS* properties (metadata), State, and RefCount
            # State and RefCount track profile load state and MUST NOT be copied
            # Copying these values causes 0xc0000409 (buffer overrun) crashes in Explorer
            if ($propName -notmatch '^PS' -and $propName -ne 'State' -and $propName -ne 'RefCount') {
                try {
                    # Special handling for ProfileImagePath
                    if ($propName -eq 'ProfileImagePath') {
                        Log-Info "Setting ProfileImagePath to: $NewProfilePath"
                        Set-ItemProperty -Path $newKeyPath -Name $propName -Value $NewProfilePath -Force
                    }
                    else {
                        Set-ItemProperty -Path $newKeyPath -Name $propName -Value $propValue -Force
                        Log-Debug "Copied registry value: $propName"
                    }
                }
                catch {
                    Log-Warning "Could not copy registry value '$propName': $_"
                }
            }
            elseif ($propName -eq 'State' -or $propName -eq 'RefCount') {
                Log-Info "Skipping '$propName' value (must be regenerated by Windows)"
            }
        }
        
        Log-Info "Registry values copied to new key"
        
        # CRITICAL: Set State to 0 to prevent Windows from treating this as a temporary profile
        # The State key must exist and be set to 0, not deleted
        # If State is missing or non-zero, Windows creates a new profile folder instead of using the existing one
        try {
            Log-Info "Setting State registry value to 0 (prevents temporary profile)"
            Set-ItemProperty -Path $newKeyPath -Name "State" -Value 0 -Type DWord -Force
            Log-Info "State value set to 0 successfully"
        }
        catch {
            Log-Warning "Could not set State value: $_"
        }
        
        # Ask user if they want to delete the old key (ONLY if different from new key)
        if ($oldKeyPath -ne $newKeyPath) {
            $response = Show-ModernDialog -Message "Profile registry has been created for the new user.`r`n`r`nDo you want to delete the old profile registry entry?`r`n`r`n(Recommended: Yes)" -Title "Delete Old Registry Entry" -Type Question -Buttons YesNo
            
            if ($response -eq "Yes") {
                Log-Info "Deleting old registry key: $oldKeyPath"
                Remove-Item -Path $oldKeyPath -Recurse -Force
                Log-Info "Old registry key deleted"
            }
            else {
                Log-Info "Old registry key retained at user's request"
            }
        }
        else {
            Log-Info "Old and New registry keys are identical (Repair Mode). Skipping deletion."
        }
        
        return @{
            Success    = $true
            NewKeyPath = $newKeyPath
        }
    }
    catch {
        Log-Error "Failed to update ProfileList registry: $_"
        return @{
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}


function Set-ProfileAcls {
    param(
        [Parameter(Mandatory = $true)][string]$ProfileFolder,
        [Parameter(Mandatory = $true)][string]$UserName,
        [string]$SourceSID,
        [string]$UserSID,
        [string]$OldProfilePath,
        [string]$NewProfilePath
    )

    # Use provided SID if available, otherwise resolve it
    try {
        $newSID = if ($UserSID) { $UserSID } else { Get-LocalUserSID -UserName $UserName }
        Log-Message "Target SID: $newSID"
    }
    catch {
        # Handle special AzureAD error case
        if ($_.Exception.Message -match '^AzureAD_NOT_JOINED:(.+)$') {
            $azUser = $matches[1]
            Log-Message "ERROR: AzureAD user '$azUser' but system is not AzureAD joined"
            throw "Cannot import AzureAD profile - system is not joined to AzureAD/Entra ID. Please join the device first and sign in with the AzureAD account."
        }
        # Re-throw other errors
        throw $_
    }
    $global:ProgressBar.Value = 85

    # Target hive path (files already extracted to ProfileFolder)
    $targetHive = Join-Path $ProfileFolder "NTUSER.DAT"

    if (-not (Test-Path $targetHive)) {
        throw "NTUSER.DAT not found in profile folder: $targetHive"
    }

    # Enable required privileges (critical for ownership/ACL operations)
    Enable-Privilege SeBackupPrivilege -IsCritical $true
    Enable-Privilege SeRestorePrivilege -IsCritical $true
    Enable-Privilege SeTakeOwnershipPrivilege -IsCritical $true

    # =====================================================================
    # CRITICAL: Apply ACLs FIRST, then do SID rewrite
    # This ensures proper ownership and permissions before hive manipulation
    # =====================================================================
    
    # Step 1: Apply standard folder ACLs
    $global:StatusText.Text = "Analyzing profile structure..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Counting files and folders for progress estimation..."
    
    # Count items for progress reporting
    $itemCount = 0
    $folderCount = 0
    $fileCount = 0
    try {
        $items = Get-ChildItem -Path $ProfileFolder -Recurse -Force -ErrorAction SilentlyContinue
        $itemCount = ($items | Measure-Object).Count
        $folderCount = ($items | Where-Object { $_.PSIsContainer } | Measure-Object).Count
        $fileCount = ($items | Where-Object { -not $_.PSIsContainer } | Measure-Object).Count
        Log-Message "Profile contains $itemCount items ($folderCount folders, $fileCount files)"
    }
    catch {
        Log-Message "Could not count profile items: $_"
    }
    
    $global:StatusText.Text = "Applying folder permissions to $folderCount folders and $fileCount files..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Starting recursive ACL application (takeown + icacls)..."
    Log-Message "This process handles: ownership transfer, permission reset, explicit ACL grants"
    
    # Start timer for ACL operations
    $aclStartTime = [DateTime]::Now
    Set-ProfileFolderAcls -ProfilePath $ProfileFolder -UserName $UserName -UserSID $newSID
    $aclElapsed = ([DateTime]::Now - $aclStartTime).TotalSeconds
    Log-Message "Folder ACL application completed in $([Math]::Round($aclElapsed, 1)) seconds"
    $global:ProgressBar.Value = 86

    # Step 2: Apply hive file ACLs using icacls
    $global:StatusText.Text = "Setting NTUSER.DAT ownership and permissions..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Applying NTUSER.DAT file ACLs (taking ownership, setting permissions)..."
    Log-Message "Operations: takeown /A, icacls /reset, icacls /grant, icacls /setowner"
    Set-ProfileHiveAcl -HiveFilePath $targetHive -OwnerSID $newSID -UserName $UserName -UserSID $newSID
    Log-Message "NTUSER.DAT ownership and ACLs applied successfully"
    $global:ProgressBar.Value = 88
    
    # Step 3: Now safe to run SID rewrite (after permissions are set)
    if ($sourceSID -and $sourceSID -ne $newSID) {
        $global:StatusText.Text = "Translating registry hive SID references..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Starting SID translation: $sourceSID -> $newSID"
        Log-Message "This rewrites binary SID references and path strings in NTUSER.DAT"
        
        # Get hive file size for progress context
        try {
            $hiveSize = (Get-Item $targetHive -Force).Length
            $hiveSizeMB = [Math]::Round($hiveSize / 1MB, 2)
            Log-Message "NTUSER.DAT size: $hiveSizeMB MB"
        }
        catch {
            Log-Message "Could not determine hive size"
        }
        
        # Step 3a: Reset UsrClass.dat (Classes Hive) - MOVED HERE TO PREVENT SKIPPING
        # CRITICAL FIX: Do NOT try to patch UsrClass.dat. It contains complex UWP/AppX binary structures
        # that break when modified, causing Explorer.exe crashes (0xc0000409).
        # Strategy: Rename it so Windows generates a fresh, clean one on login.
        # CONDITIONAL: Only force this for users with special characters (who trigger the crash most often).
        # Simple users (letters/numbers) might survive without it, preserving associations.
        
        $classesHive = Join-Path $ProfileFolder "AppData\Local\Microsoft\Windows\UsrClass.dat"
        
        # Derive SOURCE username from OldProfilePath to ensure we check the *original* name for special chars.
        # $UserName here is the TARGET username (e.g., cpeters), which is often simple.
        # We need to know if the SOURCE (e.g., Test.User) was complex.
        $sourceUserCheck = ""
        if ($OldProfilePath -match 'Users\\([^\\]+)$') {
            $sourceUserCheck = $matches[1]
        }
        
        # Check specific condition: Dot or Space in source username (most common crash triggers)
        $needsReset = $false # ($sourceUserCheck -match '[. ]') -or ($OldProfilePath -match '[^a-zA-Z0-9\:\\\-_]')
        Log-Message "Conditional Reset Check: SourceUser='$sourceUserCheck', Path='$OldProfilePath', NeedsReset=$needsReset" 'DEBUG'
        
        if (Test-Path $classesHive) {
            if ($needsReset) {
                $global:StatusText.Text = "Resetting UsrClass.dat (Fixing Explorer Crash)..."
                [System.Windows.Forms.Application]::DoEvents()
                Log-Message "Detected special characters/path complexity. RESETTING UsrClass.dat to prevent crash..."
                
                try {
                    # 1. Take ownership so we can delete it
                    # Note: Permissions already set by recursive Set-ProfileFolderAcls above
                
                    # 2. DELETE the old hive (no backup - complete cleanup)
                    Remove-Item -LiteralPath $classesHive -Force
                    Log-Message "Deleted UsrClass.dat (complete cleanup - no backup)"

                    # 3. Clean up ALL transaction logs (including GUID-formatted ones)
                    # CRITICAL: Must catch UsrClass.dat{GUID}.TM.blf and similar variants
                    $classDir = Split-Path $classesHive
                    
                    # Get ALL files starting with "UsrClass.dat" (no exclusions - we deleted the main file)
                    Get-ChildItem -Path $classDir -File -Force -ErrorAction SilentlyContinue | Where-Object {
                        $_.Name -like "UsrClass.dat*"
                    } | ForEach-Object {
                        Log-Message "  Deleting transaction log/cache: $($_.Name)" 'DEBUG'
                        try { 
                            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop 
                        }
                        catch { 
                            Log-Message "  WARNING: Failed to delete $($_.Name): $_" 'WARN' 
                        }
                    }
                    
                    # 4. Verify complete cleanup
                    if (Test-Path $classesHive) {
                        Log-Message "WARNING: UsrClass.dat still exists after cleanup attempt" 'WARN'
                    }
                    else {
                        Log-Message "UsrClass.dat completely removed. Windows will generate fresh Classes hive on login."
                    }
                }
                catch {
                    Log-Message "WARNING: Failed to reset UsrClass.dat: $_" 'WARN'
                }

                # -------------------------------------------------------------------------
                # CRITICAL FIX 2: Reset Explorer & CloudStore Caches
                # These folders contain 'IconCache.db' and Start Menu binary databases
                # that crash Explorer because they point to old paths.
                # -------------------------------------------------------------------------
                try {
                    $localWin = Join-Path $ProfileFolder "AppData\Local\Microsoft\Windows"
                    
                    # DELETE 'Explorer' folder (Icon Cache) - complete removal
                    $expFolder = Join-Path $localWin "Explorer"
                    if (Test-Path $expFolder) {
                        Remove-Item -LiteralPath $expFolder -Recurse -Force -ErrorAction SilentlyContinue
                        Log-Message "Deleted AppData\Local\...\Explorer (Forces fresh IconCache generation)"
                    }

                    # DELETE 'Shell' folder (Shell cache) - complete removal
                    $shellFolder = Join-Path $localWin "Shell"
                    if (Test-Path $shellFolder) {
                        Remove-Item -LiteralPath $shellFolder -Recurse -Force -ErrorAction SilentlyContinue
                        Log-Message "Deleted AppData\Local\...\Shell (Forces fresh Shell cache)"
                    }

                    # DELETE 'Caches' folder - complete removal
                    $cachesFolder = Join-Path $localWin "Caches"
                    if (Test-Path $cachesFolder) {
                        Remove-Item -LiteralPath $cachesFolder -Recurse -Force -ErrorAction SilentlyContinue
                        Log-Message "Deleted AppData\Local\...\Caches (Forces fresh cache generation)"
                    }

                    # DELETE 'CloudStore' folder (Start Menu / Taskbar Cache) - complete removal
                    $cloudFolder = Join-Path $localWin "CloudStore"
                    if (Test-Path $cloudFolder) {
                        Remove-Item -LiteralPath $cloudFolder -Recurse -Force -ErrorAction SilentlyContinue
                        Log-Message "Deleted AppData\Local\...\CloudStore (Forces fresh Start Menu generation)"
                    }

                    # DELETE StartMenuExperienceHost and ShellExperienceHost packages
                    $packagesPath = Join-Path $ProfileFolder "AppData\Local\Packages"
                    if (Test-Path $packagesPath) {
                        Get-ChildItem -Path $packagesPath -Directory -Force -ErrorAction SilentlyContinue | Where-Object {
                            $_.Name -like "Microsoft.Windows.StartMenuExperienceHost*" -or
                            $_.Name -like "Microsoft.Windows.ShellExperienceHost*"
                        } | ForEach-Object {
                            Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                            Log-Message "Deleted $($_.Name) package cache"
                        }
                    }

                    # DELETE ConnectedDevicesPlatform folder
                    $cdpFolder = Join-Path $ProfileFolder "AppData\Local\ConnectedDevicesPlatform"
                    if (Test-Path $cdpFolder) {
                        Remove-Item -LiteralPath $cdpFolder -Recurse -Force -ErrorAction SilentlyContinue
                        Log-Message "Deleted AppData\Local\ConnectedDevicesPlatform"
                    }

                    # DELETE OneDrive cache folder
                    $oneDriveFolder = Join-Path $ProfileFolder "AppData\Local\Microsoft\OneDrive"
                    if (Test-Path $oneDriveFolder) {
                        Remove-Item -LiteralPath $oneDriveFolder -Recurse -Force -ErrorAction SilentlyContinue
                        Log-Message "Deleted AppData\Local\Microsoft\OneDrive cache"
                    }

                    # DELETE IconCache.db file
                    $iconCacheFile = Join-Path $ProfileFolder "AppData\Local\IconCache.db"
                    if (Test-Path $iconCacheFile) {
                        Remove-Item -LiteralPath $iconCacheFile -Force -ErrorAction SilentlyContinue
                        Log-Message "Deleted AppData\Local\IconCache.db"
                    }

                    # DELETE TileDataLayer folder
                    $tileDataFolder = Join-Path $ProfileFolder "AppData\Local\TileDataLayer"
                    if (Test-Path $tileDataFolder) {
                        Remove-Item -LiteralPath $tileDataFolder -Recurse -Force -ErrorAction SilentlyContinue
                        Log-Message "Deleted AppData\Local\TileDataLayer"
                    }
                }
                catch {
                    Log-Message "WARNING: Failed to delete shell cache folders: $_" 'WARN'
                }
            }
            else {
                Log-Message "Standard user/path detected - Preserving UsrClass.dat (File Associations)" 'INFO'
                # We still clean logs to be safe
                $classDir = Split-Path $classesHive
                $classLogPattern = "UsrClass.dat.LOG*", "UsrClass.dat.blf", "UsrClass.dat*.regtrans-ms"
                Get-ChildItem -Path $classDir -Include $classLogPattern -File -Force -ErrorAction SilentlyContinue | ForEach-Object {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                }
            }
        }
        
        try {
            # -----------------------------------------------------------------
            # CRITICAL: Delete Registry Transaction Logs to force clean load
            # If these exist, Windows might ignore our changes or flag corruption
            # -----------------------------------------------------------------
            Log-Message "Removing NTUSER.DAT transaction logs to force clean hive load..."
            $logPattern = "NTUSER.DAT.LOG*", "NTUSER.DAT.blf", "NTUSER.DAT*.regtrans-ms"
            Get-ChildItem -Path $ProfileFolder -Include $logPattern -File -Force -ErrorAction SilentlyContinue | ForEach-Object {
                Log-Message "  Deleting: $($_.Name)" 'DEBUG'
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            }

            Rewrite-HiveSID -FilePath $targetHive -OldSID $sourceSID -NewSID $newSID -OldProfilePath $OldProfilePath -NewProfilePath $NewProfilePath
            Log-Message "SID translation completed successfully"
        }
        catch {
            Log-Message "SID REWRITE FAILED: $_"
            throw $_
        }


    }
    else {
        Log-Message "No SID translation required (source and target SIDs match)"
    }
 
    $global:ProgressBar.Value = 92
	
    Log-Message "Profile ACLs and SID rewrite completed successfully"
}


function Set-ProfileFolderAcls {
    param(
        [Parameter(Mandatory = $true)][string]$ProfilePath,
        [Parameter(Mandatory = $true)][string]$UserName,
        [string]$UserSID = $null
    )
    Log-Message "Applying folder ACLs to $ProfilePath..."
    try {
        # Step 1: Take ownership (processes all files recursively)
        $global:StatusText.Text = "Taking ownership of all profile files (Step 1/7)..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Step 1: Running icacls setowner Administrators (Recursive)"
        # REPLACED takeown with icacls /setowner for performance and consistency
        # /T = Recursive
        # /L = Operate on link itself (prevents infinite loops on Junctions/Symlinks)
        # /C = Continue on error
        # /Q = Quiet
        icacls "$ProfilePath" /setowner "Administrators" /T /L /C /Q >$null 2>&1
        Log-Message "Ownership transferred to Administrators group"
        
        # Step 2: Grant Administrators temporary control
        $global:StatusText.Text = "Granting Administrators access (Step 2/7)..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Step 2: Granting Administrators full control"
        icacls "$ProfilePath" /grant "Administrators:(F)" /C /Q >$null 2>&1
        
        # Step 3: Reset all ACLs (removes everything, including inherited)
        $global:StatusText.Text = "Resetting existing permissions (Step 3/7)..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Step 3: Resetting all ACLs (removing existing permissions)"
        icacls "$ProfilePath" /reset /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        # Step 4: Remove inheritance protection (so we start clean)
        $global:StatusText.Text = "Removing inheritance protection (Step 4/7)..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Step 4: Removing inheritance protection"
        icacls "$ProfilePath" /inheritance:r /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        # Step 5: Add explicit ACLs for SYSTEM, Administrators, and User
        $global:StatusText.Text = "Granting user and system permissions (Step 5/7)..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Step 5a: Adding explicit ACLs for SYSTEM and Administrators"
        icacls "$ProfilePath" /grant:r "NT AUTHORITY\SYSTEM:(F)" /Q /C >$null 2>&1
        icacls "$ProfilePath" /grant:r "BUILTIN\Administrators:(F)" /Q /C >$null 2>&1
        
        # Step 6: Set ownership to User (Standard for User Profiles)
        Log-Message "Step 6: Setting folder ownership to User (Important for Profile Service)"
        if ($UserSID) {
            icacls "$ProfilePath" /setowner "*$UserSID" /T /C /Q >$null 2>&1
        }
        
        # For domain users, use the full domain\username; for local, use shortname
        $isAzureAD = $false
        if ($UserSID -and (Test-IsAzureADSID -SID $UserSID)) {
            $isAzureAD = $true
        }
        if ($isAzureAD) {
            # Try to resolve SID to NTAccount
            $ntAccount = $null
            try {
                $sidObj = New-Object System.Security.Principal.SecurityIdentifier($UserSID)
                $ntAccount = $sidObj.Translate([System.Security.Principal.NTAccount]).Value
                Log-Message "Step 5b: AzureAD SID $UserSID resolves to NTAccount: $ntAccount"
            }
            catch {
                Log-Message "Step 5b: Could not resolve SID $UserSID to NTAccount, will use SID directly. Error: $_"
            }
            $targetPrincipal = if ($ntAccount) { $ntAccount } else { "*${UserSID}" }
            $grantCmd1 = "icacls `"$ProfilePath`" /grant:r `"*${UserSID}:(F)`" /Q /C /T"
            $grantCmd2 = "icacls `"$ProfilePath`" /grant:r `"*${UserSID}:(OI)(CI)(IO)(F)`" /Q /C"
            Invoke-Expression $grantCmd1 2>&1
            Invoke-Expression $grantCmd2 2>&1
        }
        else {
            # Always use *SID for all user types
            # CRITICAL FIX: Use /T (Recursive) to ensure user has access to ALL existing files, not just root
            Log-Message "Step 5b: Granting ACLs using *SID: $UserSID (Recursive)"
            icacls "$ProfilePath" /grant:r "*${UserSID}:(F)" /T /Q /C >$null 2>&1
            icacls "$ProfilePath" /grant:r "*${UserSID}:(OI)(CI)(IO)(F)" /Q /C >$null 2>&1
        }
        
        Log-Message "Step 5c: Adding ACLs for Administrators group"
        icacls "$ProfilePath" /grant:r "BUILTIN\Administrators:(F)" /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        # Step 6: Add inheritable ACLs for SYSTEM and Administrators
        $global:StatusText.Text = "Setting inheritable permissions for subfolders (Step 6/7)..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Step 6a: Adding inheritable ACLs (OI)(CI)(IO) for SYSTEM and Administrators"
        icacls "$ProfilePath" /grant:r "NT AUTHORITY\SYSTEM:(OI)(CI)(IO)(F)" /Q /C >$null 2>&1
        icacls "$ProfilePath" /grant:r "BUILTIN\Administrators:(OI)(CI)(IO)(F)" /Q /C >$null 2>&1
        # Add CREATOR OWNER for proper per-file ownership on create
        Log-Message "Step 6b: Adding CREATOR OWNER for new file creation"
        icacls "$ProfilePath" /grant:r "CREATOR OWNER:(OI)(CI)(IO)(F)" /Q /C >$null 2>&1
        # Remove overly broad groups that can interfere
        Log-Message "Step 6c: Removing overly permissive groups (Everyone, Users, Authenticated Users)"
        icacls "$ProfilePath" /remove:g "Everyone" "Users" "Authenticated Users" /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        # Step 7: Set owner to the user
        $global:StatusText.Text = "Setting final ownership to user (Step 7/7)..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Step 7: Setting owner to user SID: $UserSID"
        icacls "$ProfilePath" /setowner "*${UserSID}" /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        Log-Message "Folder ACLs applied successfully (all 7 steps completed)"
    }
    catch {
        Log-Message "ERROR: Failed to apply folder ACLs: $_"
        throw $_
    }
}


function Set-ProfileHiveAcl {
    param(
        [Parameter(Mandatory = $true)][string]$HiveFilePath,
        [Parameter(Mandatory = $true)][string]$OwnerSID,
        [Parameter(Mandatory = $true)][string]$UserName,
        [bool]$IsLocalUser = $true,
        [string]$UserSID = $null
    )
    
    # Validate hive file exists
    if (-not (Test-Path $HiveFilePath)) { throw "Hive file not found: $HiveFilePath" }
    
    # CRITICAL FIX: Delete transaction logs to prevent "Temporary Profile" issues
    # Transaction logs (.LOG1, .LOG2, .blf) often contain references to the old SID or are inconsistent
    Log-Message "Cleaning up hive transaction logs to ensure clean load..."
    $hiveDir = Split-Path $HiveFilePath -Parent
    $hiveName = Split-Path $HiveFilePath -Leaf
    $logsToDelete = @("$HiveFilePath.LOG*", "$HiveFilePath.blf", "$HiveFilePath.regtrans-ms")
    
    foreach ($logPattern in $logsToDelete) {
        try {
            Remove-Item -Path $logPattern -Force -ErrorAction SilentlyContinue
            Log-Message "Removed transaction log: $logPattern"
        }
        catch {}
    }
    
    Log-Message "Applying file-level ACLs to NTUSER.DAT..."
    
    try {
        # Step 1: Take ownership as Administrators
        takeown /F "$HiveFilePath" /A /R /D Y >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        # Step 2: Grant Administrators full control temporarily so we can modify ACLs
        icacls "$HiveFilePath" /grant "Administrators:(F)" /C /Q >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        # Step 3: Remove all explicit ACLs
        icacls "$HiveFilePath" /reset /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200
        
        # Step 4: Enable inheritance from parent folder
        icacls "$HiveFilePath" /inheritance:e /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200

        # Step 5: Explicitly grant required principals to avoid temp profile
        # SYSTEM and Administrators
        icacls "$HiveFilePath" /grant:r "NT AUTHORITY\SYSTEM:(F)" /Q /C >$null 2>&1
        icacls "$HiveFilePath" /grant:r "BUILTIN\Administrators:(F)" /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200

        $isAzureAD = $false
        if ($UserSID -and (Test-IsAzureADSID -SID $UserSID)) {
            $isAzureAD = $true
        }
        if ($isAzureAD) {
            # Always use *SID for all user types
            Log-Message "Granting ACLs using *SID: $UserSID"
            icacls "$HiveFilePath" /grant:r "*${UserSID}:(F)" /Q /C >$null 2>&1
        }
        else {
            Log-Message "Granting ACLs using *SID: $UserSID"
            icacls "$HiveFilePath" /grant:r "*${UserSID}:(F)" /Q /C >$null 2>&1
        }
        Start-Sleep -Milliseconds 200

        # Step 6: Set owner to the user to ensure profile loads
        Log-Message "Setting owner to user SID: $UserSID"
        icacls "$HiveFilePath" /setowner "*${UserSID}" /Q /C >$null 2>&1
        Start-Sleep -Milliseconds 200

        Log-Message "NTUSER.DAT ACLs applied and ownership set to user"
    }
    catch {
        Log-Message "ERROR: Failed to apply hive ACLs: $_"
        throw $_
    }
}


function Mount-RegistryHive {
    <#
    .SYNOPSIS
    Loads a registry hive file into HKU
    
    .DESCRIPTION
    Mounts a registry hive with logging and error handling
    
    .PARAMETER HivePath
    Path to the hive file (e.g., NTUSER.DAT)
    
    .PARAMETER MountPoint
    The registry key name under HKU (e.g., "TempHive")
    
    .OUTPUTS
    Returns $true on success, throws on error
    #>
    param(
        [Parameter(Mandatory = $true)][string]$HivePath,
        [Parameter(Mandatory = $true)][string]$MountPoint
    )
    
    Log-Message "HIVE LOAD: Loading hive at HKU\$MountPoint from $HivePath"
    
    try {
        $result = reg load "HKU\$MountPoint" $HivePath 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "reg load failed with exit code ${LASTEXITCODE}: $result"
        }
        Log-Message "HIVE LOAD: Successfully loaded hive at HKU\$MountPoint"
        return $true
    }
    catch {
        Log-Message "ERROR: Failed to load hive: $_"
        throw $_
    }
}


function Dismount-RegistryHive {
    <#
    .SYNOPSIS
    Unloads a registry hive from HKU
    
    .DESCRIPTION
    Unloads a registry hive with retry logic and garbage collection
    
    .PARAMETER MountPoint
    The registry key name under HKU to unload
    
    .PARAMETER MaxRetries
    Maximum number of retry attempts (default: 3)
    
    .OUTPUTS
    Returns $true on success, throws on error after retries
    #>
    param(
        [Parameter(Mandatory = $true)][string]$MountPoint,
        [int]$MaxRetries = 3
    )
    
    Log-Message "HIVE UNLOAD: Unloading hive at HKU\$MountPoint"
    
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            # Force garbage collection before unload
            [gc]::Collect()
            [gc]::WaitForPendingFinalizers()
            
            $result = reg unload "HKU\$MountPoint" 2>&1
            if ($LASTEXITCODE -eq 0) {
                Log-Message "HIVE UNLOAD: Successfully unloaded hive at HKU\$MountPoint"
                return $true
            }
            
            Log-Message "WARNING: Hive unload attempt $attempt failed (exit code ${LASTEXITCODE}): $result"
            
            if ($attempt -lt $MaxRetries) {
                Start-Sleep -Seconds 1
            }
        }
        catch {
            Log-Message "WARNING - Hive unload attempt $attempt threw exception - $_"
            if ($attempt -lt $MaxRetries) {
                Start-Sleep -Seconds 1
            }
        }
    }
    
    throw "Failed to unload hive HKU\$MountPoint after $MaxRetries attempts"
}


function Rewrite-HiveSID {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string]$OldSID,
        [Parameter(Mandatory = $true)][string]$NewSID,
        [Parameter(Mandatory = $false)][string]$OldProfilePath,
        [Parameter(Mandatory = $false)][string]$NewProfilePath
    )
    
    Log-Message "Rewrite-HiveSID called with path: '$FilePath'" 'DEBUG'
    Log-Message "  OldSID: $OldSID" 'DEBUG'
    Log-Message "  NewSID: $NewSID" 'DEBUG'
    Log-Message "  OldProfilePath: '$OldProfilePath'" 'DEBUG'
    Log-Message "  NewProfilePath: '$NewProfilePath'" 'DEBUG'

    # Critical privileges for hive rewrite operations
    Enable-Privilege SeBackupPrivilege -IsCritical $true
    Enable-Privilege SeRestorePrivilege -IsCritical $true
    Enable-Privilege SeTakeOwnershipPrivilege -IsCritical $true

    # Validate input SIDs before attempting rewrite
    # Accept classic SIDs and AzureAD SIDs (S-1-12-1-...)
    if (-not ($OldSID -match '^S-\d+-')) { throw "Invalid source SID format: $OldSID" }
    if (-not ($NewSID -match '^S-\d+-')) { throw "Invalid target SID format: $NewSID" }
    
    $mountPoint = "TempHive_$(Get-Random)"
    
    # Ensure no collision with existing mounts
    $maxAttempts = 5
    $attempts = 0
    while ((reg query "HKU\$mountPoint" 2>$null) -and $attempts -lt $maxAttempts) {
        $mountPoint = "TempHive_$(Get-Random)"
        $attempts++
    }
    if ($attempts -ge $maxAttempts) { throw "Cannot allocate unique hive mount point for SID rewrite" }
    
    $backup = "$FilePath.backup_$(Get-Date -f yyyyMMdd_HHmmss)"

    # === BACKUP (same method as Set-ProfileAcls) ===
    $global:StatusText.Text = "Creating backup of NTUSER.DAT..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Creating hive backup using privileged copy..."
    try {
        Copy-Item -Path $FilePath -Destination $backup -Force -ErrorAction Stop
        Log-Message "Backup created: $backup"
    }
    catch {
        Log-Message "Standard copy failed, falling back to robocopy /B..."
        $srcDir = Split-Path $FilePath -Parent
        $fileName = Split-Path $FilePath -Leaf
        robocopy "$srcDir" "$srcDir" "$fileName" /COPY:DAT /R:1 /W:1 /B /J | Out-Null
        Copy-Item "$srcDir\$fileName" $backup -Force -ErrorAction Stop
        Log-Message "Backup created via robocopy: $backup"
    }

    # Read entire hive
    $global:StatusText.Text = "Reading registry hive into memory..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Loading hive file into memory for binary processing..."
    [byte[]]$data = [IO.File]::ReadAllBytes($FilePath)
    $dataSizeMB = [Math]::Round($data.Length / 1MB, 2)
    Log-Message "Hive loaded: $dataSizeMB MB ($($data.Length) bytes)"

    # Convert SIDs
    $global:StatusText.Text = "Converting SID formats for binary replacement..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Converting text SIDs to binary format..."
    $oldSidObj = New-Object System.Security.Principal.SecurityIdentifier($OldSID)
    $newSidObj = New-Object System.Security.Principal.SecurityIdentifier($NewSID)
    $oldBin = New-Object byte[] $oldSidObj.BinaryLength
    $newBin = New-Object byte[] $newSidObj.BinaryLength
    $oldSidObj.GetBinaryForm($oldBin, 0)
    $newSidObj.GetBinaryForm($newBin, 0)
    Log-Message "Old SID binary: $($oldBin.Length) bytes"
    Log-Message "New SID binary: $($newBin.Length) bytes"

    # v2.10.63 CHANGE: RESTORE BINARY REPLACEMENT (WITH SAFETY)
    # The "Safe Core" (no binary) approach caused Sign-Out loops (ACLs insufficient).
    # The "Crash" (0xc0000409) likely caused by patching a "Dirty" hive (unmerged logs).
    # FIX: Force-Merge the hive logs BEFORE patching to ensure we patch a clean file.
    
    $doBinaryReplacement = $false
    if ($oldBin.Length -eq $newBin.Length) {
        $doBinaryReplacement = $true
    }
    else {
        Log-Message "WARNING: SID binary lengths differ ($($oldBin.Length) vs $($newBin.Length)) - skipping binary replacement to avoid corruption." 'WARN'
    }

    if ($doBinaryReplacement) {
        $global:StatusText.Text = "Preparing hive for binary update..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # 1. PRE-MERGE: Force Windows to merge any pending logs (.LOG1, .blf) into .DAT
        # This prevents "Dirty Hive" corruption when we byte-patch later.
        Log-Message "Performing Pre-Merge maintenance on hive..."
        $tempLoad = "Merge_$(Get-Random)"
        
        try {
            if (Mount-RegistryHive -HivePath $FilePath -MountPoint $tempLoad) {
                Start-Sleep -Milliseconds 500
                Dismount-RegistryHive -MountPoint $tempLoad
                Log-Message "Hive logs merged successfully."
            }
        }
        catch {
            Log-Message "WARNING: Pre-Merge failed: $_" 'WARN'
        }
        
        # 2. READ & PATCH
        $global:StatusText.Text = "Performing binary SID replacement..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # Re-read data (it might have changed during merge)
        $data = [IO.File]::ReadAllBytes($FilePath)
        Log-Message "SID binary lengths match - proceeding with binary replacement on clean hive."
        
        # Byte-level replacement function
        function Replace-Bytes([byte[]]$haystack, [byte[]]$needle, [byte[]]$replacement) {
            $list = [System.Collections.Generic.List[byte]]::new()
            $i = 0
            $replacements = 0
            $lastProgress = [DateTime]::Now
            while ($i -le $haystack.Length - $needle.Length) {
                # Update progress every second
                if (([DateTime]::Now - $lastProgress).TotalSeconds -ge 1) {
                    $percent = [Math]::Round(($i / $haystack.Length) * 100, 1)
                    $global:StatusText.Text = "Binary SID replacement: $percent% ($replacements replacements found)"
                    [System.Windows.Forms.Application]::DoEvents()
                    $lastProgress = [DateTime]::Now
                }
                
                $match = $true
                for ($j = 0; $j -lt $needle.Length; $j++) {
                    if ($haystack[$i + $j] -ne $needle[$j]) { $match = $false; break }
                }
                if ($match) {
                    $list.AddRange($replacement)
                    $i += $needle.Length
                    $replacements++
                }
                else {
                    $list.Add($haystack[$i]); $i++
                }
            }
            while ($i -lt $haystack.Length) { $list.Add($haystack[$i]); $i++ }
            Log-Message "Binary replacement completed: $replacements SID instances replaced"
            return $list.ToArray()
        }

        $data = Replace-Bytes $data $oldBin $newBin
        Log-Message "Binary SID replacement completed ($($oldBin.Length) bytes per replacement)"
    }
    else {
        Log-Message "Skipping Binary Replacement (Conditions not met)"
    }
    
    # Write patched version (only if changed, but here we always write if we entered)
    if ($doBinaryReplacement) {
        $global:StatusText.Text = "Writing patched hive to disk..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Writing patched hive to temporary file..."
        $patched = "$FilePath.patched_$(Get-Random)"
        [IO.File]::WriteAllBytes($patched, $data)
        Log-Message "Patched hive written: $patched"
    
        # Test-load patched hive
        $global:StatusText.Text = "Validating patched hive integrity..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Test-loading patched hive to verify integrity..."
        Log-Message "HIVE LOAD #1: Loading test hive at HKU\$mountPoint from $patched"
        $test = reg load "HKU\$mountPoint" $patched 2>&1
        if ($LASTEXITCODE -ne 0) {
            Log-Message "ERROR: Patched hive failed load test: $test"
            Remove-Item $patched -Force -ErrorAction SilentlyContinue
            throw "SID rewrite failed - patched hive corrupt"
        }
        Log-Message "Hive load test PASSED - patched hive is valid"
        Log-Message "HIVE UNLOAD #1: Unloading test hive at HKU\$mountPoint"
        reg unload "HKU\$mountPoint" | Out-Null
        Log-Message "HIVE UNLOAD #1: Successfully unloaded test hive"
    
        # Replace original with patched
        $global:StatusText.Text = "Replacing original hive with patched version..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Replacing original NTUSER.DAT with patched version..."
        try {
            Move-Item $patched $FilePath -Force -ErrorAction Stop
        }
        catch {
    
            Copy-Item $patched $FilePath -Force -ErrorAction Stop
            Remove-Item $patched -Force
        }
        Log-Message "Binary SID replacement completed successfully"
    }

    # Final string cleanup + OOBE fixes
    $global:StatusText.Text = "Mounting hive for registry string cleanup..."
    [System.Windows.Forms.Application]::DoEvents()
    Log-Message "Performing registry path and string cleanup"
    
    
    # Derive old and new profile paths from SIDs (only if not passed as params)
    
    try {
        # CRITICAL: Load hive INSIDE try block to ensure finally block can unload it
        Log-Message "HIVE LOAD #2: Loading hive for string cleanup at HKU\$mountPoint from $FilePath"
        reg load "HKU\$mountPoint" $FilePath | Out-Null
        Log-Message "HIVE LOAD #2: Successfully loaded hive at HKU\$mountPoint"
        
        if ([string]::IsNullOrWhiteSpace($OldProfilePath)) {
            # Get old profile path from registry (if it exists)
            $profileListBase = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
            $oldSidKey = Get-ItemProperty -Path "$profileListBase\$OldSID" -ErrorAction SilentlyContinue
            if ($oldSidKey) {
                $OldProfilePath = $oldSidKey.ProfileImagePath
                Log-Message "Old profile path from registry: $OldProfilePath"
            }
        }

        if ([string]::IsNullOrWhiteSpace($NewProfilePath)) {
            # Get new profile path
            $profileListBase = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
            $newSidKey = Get-ItemProperty -Path "$profileListBase\$NewSID" -ErrorAction SilentlyContinue
            if ($newSidKey) {
                $NewProfilePath = $newSidKey.ProfileImagePath
                Log-Message "New profile path from registry: $NewProfilePath"
            }
        }
    }
    catch {
        Log-Message "Could not resolve profile paths from SIDs: $_"
    }
    
    # SAFER APPROACH: Use native PowerShell object iteration
    # This avoids "reg export" corruption and handles special characters correctly
    function Update-RegistryStringValues {
        param(
            [string]$KeyPath,
            [string]$OldString,
            [string]$NewString
        )
        
        # Verify key exists
        $key = Get-Item -LiteralPath $KeyPath -ErrorAction SilentlyContinue
        if (-not $key) { return }
        
        try {
            Log-Message "Scanning registry key for path updates: $KeyPath"
            
            # Iterate all values in this key
            foreach ($valueName in $key.GetValueNames()) {
                try {
                    $val = $key.GetValue($valueName, $null, "DoNotExpandEnvironmentNames")
                    
                    # We only care about strings that contain the old path/user
                    if ($val -is [string] -and $val -like "*$OldString*") {
                        
                        # Case-insensitive replacement using escaped regex for safety
                        # This handles "Test.User" correctly (treating dot as literal)
                        $escapedOld = [regex]::Escape($OldString)
                        $newVal = $val -ireplace $escapedOld, $NewString
                        
                        if ($val -ne $newVal) {
                            Set-ItemProperty -LiteralPath $KeyPath -Name $valueName -Value $newVal -Force -ErrorAction Stop
                            Log-Message "Copied registry value: $valueName ($val -> $newVal)" 'DEBUG'
                        }
                    }
                }
                catch {
                    Log-Message "WARNING: Failed to update registry property '$valueName' in '$KeyPath': $_"
                }
            }
        }
        finally {
            # CRITICAL: Explicitly close the registry key handle to prevent hive unload failures
            if ($key) {
                $key.Close()
                $key.Dispose()
            }
        }
    }

    try {
        # Determine strict username for replacement (e.g. Test.User)
        $oldUsername = $null
        if ($oldProfilePath -match 'C:\\Users\\([^\\]+)') {
            $oldUsername = $matches[1]
        }
        
        # Define replacements map (What -> To What)
        $replacements = @{}
        
        if ($oldProfilePath -and $newProfilePath) {
            $replacements[$oldProfilePath] = $newProfilePath
        }
        
        if ($oldUsername) {
            # Map "C:\Users\Test.User" -> "C:\Users\cpeters"
            # This catches variants that might not match exact profile path
            $replacements["C:\Users\$oldUsername"] = $newProfilePath
        }

        # RE-ENABLED: Critical path replacements (SAFE native mode)
        $criticalKeys = @(
            "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
            "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
            "Environment",
            "Software\Microsoft\Windows\CurrentVersion\Run"
        )
        
        Log-Message "Performing targeted registry path cleanup on critical keys..."
        
        # -------------------------------------------------------------------------
        # CRITICAL: Clear Explorer BagMRU/Bags to prevent black screen
        # Standard username - Preserving Shell BagMRU settings
        Log-Message "Standard username '$oldUsername' detected - Preserving Shell BagMRU settings." 'INFO'


        foreach ($subKey in $criticalKeys) {
            # Use specific Registry::HKEY_USERS provider path to avoid drive alias issues
            $fullKey = "Registry::HKEY_USERS\$mountPoint\$subKey"
            
            # Apply all replacements
            foreach ($oldText in $replacements.Keys) {
                # Log the attempt for debugging special character handling
                # matches correctly account for dots due to [regex]::Escape() inside the function
                Update-RegistryStringValues -KeyPath $fullKey -OldString $oldText -NewString $replacements[$oldText]
            }
        }
        
        # -------------------------------------------------------------------------
        # GLOBAL PATH CLEANUP (Deep Clean for Stale Paths)
        # -------------------------------------------------------------------------
        # The critical keys above only cover standard shell folders.
        # Many apps (IE, MediaPlayer, etc.) store hardcoded paths elsewhere.
        # We must perform a recursive sweep of the ENTIRE hive to replace the old profile path.
        if ($oldProfilePath -and $newProfilePath) {
            Log-Message "Skipping GLOBAL recursive registry sweep (SAFE MODE)..." 'INFO'
            # The Global Sweep was causing Explorer crashes (0xc0000409) on clean snapshots.
            # We revert to targeted key replacement only.
            
            # [DISABLED CODE BLOCK]
            # foreach ($oldText in $replacements.Keys) { ... }
        }
        


        # -------------------------------------------------------------------------
        # Final Verification
        # -------------------------------------------------------------------------
        Log-Message "Verifying registry replacements for user '$oldUsername'..." 'DEBUG'
        $testPath = "Registry::HKEY_USERS\$mountPoint\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        $k = Get-Item $testPath -ErrorAction SilentlyContinue

        if ($k) {
            foreach ($vn in $k.GetValueNames()) {
                $vv = $k.GetValue($vn, $null, 'DoNotExpandEnvironmentNames')
                if ($vv -like "*$oldUsername*") {
                    Log-Message "ERROR: Found leftover old username reference in '$vn': $vv" 'ERROR'
                }
            }
            # CRITICAL: Dispose registry key handle
            $k.Close()
            $k.Dispose()
        }
        else {
            Log-Message "WARNING: Could not verify cleanup - 'User Shell Folders' key not found." 'WARN'
        }
        
        # Diagnostic: summarize rewrite status across critical keys
        function Report-RewriteSummary {
            param(
                [string]$Root,
                [string]$OldSid,
                [string]$NewSid,
                [string]$OldPath,
                [string]$NewPath
            )
            Log-Message "Summarizing SID/path rewrite status"
            $targets = @(
                "Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband",
                "Software\Microsoft\Windows\CurrentVersion\CloudStore",
                "Software\Microsoft\Windows\Shell\Bags",
                "Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\BagMRU",
                "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
                "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
            )
            foreach ($t in $targets) {
                $base = "Registry::$Root\\$t"
                $key = Get-Item $base -ErrorAction SilentlyContinue
                if (-not $key) { continue }
                $oldSidCount = 0; $newSidCount = 0; $oldPathCount = 0; $newPathCount = 0
                try {
                    function Count-Key([string]$p) {
                        $k = Get-Item $p -ErrorAction SilentlyContinue
                        if (-not $k) { return }
                        foreach ($vn in $k.GetValueNames()) {
                            try {
                                $vv = $k.GetValue($vn, $null, 'DoNotExpandEnvironmentNames')
                                if ($vv -is [string]) {
                                    if ($OldSid -and $vv -like "*$OldSid*") { $script:oldSidCount++ }
                                    if ($NewSid -and $vv -like "*$NewSid*") { $script:newSidCount++ }
                                    if ($OldPath -and $vv -like "*$OldPath*") { $script:oldPathCount++ }
                                    if ($NewPath -and $vv -like "*$NewPath*") { $script:newPathCount++ }
                                }
                            }
                            catch {}
                        }
                        foreach ($sk in $k.GetSubKeyNames()) { Count-Key "$p\\$sk" }
                    }
                    Count-Key $base
                    Log-Message ("Rewrite summary for '$t': OldSID=$oldSidCount, NewSID=$newSidCount, OldPath=$oldPathCount, NewPath=$newPathCount")
                }
                catch {
                    Log-Message "Summary scan error for ${t}: $_"
                }
            }
        }
        
        # CRITICAL: Force GC after registry string cleanup to close Get-Item handles
        # The Update-RegistryStringValues function and verification code use Get-Item
        # which opens registry key handles that must be closed before Report-RewriteSummary
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        Log-Message "Garbage collection after registry string cleanup"
        
        # CRITICAL: Force garbage collection before summary to minimize open handles
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Report-RewriteSummary -Root "HKU:$mountPoint" -OldSid $OldSID -NewSid $NewSID -OldPath $oldProfilePath -NewPath $newProfilePath
        
        # CRITICAL: Force aggressive garbage collection after summary to close registry handles
        # This prevents "Access is denied" errors when trying to unload the hive
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()  # Second collection to clean up finalizers
        Log-Message "Forced garbage collection to release registry handles"

        # Disable telemetry/OOBE junk - using Registry provider (avoids PSDrive issues)
        Log-Message "Disabling telemetry/OOBE via Registry provider"
        $oobeKeyPath = "Registry::HKEY_USERS\$mountPoint\Software\Microsoft\Windows\CurrentVersion\OOBE"
        
        try {
            # Ensure the OOBE key exists
            if (-not (Test-Path $oobeKeyPath)) {
                New-Item -Path $oobeKeyPath -Force | Out-Null
            }
            
            # Set OOBE values using Registry provider (no external processes, no PSDrive needed)
            New-ItemProperty -Path $oobeKeyPath -Name "SkipMachineOOBE" -Value 1 -PropertyType DWord -Force -ErrorAction SilentlyContinue | Out-Null
            New-ItemProperty -Path $oobeKeyPath -Name "DisablePrivacyExperience" -Value 1 -PropertyType DWord -Force -ErrorAction SilentlyContinue | Out-Null
            Log-Message "OOBE settings applied successfully"
        }
        catch {
            Log-Message "WARNING: Failed to set OOBE settings: $_" 'WARN'
        }
        
        # CRITICAL: Force GC to close any PowerShell registry handles
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        Log-Message "Final garbage collection before hive unload"
    }
    finally {
        # GUARANTEED cleanup: unload hive even if Fix-Key fails
        Log-Message "HIVE UNLOAD #2: Preparing to unload hive '$mountPoint' (from HIVE LOAD #2) - forcing handle cleanup..."
        
        # AGGRESSIVE: Multiple GC rounds with delays to force handle release
        for ($i = 1; $i -le 3; $i++) {
            [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::WaitForFullGCComplete() | Out-Null
            Start-Sleep -Milliseconds 100
        }
        Log-Message "Aggressive garbage collection completed with delays"
            
        $unloadAttempts = 0
        $unloadSuccess = $false
        do {
            $unloadAttempts++
            Log-Message "HIVE UNLOAD #2: Attempt $unloadAttempts of 5..." 'DEBUG'
            
            $result = reg unload "HKU\$mountPoint" 2>&1
            
            if ($LASTEXITCODE -eq 0) {
                Log-Message "HIVE UNLOAD #2: Successfully unloaded hive at HKU\$mountPoint"
                $global:StatusText.Text = "Registry cleanup complete"
                [System.Windows.Forms.Application]::DoEvents()
                $unloadSuccess = $true
                break
            }
            else {
                Log-Message "HIVE UNLOAD #2: Attempt $unloadAttempts failed (exit code: $LASTEXITCODE): $result" 'WARN'
                
                if ($unloadAttempts -ge 5) {
                    Log-Message "ERROR: Failed to unload hive HKU\$mountPoint after 5 attempts. Manual cleanup required!" 'ERROR'
                    Log-Message "ERROR: Run this command manually: reg unload `"HKU\$mountPoint`"" 'ERROR'
                    break
                }
                
                # Wait before retry
                Start-Sleep -Milliseconds $Config.HiveUnloadWaitMs
            }
        } while ($true)
        
        if (-not $unloadSuccess) {
            # Last-ditch effort: try to force close handles
            Log-Message "Attempting emergency handle cleanup..." 'WARN'
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            Start-Sleep -Seconds 2
            
            $result = reg unload "HKU\$mountPoint" 2>&1
            if ($LASTEXITCODE -eq 0) {
                Log-Message "Emergency unload succeeded!" 'WARN'
            }
            else {
                Log-Message "CRITICAL: Emergency unload also failed. Hive HKU\$mountPoint remains mounted!" 'ERROR'
            }
        }
    }

    Log-Message "Rewrite-HiveSID completed successfully: $OldSID -> $NewSID"
}


function Test-AndFix-ProfileHive {
    param(
        [string]$HivePath,
        [string]$UserId,
        [string]$UserName
    )

    $testMount = "HKU\ProfileCheck_$(Get-Random)"
    try {
        Log-Message "Attempting to load hive for verification: $HivePath"
        $out = reg load "$testMount" "$HivePath" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Log-Message "Initial hive load failed: $out"
            Log-Message "Attempting to repair NTUSER.DAT ownership and ACLs..."

            # Take ownership as Administrators and grant full control, then set to user
            try {
                takeown /F "$HivePath" /A /R /D Y >$null 2>&1
                icacls "$HivePath" /grant "Administrators:(F)" /C /Q >$null 2>&1
                icacls "$HivePath" /setowner *"$UserName" /Q >$null 2>&1
                icacls "$HivePath" /grant:r *"${UserName}:(F)" /Q >$null 2>&1
                Start-Sleep -Milliseconds 500
            }
            catch {
                Log-Message "Failed to adjust NTUSER.DAT ACLs: $_"
            }

            # Retry load
            $out2 = reg load "$testMount" "$HivePath" 2>&1
            if ($LASTEXITCODE -ne 0) {
                Log-Message "Retry hive load failed: $out2"
                return $false
            }
        }

        # Successfully loaded - unload and return success
        reg unload "$testMount" | Out-Null
        Log-Message "Hive verification succeeded"
        return $true
    }
    catch {
        Log-Message "Exception during hive verification: $_"
        try { reg unload "$testMount" | Out-Null } catch {}
        return $false
    }
}


function Remove-FolderRobust {
    param([string]$Path)
    if (Test-Path $Path) {
        try { Remove-Item $Path -Recurse -Force -ErrorAction Stop }
        catch {
            $empty = "$env:TEMP\Empty_$(Get-Random)"
            New-Item -ItemType Directory -Path $empty -Force | Out-Null
            robocopy $empty $Path /MIR /R:1 /W:1 | Out-Null
            Remove-Item $Path -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item $empty -Recurse -Force
        }
    }
}


function Get-FolderSize {
    <#
    .SYNOPSIS
    Calculates the total size of a folder in bytes
    
    .DESCRIPTION
    Recursively calculates folder size with error handling
    
    .PARAMETER Path
    The folder path to measure
    
    .OUTPUTS
    Returns size in bytes, or 0 if path doesn't exist or on error
    #>
    param([Parameter(Mandatory = $true)][string]$Path)
    
    if (-not (Test-Path $Path)) { return 0 }
    
    try {
        $size = (Get-ChildItem -Path $Path -Recurse -File -Force -ErrorAction SilentlyContinue | 
            Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        
        if ($size) { return $size } else { return 0 }
    }
    catch {
        Log-Message "WARNING: Size calculation failed for '$Path': $_"
        return 0
    }
}


function Get-RobocopyExclusions {
    param(
        [ValidateSet('Export', 'Import', 'AppData')]
        [string]$Mode = 'Export'
    )
    
    $excludeFiles = @('Thumbs.db', 'desktop.ini')
    $excludeDirs = @('AppData\Local\Microsoft\Edge\User Data\ShaderCache', 'AppData\Local\Temp')
    
    $result = @{
        Files = $excludeFiles
        Dirs  = $excludeDirs
    }
    
    # Export mode: exclude everything in AppData except what we need
    if ($Mode -eq 'Export') {
        # Only include specific AppData subdirectories (everything else in AppData is excluded)
        # This list is complementary - we EXCLUDE everything EXCEPT these
        $result.AppDataInclude = @(
            'Roaming\Microsoft\Signatures',
            'Roaming\Microsoft\Outlook',
            'Roaming\Microsoft\Office',
            'Roaming\Microsoft\Credentials',
            'Local\Microsoft\Credentials',
            'Local\Microsoft\Protect',
            'Local\Packages\Microsoft.MicrosoftStickyNotes_*',
            'Local\Google\Chrome\User Data\Default',
            'Local\Google\Chrome\User Data\Profile *',
            'Local\Microsoft\Edge\User Data\Default',
            'Local\Microsoft\Edge\User Data\Profile *',
            'Roaming\Mozilla\Firefox\Profiles\*'
        )
    }
    
    # Import mode excludes only LOG files and temp data
    # NTUSER.DAT must be extracted for 7-Zip import (it's extracted, then rewritten with SID replacement)
    if ($Mode -eq 'Import') {
        # Exclude specific LOG files but NOT NTUSER.DAT itself
        $result.Files += @('ntuser.dat.LOG1', 'ntuser.dat.LOG2', 'NTUSER.DAT.LOG1', 'NTUSER.DAT.LOG2', 'UsrClass.dat.LOG1', 'UsrClass.dat.LOG2')
        # Import should only include specific AppData paths, exclude everything else
        # We'll handle this with explicit exclusion in the import logic
        $result.AppDataInclude = @(
            'Roaming\Microsoft\Signatures',
            'Roaming\Microsoft\Outlook',
            'Roaming\Microsoft\Office',
            'Roaming\Microsoft\Credentials',
            'Local\Microsoft\Credentials',
            'Local\Microsoft\Protect',
            'Local\Packages\Microsoft.MicrosoftStickyNotes_*',
            'Local\Google\Chrome\User Data\Default',
            'Local\Google\Chrome\User Data\Profile *',
            'Local\Microsoft\Edge\User Data\Default',
            'Local\Microsoft\Edge\User Data\Profile *',
            'Roaming\Mozilla\Firefox\Profiles\*'
        )
    }
    
    # AppData bonus copy only needs cache exclusions
    if ($Mode -eq 'AppData') {
        $result.Files = @()
        $result.Dirs = @('Cache', 'Code Cache', 'GPUCache', 'Temp')
    }
    
    return $result
}


function Enable-Privilege {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [bool]$IsCritical = $false
    )
    if (-not $Name) { throw "Privilege name cannot be empty" }
    $ok = [Priv]::EnablePrivilege($Name)
    if (-not $ok) {
        $msg = "Failed to enable privilege: $Name"
        Write-Host "WARNING: $msg" -ForegroundColor Yellow
        Log-Warning $msg
        if ($IsCritical) {
            throw "CRITICAL: Cannot enable required privilege $Name - import cannot proceed"
        }
    }
    return $ok
}


function New-CleanupItem {
    <#
    .SYNOPSIS
    Creates a cleanup item hashtable for the cleanup wizard
    
    .DESCRIPTION
    Standardizes cleanup item creation with all required properties
    
    .OUTPUTS
    Returns properly formatted cleanup item hashtable
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Category,
        [Parameter(Mandatory = $true)][string]$Description,
        [Parameter(Mandatory = $true)][long]$Size,
        [Parameter(Mandatory = $false)][array]$Paths = @(),
        [bool]$DefaultChecked = $false,
        [array]$Details = @(),
        [bool]$HasIndividualSelection = $false,
        [array]$IndividualItems = @(),
        [array]$DuplicateGroups = @()
    )
    
    $item = @{
        Category    = $Category
        Description = $Description
        Size        = $Size
        Paths       = $Paths
        Checked     = $DefaultChecked
    }
    
    if ($Details.Count -gt 0) {
        $item.Details = $Details
    }
    
    if ($HasIndividualSelection) {
        $item.HasIndividualSelection = $true
        $item.IndividualItems = $IndividualItems
        $item.SelectedIndividualPaths = @()
        
        # Extended support for Duplicate Files logic
        if ($DuplicateGroups) {
            $item.HasDuplicateGroups = $true
            $item.DuplicateGroups = $DuplicateGroups
        }
    }
    
    return $item
}


function Get-FriendlyAppName {
    param(
        [string]$PackageId
    )
    
    if (-not $PackageId) {
        return "Unknown"
    }
    
    # Map common package IDs to friendly names
    $nameMap = @{
        '7zip.7zip'                         = '7-Zip'
        'Microsoft.msodbcsql.17'            = 'Microsoft ODBC Driver 17 for SQL Server'
        'Google.Chrome'                     = 'Google Chrome'
        'GlavSoft.TightVNC'                 = 'TightVNC'
        'HARMAN.AdobeAIR'                   = 'Adobe AIR'
        'Microsoft.Edge'                    = 'Microsoft Edge'
        'Cisco.UmbrellaRoamingClient'       = 'Cisco Umbrella Roaming Client'
        'Oracle.JavaRuntimeEnvironment'     = 'Java Runtime Environment'
        'ForensiT.Transwiz'                 = 'Transwiz'
        'Microsoft.VCRedist.2015+.x86'      = 'Visual C++ Redistributable 2015+ (x86)'
        'Microsoft.VCRedist.2015+.x64'      = 'Visual C++ Redistributable 2015+ (x64)'
        'Microsoft.Teams.Classic'           = 'Microsoft Teams Classic'
        'Microsoft.PowerBI'                 = 'Power BI Desktop'
        'Microsoft.DotNet.DesktopRuntime.8' = '.NET Desktop Runtime 8'
        'Microsoft.OneDrive'                = 'OneDrive'
        'Microsoft.AppInstaller'            = 'App Installer'
        'Microsoft.UI.Xaml.2.7'             = 'Windows UI Library 2.7'
        'Microsoft.UI.Xaml.2.8'             = 'Windows UI Library 2.8'
        'Microsoft.VCLibs.Desktop.14'       = 'Visual C++ Libraries Desktop 14'
    }
    
    # Check if we have a friendly name mapping
    if ($nameMap.ContainsKey($PackageId)) {
        return $nameMap[$PackageId]
    }
    
    # Fallback: try to extract a reasonable name from the package ID
    # Format is usually Publisher.ProductName or Publisher.ProductName.Version
    $parts = $PackageId -split '\.'
    
    if ($parts.Count -ge 2) {
        # Take the second part (product name) and make it more readable
        $productName = $parts[1]
        
        # Handle special cases
        if ($productName -match '^[a-z]+$') {
            # All lowercase - capitalize first letter
            return $productName.Substring(0, 1).ToUpper() + $productName.Substring(1)
        }
        elseif ($productName -match '^[A-Z]+$') {
            # All uppercase - keep as is (likely an acronym)
            return $productName
        }
        else {
            # Mixed case - likely already formatted correctly
            return $productName
        }
    }
    
    # Last resort: return the full package ID
    return $PackageId
}


function Get-DomainFQDN {
    param([Parameter(Mandatory = $true)][string]$NetBIOSName)
    try {
        # Try getting current domain if machine is domain-joined
        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        if ($domain.Name -match "^$NetBIOSName\.") { return $domain.Name }
    }
    catch {}
    
    # Fallback: DNS lookup for common TLDs
    $commonTLDs = @('.local', '.lan', '.corp', '.com', '.net', '.org')
    foreach ($tld in $commonTLDs) {
        $fqdn = "$NetBIOSName$tld"
        try {
            $result = Resolve-DnsName -Name $fqdn -ErrorAction Stop -QuickTimeout
            if ($result) { return $fqdn }
        }
        catch {}
    }
    
    # If all else fails, return original
    return $NetBIOSName
}


function Test-DomainReachability {
    param([string]$DomainName)
    try {
        Log-Message "Testing domain reachability: $DomainName"
        
        # DNS with timeout
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $dnsResult = $null
        try {
            $dnsResult = Resolve-DnsName -Name $DomainName -ErrorAction Stop -QuickTimeout
        }
        catch {
            if ($sw.ElapsedMilliseconds -gt $Config.DomainReachabilityTimeout) {
                return @{ Success = $false; Error = "DNS resolution timeout for '$DomainName'"; ErrorCode = "DNS_TIMEOUT" }
            }
            return @{ Success = $false; Error = "DNS resolution failed for domain '$DomainName'"; ErrorCode = "DNS_FAIL" }
        }
        
        if (-not $dnsResult) {
            return @{ Success = $false; Error = "DNS resolution failed for domain '$DomainName'"; ErrorCode = "DNS_FAIL" }
        }
        Log-Message "DNS resolution successful"
        
        # LDAP connectivity with proper timeout
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        try {
            $connect = $tcpClient.BeginConnect($DomainName, 389, $null, $null)
            $wait = $connect.AsyncWaitHandle.WaitOne($Config.DomainReachabilityTimeout, $false)
            if (-not $wait) {
                $tcpClient.Close()
                return @{ Success = $false; Error = "Cannot reach domain controller on port 389 (LDAP). Check network connectivity."; ErrorCode = "LDAP_UNREACHABLE" }
            }
            $tcpClient.EndConnect($connect)
            $tcpClient.Close()
            Log-Message "Domain controller reachable on LDAP port"
            return @{ Success = $true; Error = $null; ErrorCode = $null }
        }
        finally {
            $tcpClient.Dispose()
        }
    }
    catch {
        return @{ Success = $false; Error = "Domain reachability test failed: $($_.Exception.Message)"; ErrorCode = "NETWORK_ERROR" }
    }
}


function Get-DomainCredential {
    param([string]$Domain)
    
    # Theme support
    $currentTheme = if ($global:CurrentTheme) { $global:CurrentTheme } else { "Light" }
    $theme = $Themes[$currentTheme]
    
    # Show custom themed credential dialog similar to line 5599
    $credForm = New-Object System.Windows.Forms.Form
    $credForm.Text = "Domain Credentials"
    $credForm.Size = New-Object System.Drawing.Size(500, 320)
    $credForm.StartPosition = "CenterScreen"
    $credForm.FormBorderStyle = "FixedDialog"
    $credForm.MaximizeBox = $false
    $credForm.MinimizeBox = $false
    $credForm.TopMost = $true
    $credForm.BackColor = $theme.FormBackColor
    $credForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Header panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size = New-Object System.Drawing.Size(500, 70)
    $headerPanel.BackColor = $theme.HeaderBackColor
    $credForm.Controls.Add($headerPanel)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(460, 25)
    $lblTitle.Text = "Domain Credentials Required"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $theme.HeaderTextColor
    $headerPanel.Controls.Add($lblTitle)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = New-Object System.Drawing.Point(22, 43)
    $lblInfo.Size = New-Object System.Drawing.Size(460, 20)
    $lblInfo.Text = "Enter credentials to join domain: $Domain"
    $lblInfo.ForeColor = $theme.SubHeaderTextColor
    $headerPanel.Controls.Add($lblInfo)

    # Main content card
    $contentCard = New-Object System.Windows.Forms.Panel
    $contentCard.Location = New-Object System.Drawing.Point(15, 85)
    $contentCard.Size = New-Object System.Drawing.Size(460, 140)
    $contentCard.BackColor = $theme.PanelBackColor
    $credForm.Controls.Add($contentCard)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Location = New-Object System.Drawing.Point(20, 25)
    $lblUser.Size = New-Object System.Drawing.Size(100, 20)
    $lblUser.Text = "Username:"
    $lblUser.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblUser.ForeColor = $theme.LabelTextColor
    $contentCard.Controls.Add($lblUser)

    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = New-Object System.Drawing.Point(130, 23)
    $txtUser.Size = New-Object System.Drawing.Size(300, 25)
    $txtUser.Text = "$Domain\"
    $txtUser.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtUser.BackColor = $theme.TextBoxBackColor
    $txtUser.ForeColor = $theme.TextBoxForeColor
    $contentCard.Controls.Add($txtUser)

    $lblPass = New-Object System.Windows.Forms.Label
    $lblPass.Location = New-Object System.Drawing.Point(20, 65)
    $lblPass.Size = New-Object System.Drawing.Size(100, 20)
    $lblPass.Text = "Password:"
    $lblPass.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblPass.ForeColor = $theme.LabelTextColor
    $contentCard.Controls.Add($lblPass)

    $txtPass = New-Object System.Windows.Forms.TextBox
    $txtPass.Location = New-Object System.Drawing.Point(130, 63)
    $txtPass.Size = New-Object System.Drawing.Size(300, 25)
    $txtPass.UseSystemPasswordChar = $true
    $txtPass.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtPass.BackColor = $theme.TextBoxBackColor
    $txtPass.ForeColor = $theme.TextBoxForeColor
    $contentCard.Controls.Add($txtPass)

    $chkShowPass = New-Object System.Windows.Forms.CheckBox
    $chkShowPass.Location = New-Object System.Drawing.Point(130, 100)
    $chkShowPass.Size = New-Object System.Drawing.Size(150, 23)
    $chkShowPass.Text = "Show password"
    $chkShowPass.ForeColor = $theme.LabelTextColor
    $chkShowPass.Add_CheckedChanged({ $txtPass.UseSystemPasswordChar = -not $chkShowPass.Checked })
    $contentCard.Controls.Add($chkShowPass)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(250, 240)
    $btnOK.Size = New-Object System.Drawing.Size(110, 35)
    $btnOK.Text = "Connect"
    $btnOK.BackColor = $theme.ButtonPrimaryBackColor
    $btnOK.ForeColor = $theme.ButtonPrimaryForeColor
    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOK.FlatAppearance.BorderSize = 0
    $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOK.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnOK.Add_MouseEnter({ $this.BackColor = $theme.ButtonPrimaryHoverBackColor })
    $btnOK.Add_MouseLeave({ $this.BackColor = $theme.ButtonPrimaryBackColor })
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $credForm.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(370, 240)
    $btnCancel.Size = New-Object System.Drawing.Size(105, 35)
    $btnCancel.Text = "Cancel"
    $btnCancel.BackColor = $theme.ButtonSecondaryBackColor
    $btnCancel.ForeColor = $theme.ButtonSecondaryForeColor
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.FlatAppearance.BorderSize = 0
    $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCancel.Add_MouseEnter({ $this.BackColor = $theme.ButtonSecondaryHoverBackColor })
    $btnCancel.Add_MouseLeave({ $this.BackColor = $theme.ButtonSecondaryBackColor })
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $credForm.Controls.Add($btnCancel)

    $credForm.AcceptButton = $btnOK
    $credForm.CancelButton = $btnCancel
    $txtUser.Select($txtUser.Text.Length, 0)
    
    if ($credForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $u = $txtUser.Text.Trim()
        if ($u -notmatch '\\') { $u = "$Domain\$u" }
        $p = $txtPass.Text | ConvertTo-SecureString -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential($u, $p)
        $credForm.Dispose()
        return $cred
    }
    $credForm.Dispose()
    return $null
}


function Get-DomainAdminCredential {
    param(
        [string]$DomainName,
        [System.Management.Automation.PSCredential]$InitialCredential
    )
    
    $cred = $InitialCredential
    $retryCount = 0
    
    while ($true) {
        # 1. Get Credentials if not provided
        if (-not $cred) {
            # Use custom theme-aware credential dialog for consistency across app
            $cred = Get-DomainCredential -Domain $DomainName
            if (-not $cred) { return $null } # User cancelled
        }
        
        $global:StatusText.Text = "Validating credentials..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # 2. Basic Validation (LDAP Bind)
        $bindTest = Test-DomainCredentials -DomainName $DomainName -Credential $cred
        if (-not $bindTest.Success) {
            Log-Message "Credential validation failed: $($bindTest.Error)"
            
            # Special handling for network/reachability errors
            if ($bindTest.ErrorCode -eq "DC_NOT_OPERATIONAL" -or $bindTest.ErrorCode -eq "NETWORK_ERROR" -or $bindTest.ErrorCode -eq "LDAP_UNREACHABLE") {
                $res = Show-ModernDialog -Message "Network Connection Failed:`r`n`r`n$($bindTest.Error)`r`n`r`nRetry connection?" -Title "Connection Error" -Type Error -Buttons YesNo
                if ($res -eq "Yes") {
                    $retryCount++
                    # Do NOT clear $cred, strictly retry the connection check
                    continue
                }
                throw "Network Error: $($bindTest.Error)"
            }

            $res = Show-ModernDialog -Message "Authentication Failed:`r`n`r`n$($bindTest.Error)`r`n`r`nRetry with different credentials?" -Title "Invalid Credentials" -Type Warning -Buttons YesNo
            if ($res -eq "Yes") {
                $cred = $null
                $retryCount++
                continue
            }
            throw "Authentication failed: $($bindTest.Error)"
        }
        
        # 3. Permission Validation (Domain Admin Check)
        try {
            Log-Message "Checking domain join permissions..."
            Add-Type -AssemblyName System.DirectoryServices.AccountManagement
            $context = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain, $DomainName, $cred.UserName, $cred.GetNetworkCredential().Password)
            $user = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($context, $cred.UserName)
            
            if (-not $user) { throw "User account '$($cred.UserName)' not found in domain" }
            
            $groups = $user.GetAuthorizationGroups()
            $isDomainAdmin = $false
            foreach ($group in $groups) {
                if ($group.Name -eq "Domain Admins" -or $group.Name -eq "Administrators") {
                    $isDomainAdmin = $true
                    break
                }
            }
            
            if (-not $isDomainAdmin) {
                # WARNING: Not a Domain Admin
                Log-Message "WARNING: User is not in Domain Admins group"
                $res = Show-ModernDialog -Message "User is not in Domain Admins group. Domain join may fail if delegated permissions are not configured.`n`nContinue with these credentials anyway?" -Title "Permission Warning" -Type Warning -Buttons YesNo
                if ($res -eq "No") {
                    # User wants to try different credentials
                    $res2 = Show-ModernDialog -Message "Do you want to try entering different credentials?" -Title "Retry?" -Type Question -Buttons YesNo
                    if ($res2 -eq "Yes") {
                        $cred = $null
                        $retryCount++
                        continue
                    }
                    return $null # Cancelled completely
                }
            }
        }
        catch {
            $exMsg = $_.Exception.Message
            if ($_.Exception.InnerException) {
                $exMsg = $_.Exception.InnerException.Message
            }
            # Robust cleanup
            if ($exMsg -like '*: "*') {
                $parts = $exMsg -split ': "'
                if ($parts.Count -gt 1) { $exMsg = $parts[$parts.Count - 1].Trim('"') }
            }
            $exMsg = $exMsg.Trim()
            Log-Message "Permission validation failed: $exMsg"
            
            # Check for invalid credentials during permission check (Retry Logic)
            if ($exMsg -match "user\s*name|password|logon failure|invalid credentials|unknown user") {
                $res = Show-ModernDialog -Message "Permission Check Failed:`r`n`r`n$exMsg`r`n`r`nRetry with different credentials?" -Title "Authentication Failed" -Type Warning -Buttons YesNo
                if ($res -eq "Yes") {
                    $cred = $null
                    $retryCount++
                    continue
                }
            }
            # Check for network/timeout during permission check
            elseif ($exMsg -match "server.*not.*operational|timeout|network path was not found") {
                $res = Show-ModernDialog -Message "Connection Lost:`r`n`r`n$exMsg`r`n`r`nRetry connection?" -Title "Connection Error" -Type Error -Buttons YesNo
                if ($res -eq "Yes") {
                    $retryCount++
                    continue
                }
            }
            
            throw "Permission check failed: $exMsg"
        }
        
        # If we got here, credentials are valid (or user accepted warning)
        return $cred
    }
}


function Test-DomainCredentials {
    param([string]$DomainName, [System.Management.Automation.PSCredential]$Credential)
    try {
        Log-Message "Validating domain credentials..."
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $username = $Credential.UserName
        if ($username -match '\\') { $username = ($username -split '\\')[1] }
        Log-Message "Testing credentials for user: $username in domain: $DomainName"
        try {
            $contextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
            $principalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext($contextType, $DomainName, $username, $Credential.GetNetworkCredential().Password)
            $userPrincipal = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($principalContext, $username)
            if ($userPrincipal) {
                Log-Message "Credentials validated successfully (User: $($userPrincipal.SamAccountName), Display: $($userPrincipal.DisplayName))"
                $principalContext.Dispose()
                return @{ Success = $true; Error = $null; ErrorCode = $null }
            }
            else {
                Log-Message "WARNING: Context created but user not found - credentials likely valid"
                $principalContext.Dispose()
                return @{ Success = $true; Warning = "Could not verify user account exists, but credentials appear valid for domain"; ErrorCode = $null }
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            if ($errorMsg -match "username or password|logon failure|invalid credentials|unknown user|0x8007052e|0x52e") {
                Log-Message "Credential validation failed: Invalid username or password"
                return @{ Success = $false; Error = "Invalid credentials for domain '$DomainName' (username or password incorrect)"; ErrorCode = "INVALID_CREDENTIALS" }
            }
            elseif ($errorMsg -match "account.*locked|account.*disabled|0x80070533|0x533") {
                Log-Message "Credential validation failed: Account locked or disabled"
                return @{ Success = $false; Error = "User account is locked or disabled in domain '$DomainName'"; ErrorCode = "ACCOUNT_LOCKED_OR_DISABLED" }
            }
            elseif ($errorMsg -match "server.*not.*operational|0x8007203a|0x203a") {
                Log-Message "Domain controller not available"
                return @{ Success = $false; Error = "Domain controller not operational or unreachable"; ErrorCode = "DC_NOT_OPERATIONAL" }
            }
            else {
                Log-Message "WARNING: Credential check inconclusive: $errorMsg"
                return @{ Success = $true; Warning = "Could not fully validate credentials (network or configuration issue), but will attempt domain join anyway. Error: $errorMsg"; ErrorCode = $null }
            }
        }
    }
    catch {
        Log-Message "WARNING: Credential validation system failed: $($_.Exception.Message)"
        return @{ Success = $true; Warning = "Credential validation unavailable (will proceed with domain join attempt): $($_.Exception.Message)"; ErrorCode = $null }
    }
}


function Get-DomainJoinErrorDetails {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    $errorMessage = $ErrorRecord.Exception.Message
    $errorCode = "UNKNOWN"
    $userFriendlyMessage = $errorMessage
    $suggestion = ""
    if ($errorMessage -match "0x0000052E") {
        $errorCode = "LOGON_FAILURE"
        $userFriendlyMessage = "Invalid username or password"
        $suggestion = "Verify credentials and try again. Ensure you're using DOMAIN\username format."
    }
    elseif ($errorMessage -match "0x00000035") {
        $errorCode = "NETWORK_PATH_NOT_FOUND"
        $userFriendlyMessage = "Cannot find the domain controller"
        $suggestion = "Check network connectivity and DNS settings. Ensure the domain name is correct."
    }
    elseif ($errorMessage -match "0x0000054B") {
        $errorCode = "COMPUTER_ALREADY_EXISTS"
        $userFriendlyMessage = "A computer account with this name already exists in the domain"
        $suggestion = "Either use a different computer name, or have a domain admin delete the existing computer account."
    }
    elseif ($errorMessage -match "0x00000569") {
        $errorCode = "ACCOUNT_RESTRICTION"
        $userFriendlyMessage = "User account does not have permission to join computers to the domain"
        $suggestion = "Contact your domain administrator to grant domain join permissions or use an account with Domain Admin rights."
    }
    elseif ($errorMessage -match "0x0000232A") {
        $errorCode = "DNS_FAILURE"
        $userFriendlyMessage = "DNS name does not exist"
        $suggestion = "Verify the domain name is spelled correctly and DNS is properly configured."
    }
    elseif ($errorMessage -match "0x0000232B") {
        $errorCode = "DNS_SERVER_FAILURE"
        $userFriendlyMessage = "DNS server failure"
        $suggestion = "Check your DNS server settings and network connectivity."
    }
    elseif ($errorMessage -match "0x00000005") {
        $errorCode = "ACCESS_DENIED"
        $userFriendlyMessage = "Access denied - insufficient permissions"
        $suggestion = "You need Domain Admin rights or delegated permissions to join computers to the domain."
    }
    elseif ($errorMessage -match "password|credential") {
        $errorCode = "CREDENTIAL_ERROR"
        $userFriendlyMessage = "Credential authentication failed"
        $suggestion = "Verify username and password are correct. Check for typos and ensure CAPS LOCK is off."
    }
    elseif ($errorMessage -match "network|connection|timeout") {
        $errorCode = "NETWORK_ERROR"
        $userFriendlyMessage = "Network connectivity issue"
        $suggestion = "Check network cables, Wi-Fi connection, and firewall settings. Ensure you can ping the domain controller."
    }
    return @{
        ErrorCode           = $errorCode
        OriginalMessage     = $errorMessage
        UserFriendlyMessage = $userFriendlyMessage
        Suggestion          = $suggestion
    }
}


function Show-AzureADJoinDialog {
    param([string]$Username)
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    # Apply theme to dialog
    $currentTheme = if ($global:CurrentTheme) { $global:CurrentTheme } else { "Light" }
    $theme = $Themes[$currentTheme]
    
    $azForm = New-Object System.Windows.Forms.Form
    $azForm.Text = "AzureAD/Entra ID Join Required"
    $azForm.Size = New-Object System.Drawing.Size(600, 400)
    $azForm.StartPosition = "CenterScreen"
    $azForm.FormBorderStyle = "FixedDialog"
    $azForm.MaximizeBox = $false
    $azForm.MinimizeBox = $false
    $azForm.TopMost = $true
    $azForm.BackColor = $theme.FormBackColor
    $azForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Header
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size = New-Object System.Drawing.Size(600, 70)
    $headerPanel.BackColor = if ($currentTheme -eq "Dark") { [System.Drawing.Color]::FromArgb(139, 69, 19) } else { [System.Drawing.Color]::FromArgb(232, 17, 35) }
    $azForm.Controls.Add($headerPanel)
    
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(560, 25)
    $lblTitle.Text = "Microsoft Entra ID Join Required"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::White
    $headerPanel.Controls.Add($lblTitle)
    
    $lblSubtitle = New-Object System.Windows.Forms.Label
    $lblSubtitle.Location = New-Object System.Drawing.Point(22, 43)
    $lblSubtitle.Size = New-Object System.Drawing.Size(560, 20)
    $lblSubtitle.Text = "This profile requires an AzureAD/Entra ID joined device"
    $lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $headerPanel.Controls.Add($lblSubtitle)
    
    # Message - use multiple labels for proper line breaks
    $lblMessage1 = New-Object System.Windows.Forms.Label
    $lblMessage1.Location = New-Object System.Drawing.Point(20, 90)
    $lblMessage1.Size = New-Object System.Drawing.Size(550, 40)
    $lblMessage1.Text = "You are trying to import an AzureAD/Entra ID profile ($Username), but this computer is not joined to AzureAD/Entra ID."
    $lblMessage1.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblMessage1.ForeColor = $theme.LabelTextColor
    $azForm.Controls.Add($lblMessage1)
    
    $lblMessage2 = New-Object System.Windows.Forms.Label
    $lblMessage2.Location = New-Object System.Drawing.Point(20, 135)
    $lblMessage2.Size = New-Object System.Drawing.Size(550, 20)
    $lblMessage2.Text = "To import this profile, you must:"
    $lblMessage2.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblMessage2.ForeColor = $theme.LabelTextColor
    $azForm.Controls.Add($lblMessage2)
    
    $lblSteps = New-Object System.Windows.Forms.Label
    $lblSteps.Location = New-Object System.Drawing.Point(40, 160)
    $lblSteps.Size = New-Object System.Drawing.Size(530, 60)
    $lblSteps.Text = "1. Join this device to Microsoft Entra ID" + [Environment]::NewLine + "2. Click the button below to open AzureAD join settings"
    $lblSteps.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSteps.ForeColor = $theme.LabelTextColor
    $azForm.Controls.Add($lblSteps)
    
    # Instructions
    $lblInstructions = New-Object System.Windows.Forms.Label
    $lblInstructions.Location = New-Object System.Drawing.Point(20, 230)
    $lblInstructions.Size = New-Object System.Drawing.Size(550, 75)
    $lblInstructions.Text = "In Settings, look for:" + [Environment]::NewLine + "  - 'Access work or school' > 'Connect'" + [Environment]::NewLine + "  - Select 'Join this device to Microsoft Entra ID'" + [Environment]::NewLine + "  - Follow the prompts to sign in with your work/school account"
    $lblInstructions.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblInstructions.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $azForm.Controls.Add($lblInstructions)
    
    # Buttons
    $btnOpenSettings = New-Object System.Windows.Forms.Button
    $btnOpenSettings.Location = New-Object System.Drawing.Point(280, 320)
    $btnOpenSettings.Size = New-Object System.Drawing.Size(150, 35)
    $btnOpenSettings.Text = "Open Settings"
    $btnOpenSettings.BackColor = $theme.ButtonPrimaryBackColor
    $btnOpenSettings.ForeColor = $theme.ButtonPrimaryForeColor
    $btnOpenSettings.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOpenSettings.FlatAppearance.BorderSize = 0
    $btnOpenSettings.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOpenSettings.Cursor = [System.Windows.Forms.Cursors]::Hand
    $azForm.Controls.Add($btnOpenSettings)
    
    # Continue Import button (hidden initially)
    $btnContinue = New-Object System.Windows.Forms.Button
    $btnContinue.Location = New-Object System.Drawing.Point(280, 320)
    $btnContinue.Size = New-Object System.Drawing.Size(150, 35)
    $btnContinue.Text = "Continue Import"
    $btnContinue.BackColor = $theme.ButtonSuccessBackColor
    $btnContinue.ForeColor = $theme.ButtonSuccessForeColor
    $btnContinue.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnContinue.FlatAppearance.BorderSize = 0
    $btnContinue.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnContinue.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnContinue.DialogResult = "OK"
    $btnContinue.Visible = $false
    $azForm.Controls.Add($btnContinue)
    
    # When Open Settings is clicked, launch settings and swap buttons
    $btnOpenSettings.Add_Click({
            try {
                Start-Process "ms-settings:workplace"
                Log-Message "Opened AzureAD join settings (ms-settings:workplace)"
                $btnOpenSettings.Visible = $false
                $btnContinue.Visible = $true
                $lblMessage2.Text = "After joining to AzureAD, click Continue:"
                $azForm.AcceptButton = $btnContinue
            }
            catch {
                Log-Message "Failed to open settings: $_"
                Show-ModernDialog -Message ("Could not open settings automatically." + [Environment]::NewLine + [Environment]::NewLine + "Please open Settings manually and go to:" + [Environment]::NewLine + "Accounts > Access work or school") -Title "Error" -Type Error -Buttons OK
            }
        })
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(440, 320)
    $btnCancel.Size = New-Object System.Drawing.Size(130, 35)
    $btnCancel.Text = "Cancel Import"
    $btnCancel.BackColor = $theme.ButtonSecondaryBackColor
    $btnCancel.ForeColor = $theme.ButtonSecondaryForeColor
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.FlatAppearance.BorderSize = 0
    $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCancel.DialogResult = "Cancel"
    $azForm.Controls.Add($btnCancel)
    
    $azForm.AcceptButton = $btnOpenSettings
    $azForm.CancelButton = $btnCancel
    
    $result = $azForm.ShowDialog()
    $azForm.Dispose()
    return $result
}


function Invoke-AzureADUnjoin {
    <#
    .SYNOPSIS
    Unjoins the computer from AzureAD/Entra ID
    
    .DESCRIPTION
    Uses dsregcmd /leave to remove the device from AzureAD.
    Requires Administrator privileges.
    
    .RETURNS
    Hashtable with Success (bool) and Message (string)
    #>
    
    try {
        Log-Info "Unjoining from AzureAD..."
        
        # Verify we're actually AzureAD joined
        if (-not (Test-IsAzureADJoined)) {
            return @{
                Success = $false
                Message = "Device is not AzureAD joined"
            }
        }
        
        # Run dsregcmd /leave
        Log-Info "Executing: dsregcmd /leave"
        $result = & dsregcmd /leave 2>&1
        
        # Check if successful
        if ($LASTEXITCODE -eq 0) {
            Log-Info "Successfully unjoined from AzureAD"
            return @{
                Success = $true
                Message = "Device successfully unjoined from AzureAD"
            }
        }
        else {
            $errorMsg = $result -join "`n"
            Log-Error "Failed to unjoin from AzureAD: $errorMsg"
            return @{
                Success = $false
                Message = "Failed to unjoin: $errorMsg"
            }
        }
    }
    catch {
        Log-Error "Exception during AzureAD unjoin: $_"
        return @{
            Success = $false
            Message = "Exception: $_"
        }
    }
}


function Invoke-DomainUnjoin {
    <#
    .SYNOPSIS
    Unjoins the computer from the domain
    
    .DESCRIPTION
    Removes the computer from the domain and converts to workgroup.
    Requires Administrator privileges.
    
    .RETURNS
    Hashtable with Success (bool) and Message (string)
    #>
    
    try {
        Log-Info "Unjoining from domain..."
        
        # Get current domain
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
        if (-not $computerSystem.PartOfDomain) {
            return @{
                Success = $false
                Message = "Computer is not domain-joined"
            }
        }
        
        $currentDomain = $computerSystem.Domain
        Log-Info "Current domain: $currentDomain"
        
        # Unjoin from domain (convert to workgroup)
        $workgroupName = "WORKGROUP"
        
        Log-Info "Executing: Remove-Computer -WorkgroupName $workgroupName -Force"
        $result = Remove-Computer -WorkgroupName $workgroupName -Force -PassThru 2>&1
        
        # Check if successful
        if ($LASTEXITCODE -eq 0 -or $result.HasSucceeded) {
            Log-Info "Successfully unjoined from domain"
            return @{
                Success = $true
                Message = "Device successfully unjoined from $currentDomain"
            }
        }
        else {
            $errorMsg = $result -join "`n"
            Log-Error "Failed to unjoin from domain: $errorMsg"
            return @{
                Success = $false
                Message = "Failed to unjoin: $errorMsg"
            }
        }
    }
    catch {
        Log-Error "Exception during domain unjoin: $_"
        return @{
            Success = $false
            Message = "Exception: $_"
        }
    }
}


function Invoke-ForceUserLogoff {
    param(
        [string]$Username
    )
    
    try {
        Log-Info "Attempting to force logoff for user: $Username"
        
        # Strip domain if present for quser matching
        $shortName = if ($Username -match '\\') { ($Username -split '\\', 2)[1] } else { $Username }
        
        # Get session ID via quser
        $quserOutput = quser 2>&1 | Out-String
        $sessionId = $null
        
        $lines = $quserOutput -split "`n"
        foreach ($line in $lines) {
            # Skip header
            if ($line -match "^ USERNAME") { continue }
            
            # Simple containment check for username
            if ($line -match "^\s*>?\s*$([Regex]::Escape($shortName))\s+") {
                # Split by whitespace and filter empty
                $parts = ($line -split '\s+') | Where-Object { $_ -ne "" }
                
                # ID is usually the first integer found (index 2 or 1)
                foreach ($part in $parts) {
                    if ($part -match '^\d+$') {
                        $sessionId = $part
                        break
                    }
                }
            }
        }
        
        if ($sessionId) {
            Log-Info "Found session ID $sessionId for user $shortName. Executing logoff..."
            logoff $sessionId
            
            # Verification loop
            Log-Info "Waiting for user to log off..."
            $maxRetries = 10
            for ($i = 0; $i -lt $maxRetries; $i++) {
                Start-Sleep -Seconds 2
                
                $checkQuser = quser 2>&1 | Out-String
                if ($checkQuser -notmatch "^\s*>?\s*$([Regex]::Escape($shortName))\s+") {
                    Log-Info "User $shortName is no longer listed in quser. Logoff successful."
                    
                    # CRITICAL: Clear the State registry key to prevent temporary profile
                    # When Windows force-logs off a user, it sets State=516 (0x204) which indicates an error
                    # We need to reset this to 0 to allow normal login
                    try {
                        $userSid = $null
                        try {
                            $obj = New-Object System.Security.Principal.NTAccount($Username)
                            $userSid = $obj.Translate([System.Security.Principal.SecurityIdentifier]).Value
                        }
                        catch {
                            Log-Warning "Could not resolve SID for $Username to clear State flag: $_"
                        }
                        
                        if ($userSid) {
                            $profileKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$userSid"
                            if (Test-Path $profileKey) {
                                $currentState = (Get-ItemProperty -Path $profileKey -Name State -ErrorAction SilentlyContinue).State
                                if ($currentState -and $currentState -ne 0) {
                                    Log-Info "Clearing State registry flag (was $currentState) to prevent temporary profile"
                                    Set-ItemProperty -Path $profileKey -Name State -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
                                    Log-Info "State flag cleared successfully"
                                }
                            }
                        }
                    }
                    catch {
                        Log-Warning "Could not clear State registry flag: $_"
                    }
                    
                    return $true
                }
            }
            
            Log-Warning "Timed out waiting for user $shortName to disappear from quser."
            return $false
        }
        else {
            Log-Warning "Could not find active session ID for user $shortName via quser."
            return $false
        }
    }
    catch {
        Log-Error "Failed to force logoff: $_"
        return $false
    }
}


function Test-WingetFunctionality {
    Write-Host "  - Verifying Winget health..." -ForegroundColor Gray
    try {
        # Resolve Winget Path
        $wingetExe = "winget.exe"
        if (-not (Get-Command $wingetExe -ErrorAction SilentlyContinue)) {
            $localPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
            if (Test-Path $localPath) {
                $wingetExe = $localPath
            }
            else {
                return @{ Success = $false; Error = "Winget executable not found in PATH or standard location" }
            }
        }

        # Use Start-Process with a timeout to prevent hanging
        # Use 'source list' to verify the source configuration is actually readable
        $proc = Start-Process $wingetExe -ArgumentList "source", "list" -NoNewWindow -PassThru -ErrorAction SilentlyContinue
        
        # Wait for up to 10 seconds
        $timeout = 10
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        while (-not $proc.HasExited -and $timer.Elapsed.TotalSeconds -lt $timeout) {
            Start-Sleep -Milliseconds 500
        }

        if (-not $proc.HasExited) {
            $proc.Kill()
            return @{ Success = $false; Error = "Timed out after ${timeout}s" }
        }

        # Ensure process handle state is final
        $proc.WaitForExit()

        # Winget shim sometimes returns null ExitCode on success. 
        # Only fail if we have a specific non-zero exit code.
        if ($null -ne $proc.ExitCode -and $proc.ExitCode -ne 0) {
            return @{ Success = $false; Error = "Exit code $($proc.ExitCode)" }
        }
        
        return @{ Success = $true }
    }
    catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}


function Repair-WingetSources {
    Log-Message "Attempting to repair Winget sources..." 'WARN'
    try {
        # Update root certificates
        try {
            certutil.exe -generateSSTFromWU C:\Windows\Temp\roots.sst | Out-Null
            Import-Certificate -FilePath C:\Windows\Temp\roots.sst -CertStoreLocation Cert:\LocalMachine\Root | Out-Null
        }
        catch { }

        # Reset ALL sources to defaults (don't specify individual sources)
        Write-Host "  - Resetting Winget sources to defaults..." -ForegroundColor Gray
        
        # Resolve Winget Path
        $wingetExe = "winget.exe"
        if (-not (Get-Command $wingetExe -ErrorAction SilentlyContinue)) {
            $localPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
            if (Test-Path $localPath) {
                $wingetExe = $localPath
            }
        }
        
        $resetAll = Start-Process -FilePath $wingetExe -ArgumentList "source", "reset", "--force" -PassThru -NoNewWindow
        $resetAll.WaitForExit(30000)
        
        Log-Message "Winget source reset completed." 'INFO'
        return $true
    }
    catch {
        Log-Message "Winget repair failed: $_" 'ERROR'
        return $false
    }
}


function Install-WingetAppsFromExport {
    param(
        [Parameter(Mandatory = $true)][string]$TargetProfilePath
    )
    # In flat ZIP/import, Winget JSON is at the profile root
    $jsonPath = Join-Path $TargetProfilePath 'Winget-Packages.json'

    # --- Winget msstore source troubleshooting ---
    Log-Message "Checking winget functionality..." 'INFO'
    $health = Test-WingetFunctionality
    if (-not $health.Success) {
        Log-Warning "Winget issue detected: $($health.Error)"
        Repair-WingetSources | Out-Null
    }
    else {
        Log-Message "Winget sources appear healthy." 'INFO'
    }

    # 1. Safe file check
    if (-not (Test-Path $jsonPath) -or (Get-Item $jsonPath).Length -lt 10) {
        Log-Message "No valid Winget-Packages.json found - skipping app reinstall"
        return
    }

    $global:StatusText.Text = "Loading Winget app list..."
    $global:ProgressBar.Value = 98
    [System.Windows.Forms.Application]::DoEvents()

    # 2. Safe JSON parsing
    try {
        $json = Get-Content $jsonPath -Raw | ConvertFrom-Json -ErrorAction Stop
        $apps = $json.Sources.Packages
        if (-not $apps -or $apps.Count -eq 0) {
            Log-Message "Winget export is empty - nothing to install"
            return
        }
    }
    catch {
        Log-Message "Failed to parse Winget-Packages.json: $_ - skipping apps"
        $emptyJson = '{"Sources":[{"Packages":[]}]}'
        Set-Content $jsonPath -Value $emptyJson -Force -ErrorAction SilentlyContinue
        Show-ModernDialog -Message "The Winget app list is corrupted or unreadable.`r`n`r`nSkipping application reinstall." -Title "Winget List Error" -Type Warning -Buttons OK
        return
    }

    # 3. GUI to select apps - MODERN DESIGN
    $global:StatusText.Text = "Preparing app selection UI..."
    [System.Windows.Forms.Application]::DoEvents()
    
    # Apply theme
    $currentTheme = if ($global:CurrentTheme) { $global:CurrentTheme } else { "Light" }
    $theme = $Themes[$currentTheme]

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Application Installer"
    $form.Size = New-Object System.Drawing.Size(800, 700)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor = $theme.FormBackColor
    $form.MinimumSize = New-Object System.Drawing.Size(800, 700)

    # Header panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size = New-Object System.Drawing.Size(800, 70)
    $headerPanel.BackColor = $theme.HeaderBackColor
    $form.Controls.Add($headerPanel)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(500, 28)
    $lblTitle.Text = "Select Applications to Install"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $theme.HeaderTextColor
    $headerPanel.Controls.Add($lblTitle)

    $lblSubtitle = New-Object System.Windows.Forms.Label
    $lblSubtitle.Location = New-Object System.Drawing.Point(22, 45)
    $lblSubtitle.Size = New-Object System.Drawing.Size(750, 20)
    $lblSubtitle.Text = "Found $($apps.Count) applications from source PC - Select which ones to reinstall"
    $lblSubtitle.ForeColor = $theme.SubHeaderTextColor
    $headerPanel.Controls.Add($lblSubtitle)

    # Main content card
    $contentCard = New-Object System.Windows.Forms.Panel
    $contentCard.Location = New-Object System.Drawing.Point(15, 85)
    $contentCard.Size = New-Object System.Drawing.Size(755, 500)
    $contentCard.BackColor = $theme.PanelBackColor
    $form.Controls.Add($contentCard)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = New-Object System.Drawing.Point(15, 15)
    $lblInfo.Size = New-Object System.Drawing.Size(720, 35)
    $lblInfo.Text = "Winget will automatically install or upgrade these applications.`nUse the checkboxes to select which apps you want:"
    $lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblInfo.ForeColor = $theme.LabelTextColor
    $contentCard.Controls.Add($lblInfo)

    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(15, 55)
    $clb.Size = New-Object System.Drawing.Size(720, 350)
    $clb.CheckOnClick = $true
    $clb.Font = New-Object System.Drawing.Font("Consolas", 9)
    $clb.BackColor = $theme.LogBoxBackColor
    $clb.ForeColor = $theme.LogBoxForeColor
    $clb.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    foreach ($app in $apps) {
        $name = $app.PackageIdentifier
        if ($app.Version) { $name += "  [v$($app.Version)]" }
        $clb.Items.Add($name, $true) | Out-Null
    }
    $contentCard.Controls.Add($clb)

    # Selection controls
    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Location = New-Object System.Drawing.Point(15, 415)
    $btnAll.Size = New-Object System.Drawing.Size(120, 35)
    $btnAll.Text = "Select All"
    $btnAll.BackColor = $theme.ButtonPrimaryBackColor
    $btnAll.ForeColor = $theme.ButtonPrimaryForeColor
    $btnAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnAll.FlatAppearance.BorderSize = 0
    $btnAll.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnAll.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnAll.Add_MouseEnter({ $this.BackColor = $theme.ButtonPrimaryHoverBackColor })
    $btnAll.Add_MouseLeave({ $this.BackColor = $theme.ButtonPrimaryBackColor })
    $btnAll.Add_Click({ for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $true) } })
    $contentCard.Controls.Add($btnAll)

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Location = New-Object System.Drawing.Point(145, 415)
    $btnNone.Size = New-Object System.Drawing.Size(120, 35)
    $btnNone.Text = "Select None"
    $btnNone.BackColor = $theme.ButtonSecondaryBackColor
    $btnNone.ForeColor = $theme.ButtonSecondaryForeColor
    $btnNone.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnNone.FlatAppearance.BorderSize = 0
    $btnNone.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnNone.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnNone.Add_MouseEnter({ $this.BackColor = $theme.ButtonSecondaryHoverBackColor })
    $btnNone.Add_MouseLeave({ $this.BackColor = $theme.ButtonSecondaryBackColor })
    $btnNone.Add_Click({ for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $false) } })
    $contentCard.Controls.Add($btnNone)

    # Selection counter
    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Location = New-Object System.Drawing.Point(15, 460)
    $lblCount.Size = New-Object System.Drawing.Size(720, 25)
    $lblCount.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblCount.ForeColor = $theme.HeaderTextColor
    $contentCard.Controls.Add($lblCount)
    
    # Update counter function
    $updateCounter = {
        $checked = 0
        for ($i = 0; $i -lt $clb.Items.Count; $i++) {
            if ($clb.GetItemChecked($i)) { $checked++ }
        }
        $lblCount.Text = "$checked of $($clb.Items.Count) applications selected for installation"
    }
    & $updateCounter
    $clb.Add_ItemCheck({ 
            # Delay to let the check state update
            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 10
            $timer.Add_Tick({
                    & $updateCounter
                    $this.Stop()
                    $this.Dispose()
                })
            $timer.Start()
        })

    # Action buttons at bottom
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(540, 600)
    $btnOK.Size = New-Object System.Drawing.Size(110, 40)
    $btnOK.Text = "Install"
    $btnOK.BackColor = $theme.ButtonSuccessBackColor
    $btnOK.ForeColor = $theme.ButtonSuccessForeColor
    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOK.FlatAppearance.BorderSize = 0
    $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnOK.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnOK.Add_MouseEnter({ $this.BackColor = $theme.ButtonSuccessHoverBackColor })
    $btnOK.Add_MouseLeave({ $this.BackColor = $theme.ButtonSuccessBackColor })
    $btnOK.DialogResult = "OK"
    $form.Controls.Add($btnOK)

    $btnSkip = New-Object System.Windows.Forms.Button
    $btnSkip.Location = New-Object System.Drawing.Point(660, 600)
    $btnSkip.Size = New-Object System.Drawing.Size(110, 40)
    $btnSkip.Text = "Skip All"
    $btnSkip.BackColor = $theme.ButtonSecondaryBackColor
    $btnSkip.ForeColor = $theme.ButtonSecondaryForeColor
    $btnSkip.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSkip.FlatAppearance.BorderSize = 0
    $btnSkip.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnSkip.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnSkip.Add_MouseEnter({ $this.BackColor = $theme.ButtonSecondaryHoverBackColor })
    $btnSkip.Add_MouseLeave({ $this.BackColor = $theme.ButtonSecondaryBackColor })
    $btnSkip.DialogResult = "Cancel"
    $form.Controls.Add($btnSkip)

    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnSkip

    if ($form.ShowDialog() -ne "OK") {
        Log-Message "User skipped Winget app installation"
        $form.Dispose()
        return
    }
    $form.Dispose()

    # 4. Collect selected apps and create filtered JSON
    $selectedApps = @()
    for ($i = 0; $i -lt $clb.Items.Count; $i++) {
        if ($clb.GetItemChecked($i)) {
            $selectedApps += $apps[$i]
        }
    }

    if ($selectedApps.Count -eq 0) {
        Log-Message "No apps selected for install"
        return
    }

    # 5. Track installed apps globally for report
    $global:InstalledAppsList = @()
    
    # 6. Install apps individually for better tracking
    $totalApps = $selectedApps.Count
    $currentAppIndex = 0
    $successCount = 0
    $failCount = 0
    
    Log-Message "Starting individual Winget app installation for $totalApps applications"
    
    foreach ($app in $selectedApps) {
        $currentAppIndex++
        $appId = $app.PackageIdentifier
        
        $global:StatusText.Text = "Installing ($currentAppIndex/$totalApps): $appId"
        # Progress range 98-100%
        $progressCalc = 98 + ([Math]::Round(($currentAppIndex / $totalApps) * 2))
        $global:ProgressBar.Value = [Math]::Min(100, $progressCalc)
        [System.Windows.Forms.Application]::DoEvents()
        
        Log-Message "Installing [$currentAppIndex/$totalApps]: $appId" 'INFO'
        
        try {
            # Execute winget install
            # --silent: No UI
            # --accept-package-agreements / --accept-source-agreements: Skip prompts
            # --source winget: Force community source to avoid msstore certificate issues
            $proc = Start-Process -FilePath "winget" -ArgumentList "install", "--id", "$appId", "--accept-package-agreements", "--accept-source-agreements", "--silent", "--source", "winget" -Wait -PassThru -ErrorAction Stop
            
            # 0 = Success
            # 0x8A15002B (-1978335189) = Already installed
            $isAlreadyInstalled = ($proc.ExitCode -eq -1978335189)
            
            if ($proc.ExitCode -eq 0 -or $isAlreadyInstalled) {
                if ($isAlreadyInstalled) {
                    Log-Message "${appId} is already installed. Checking for updates..." 'INFO'
                    # Try to upgrade if it already exists
                    Start-Process -FilePath "winget" -ArgumentList "upgrade", "--id", "$appId", "--accept-package-agreements", "--accept-source-agreements", "--silent", "--source", "winget" -Wait -NoNewWindow
                }
                Log-Message "Successfully processed ${appId}" 'INFO'
                $global:InstalledAppsList += $app
                $successCount++
            }
            else {
                Log-Message "Failed to install ${appId} (Exit code: $($proc.ExitCode))" 'WARN'
                $failCount++
            }
        }
        catch {
            Log-Message "Exception installing ${appId}: $($_.Exception.Message)" 'ERROR'
            $failCount++
        }
    }
    
    $summary = "Winget process complete. $successCount successful, $failCount failed of $totalApps total."
    Log-Message $summary
    $global:StatusText.Text = "Apps finished ($successCount/$totalApps installed)"
    $global:ProgressBar.Value = 100
    [System.Windows.Forms.Application]::DoEvents()
    
    $dialogType = if ($failCount -eq 0) { 'Success' } else { 'Warning' }
    Show-ModernDialog -Message $summary -Title "Installation Complete" -Type $dialogType -Buttons OK
}


function Add-AppxReregistrationActiveSetup {
    param(
        [Parameter(Mandatory = $true)][string]$Username,
        [Parameter(Mandatory = $true)][string]$OperationType  # "Conversion" or "Import"
    )
    
    try {
        Log-Info "Setting up AppX package re-registration via Active Setup..."
        
        # Create script directory if it doesn't exist
        $scriptDir = "C:\ProgramData\ProfileMigration"
        if (-not (Test-Path $scriptDir)) {
            New-Item -ItemType Directory -Path $scriptDir -Force | Out-Null
        }
        
        # GRANT PERMISSIONS: Allow Users to write/modify files in this directory (for logging)
        try {
            $acl = Get-Acl $scriptDir
            # Check if rule already exists to avoid duplication? AddAccessRule handles it reasonably well
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule("Users", "Modify", "ContainerInherit,ObjectInherit", "None", "Allow")
            $acl.AddAccessRule($rule)
            Set-Acl -Path $scriptDir -AclObject $acl
        }
        catch {
            Log-Warning "Failed to set permissions on log directory: $_"
        }
        
        # Create timestamp and unique ID
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $sanitizedUser = $Username -replace '[\\/:*?"<>| ]', '_'
        $scriptPath = Join-Path $scriptDir "ReregisterApps_${sanitizedUser}_${timestamp}.ps1"
        $componentID = "ProfileMigration_AppXRepair_${sanitizedUser}_${timestamp}"
        
        # Extract short username for validation (handle DOMAIN\user and user@domain.com)
        $shortTargetUser = if ($Username -match '\\|@') {
            if ($Username -match '@') { ($Username -split '@')[0] }
            else { ($Username -split '\\', 2)[1] }
        }
        else { $Username }
        
        # Resolve Theme Colors for the splash screen
        $currTheme = if ($global:CurrentTheme) { $global:CurrentTheme } else { "Dark" }
        $t = $global:Themes[$currTheme]
        # Helper to format color as "R, G, B"
        $bgColor = "$($t.FormBackColor.R), $($t.FormBackColor.G), $($t.FormBackColor.B)"
        $fgColor = "$($t.LabelTextColor.R), $($t.LabelTextColor.G), $($t.LabelTextColor.B)"
        $headerColor = "$($t.HeaderTextColor.R), $($t.HeaderTextColor.G), $($t.HeaderTextColor.B)"

        # Create the re-registration script
        # Includes a check to ensure it only runs for the target user
        $reregScript = @"
# AppX Package Re-registration Script (Active Setup)
# Generated by Profile Migration Tool v$($global:Config.Version)
# Operation: $OperationType
# Target User: $Username
# Created: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

# Set up logging
`$logPath = "C:\ProgramData\ProfileMigration\AppxReregistration_${sanitizedUser}_${timestamp}.log"
function Write-Log {
    param([string]`$Message)
    `$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[`$timestamp] `$Message" | Out-File -FilePath `$logPath -Append -Encoding UTF8
}

Write-Log "=== AppX Package Re-registration Started (Active Setup) ==="
Write-Log "User: `$env:USERNAME"
Write-Log "Domain: `$env:USERDOMAIN"
Write-Log "Operation Type: $OperationType"

# Log the user running this script (for diagnostics)
`$targetUser = '$Username'
`$targetShortName = '$shortTargetUser'
`$currentShortName = `$env:USERNAME

Write-Log \"Target user: `$targetUser (Short: `$targetShortName)\"
Write-Log \"Running as: `$env:USERNAME (Domain: `$env:USERDOMAIN)\"

# VALIDATION: Ensure we are running as the correct user
# Since simple string matching can be tricky with domains, we compare the "short" username
# which matches $env:USERNAME behavior
if (`$currentShortName -ne `$targetShortName) {
    Write-Log "MISMATCH: Current user (`$currentShortName) != Target (`$targetShortName). Exiting to preserve script for target user."
    exit
}

Write-Log \"User match confirmed. Proceeding...\"

try {
    # UI SETUP: Create a simple splash screen
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    `$form = New-Object System.Windows.Forms.Form
    `$form.Text = "Profile Migration"
    `$form.Size = New-Object System.Drawing.Size(450, 160)
    `$form.StartPosition = "CenterScreen"
    `$form.FormBorderStyle = "FixedDialog"
    `$form.ControlBox = `$false
    `$form.TopMost = `$true
    `$form.BackColor = [System.Drawing.Color]::FromArgb($bgColor)

    `$labelTitle = New-Object System.Windows.Forms.Label
    `$labelTitle.Location = New-Object System.Drawing.Point(20, 15)
    `$labelTitle.Size = New-Object System.Drawing.Size(400, 25)
    `$labelTitle.Text = "Finalizing Profile Setup"
    `$labelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    `$labelTitle.ForeColor = [System.Drawing.Color]::FromArgb($headerColor)
    `$form.Controls.Add(`$labelTitle)

    `$labelDesc = New-Object System.Windows.Forms.Label
    `$labelDesc.Location = New-Object System.Drawing.Point(22, 45)
    `$labelDesc.Size = New-Object System.Drawing.Size(400, 20)
    `$labelDesc.Text = "Please wait while we update your applications..."
    `$labelDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    `$labelDesc.ForeColor = [System.Drawing.Color]::FromArgb($fgColor)
    `$form.Controls.Add(`$labelDesc)

    `$lblStatus = New-Object System.Windows.Forms.Label
    `$lblStatus.Location = New-Object System.Drawing.Point(22, 75)
    `$lblStatus.Size = New-Object System.Drawing.Size(400, 20)
    `$lblStatus.Text = "Starting..."
    `$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    `$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb($fgColor)
    `$form.Controls.Add(`$lblStatus)

    `$progressBar = New-Object System.Windows.Forms.ProgressBar
    `$progressBar.Location = New-Object System.Drawing.Point(20, 100)
    `$progressBar.Size = New-Object System.Drawing.Size(400, 10)
    `$progressBar.Style = "Continuous" 
    `$form.Controls.Add(`$progressBar)

    # Show form non-modally
    `$form.Show()
    `$form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()

    # Get AppX packages for CURRENT USER only (avoid Access Denied)
    `$packages = Get-AppxPackage
    `$successCount = 0
    `$failCount = 0
    `$total = `$packages.Count
    
    `$progressBar.Maximum = `$total
    `$current = 0

    Write-Log "Found `$total AppX packages to re-register"
    
    foreach (`$package in `$packages) {
        `$current++
        
        # Update UI every few items to stay responsive
        if (`$current % 3 -eq 0 -or `$current -eq `$total) {
            `$lblStatus.Text = "Updating: `$(`$package.Name)"
            `$progressBar.Value = `$current
            [System.Windows.Forms.Application]::DoEvents()
        }

        try {
            `$manifest = "`$(`$package.InstallLocation)\AppXManifest.xml"
            if (Test-Path `$manifest) {
                Add-AppxPackage -DisableDevelopmentMode -Register `$manifest -ErrorAction Stop
                `$successCount++
                Write-Log "SUCCESS: Re-registered `$(`$package.Name)"
            }
            else {
                Write-Log "SKIP: Manifest not found for `$(`$package.Name)"
            }
        }
        catch {
            `$failCount++
            Write-Log "FAILED: Could not re-register `$(`$package.Name) - `$_"
        }
    }
    
    `$lblStatus.Text = "Done!"
    `$progressBar.Value = `$total
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 500
    
    `$form.Close()
    `$form.Dispose()

    Write-Log "=== AppX Package Re-registration Completed ==="
    Write-Log "Successfully re-registered: `$successCount packages"
    Write-Log "Failed to re-register: `$failCount packages"
    
    # Clean up this script after execution
    Start-Sleep -Seconds 1
    Remove-Item -Path `$PSCommandPath -Force -ErrorAction SilentlyContinue
}
catch {
    Write-Log "ERROR: Re-registration failed - `$_"
    if (`$form) { `$form.Close() }
}
"@
        
        # Write the script to disk
        Set-Content -Path $scriptPath -Value $reregScript -Encoding UTF8
        Log-Info "Created AppX re-registration script: $scriptPath"
        
        # Add Active Setup registry key
        # Reads from HKLM, runs for each user once (tracked in HKCU)
        $activeSetupPath = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\$componentID"
        if (-not (Test-Path $activeSetupPath)) {
            New-Item -Path $activeSetupPath -Force | Out-Null
        }
        
        # StubPath: The command to run
        $stubPath = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
        
        Set-ItemProperty -Path $activeSetupPath -Name "StubPath" -Value $stubPath
        Set-ItemProperty -Path $activeSetupPath -Name "Version" -Value "1,0,0"
        Set-ItemProperty -Path $activeSetupPath -Name "(default)" -Value "Profile Migration - AppX Repair"
        Set-ItemProperty -Path $activeSetupPath -Name "IsInstalled" -Value 1 -Type DWord
        
        Log-Info "Added Active Setup key: $componentID"
        Log-Info "AppX packages will be re-registered on $Username's next login (before desktop load)"
        
        return $true
    }
    catch {

        Log-Warning "Failed to set up AppX re-registration: $_"
        return $false
    }
}


function Show-SevenZipRecoveryDialog {
    param([string]$Message)
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    # Apply theme to dialog
    $currentTheme = if ($global:CurrentTheme) { $global:CurrentTheme } else { "Light" }
    $theme = $Themes[$currentTheme]
    
    $recoveryForm = New-Object System.Windows.Forms.Form
    $recoveryForm.Text = "7-Zip Required"
    $recoveryForm.Size = New-Object System.Drawing.Size(600, 380)
    $recoveryForm.StartPosition = "CenterScreen"
    $recoveryForm.FormBorderStyle = "FixedDialog"
    $recoveryForm.MaximizeBox = $false
    $recoveryForm.MinimizeBox = $false
    $recoveryForm.TopMost = $true
    $recoveryForm.BackColor = $theme.FormBackColor
    $recoveryForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Header panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size = New-Object System.Drawing.Size(600, 70)
    $headerPanel.BackColor = if ($currentTheme -eq "Dark") { [System.Drawing.Color]::FromArgb(139, 69, 19) } else { [System.Drawing.Color]::FromArgb(232, 17, 35) }
    $recoveryForm.Controls.Add($headerPanel)
    
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(560, 25)
    $lblTitle.Text = "7-Zip Not Found"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::White
    $headerPanel.Controls.Add($lblTitle)
    
    $lblSubtitle = New-Object System.Windows.Forms.Label
    $lblSubtitle.Location = New-Object System.Drawing.Point(22, 43)
    $lblSubtitle.Size = New-Object System.Drawing.Size(560, 20)
    $lblSubtitle.Text = "7-Zip is required for profile migration"
    $lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 200)
    $headerPanel.Controls.Add($lblSubtitle)
    
    # Content panel
    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Location = New-Object System.Drawing.Point(15, 85)
    $contentPanel.Size = New-Object System.Drawing.Size(560, 180)
    $contentPanel.BackColor = [System.Drawing.Color]::White
    $recoveryForm.Controls.Add($contentPanel)
    
    $lblMessage = New-Object System.Windows.Forms.Label
    $lblMessage.Location = New-Object System.Drawing.Point(15, 15)
    $lblMessage.Size = New-Object System.Drawing.Size(530, 60)
    $lblMessage.Text = $Message
    $lblMessage.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $contentPanel.Controls.Add($lblMessage)
    
    $lblOptions = New-Object System.Windows.Forms.Label
    $lblOptions.Location = New-Object System.Drawing.Point(15, 80)
    $lblOptions.Size = New-Object System.Drawing.Size(530, 20)
    $lblOptions.Text = "Choose an option:"
    $lblOptions.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $contentPanel.Controls.Add($lblOptions)
    
    # Browse for 7z.exe button
    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Location = New-Object System.Drawing.Point(30, 110)
    $btnBrowse.Size = New-Object System.Drawing.Size(240, 35)
    $btnBrowse.Text = "Browse for 7z.exe"
    $btnBrowse.BackColor = $theme.ButtonPrimaryBackColor
    $btnBrowse.ForeColor = $theme.ButtonPrimaryForeColor
    $btnBrowse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnBrowse.FlatAppearance.BorderSize = 0
    $btnBrowse.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnBrowse.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnBrowse.Add_MouseEnter({ $this.BackColor = $theme.ButtonPrimaryHoverBackColor })
    $btnBrowse.Add_MouseLeave({ $this.BackColor = $theme.ButtonPrimaryBackColor })
    $btnBrowse.Add_Click({
            $openDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openDialog.Title = "Locate 7z.exe"
            $openDialog.Filter = "7-Zip Executable (7z.exe)|7z.exe|All Files (*.*)|*.*"
            $openDialog.InitialDirectory = ${env:ProgramFiles}
            if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                if (Test-Path $openDialog.FileName) {
                    $recoveryForm.Tag = @{ Action = "Browse"; Path = $openDialog.FileName }
                    $recoveryForm.Close()
                }
            }
        })
    $contentPanel.Controls.Add($btnBrowse)
    
    # Retry winget install button
    $btnRetry = New-Object System.Windows.Forms.Button
    $btnRetry.Location = New-Object System.Drawing.Point(290, 110)
    $btnRetry.Size = New-Object System.Drawing.Size(240, 35)
    $btnRetry.Text = "Retry Winget Install"
    $btnRetry.BackColor = $theme.ButtonSuccessBackColor
    $btnRetry.ForeColor = $theme.ButtonSuccessForeColor
    $btnRetry.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRetry.FlatAppearance.BorderSize = 0
    $btnRetry.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnRetry.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnRetry.Add_MouseEnter({ $this.BackColor = $theme.ButtonSuccessHoverBackColor })
    $btnRetry.Add_MouseLeave({ $this.BackColor = $theme.ButtonSuccessBackColor })
    $btnRetry.Add_Click({
            $recoveryForm.Tag = @{ Action = "Retry" }
            $recoveryForm.Close()
        })
    $contentPanel.Controls.Add($btnRetry)
    
    # Download from website button
    $btnDownload = New-Object System.Windows.Forms.Button
    $btnDownload.Location = New-Object System.Drawing.Point(160, 155)
    $btnDownload.Size = New-Object System.Drawing.Size(240, 35)
    $btnDownload.Text = "Download from Website"
    $btnDownload.BackColor = $theme.ButtonSecondaryBackColor
    $btnDownload.ForeColor = $theme.ButtonSecondaryForeColor
    $btnDownload.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDownload.FlatAppearance.BorderSize = 0
    $btnDownload.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnDownload.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnDownload.Add_MouseEnter({ $this.BackColor = $theme.ButtonSecondaryHoverBackColor })
    $btnDownload.Add_MouseLeave({ $this.BackColor = $theme.ButtonSecondaryBackColor })
    $btnDownload.Add_Click({
            try {
                Start-Process "https://www.7-zip.org/download.html"
                Show-ModernDialog -Message "7-Zip download page opened in browser.`r`n`r`nAfter installation, click 'Browse for 7z.exe' or 'Retry Winget Install'." -Title "Download Started" -Type Info -Buttons OK
            }
            catch {
                Show-ModernDialog -Message "Failed to open browser: $_" -Title "Error" -Type Error -Buttons OK
            }
        })
    $contentPanel.Controls.Add($btnDownload)
    
    # Cancel button
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(450, 290)
    $btnCancel.Size = New-Object System.Drawing.Size(120, 35)
    $btnCancel.Text = "Exit"
    $btnCancel.BackColor = $theme.ButtonSecondaryBackColor
    $btnCancel.ForeColor = $theme.ButtonSecondaryForeColor
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.FlatAppearance.BorderSize = 0
    $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCancel.Add_MouseEnter({ $this.BackColor = $theme.ButtonSecondaryHoverBackColor })
    $btnCancel.Add_MouseLeave({ $this.BackColor = $theme.ButtonSecondaryBackColor })
    $btnCancel.Add_Click({
            $recoveryForm.Tag = @{ Action = "Exit" }
            $recoveryForm.Close()
        })
    $recoveryForm.Controls.Add($btnCancel)
    
    $recoveryForm.ShowDialog() | Out-Null
    return $recoveryForm.Tag
}


function Get-ZipUncompressedSize {
    param([string]$SevenZipPath, [string]$ZipPath)
    
    if (-not (Test-Path $SevenZipPath) -or -not (Test-Path $ZipPath)) { return 0 }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $SevenZipPath
    # Use -slt for technical listing (reliable parsing of Size = fields)
    $psi.Arguments = "l -slt `"$ZipPath`""
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.CreateNoWindow = $true
    
    $proc = [System.Diagnostics.Process]::Start($psi)
    # Read output line by line to avoid memory issues with huge logs? 
    # ReadToEnd is fine for text output usually, but for 100k files it might be large.
    # However, Generate-ImportReport uses ReadToEnd(), so it should be fine here.
    $output = $proc.StandardOutput.ReadToEnd()
    $proc.WaitForExit()
    
    $totalSize = 0
    $lines = $output -split "`n"
    foreach ($line in $lines) {
        if ($line -match '^\s*Size\s*=\s*(\d+)\s*$') {
            $totalSize += [long]$matches[1]
        }
    }
    
    return $totalSize
}


function Test-ZipIntegrity {
    param(
        [string]$SevenZipPath,
        [string]$ZipPath
    )
    
    Log-Info "Verifying ZIP integrity: $ZipPath"
    
    if (-not (Test-Path $ZipPath)) {
        Log-Error "ZIP file not found: $ZipPath"
        return $false
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $SevenZipPath
    $psi.Arguments = "t `"$ZipPath`""
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.Start() | Out-Null
    $proc.WaitForExit()
    
    if ($proc.ExitCode -eq 0) {
        Log-Info "ZIP integrity check passed."
        return $true
    }
    else {
        Log-Error "ZIP integrity check FAILED (Exit Code: $($proc.ExitCode))"
        return $false
    }
}


function Invoke-ProactiveUserCheck {
    param(
        [string]$Username
    )

    if (-not $Username) { return $true }
    
    # Normalize username
    $checkName = $Username
    if ($checkName -match '^(.+?)\s+-\s+\[.+\]$') { $checkName = $matches[1] }
    
    # Resolve SID
    $checkSid = $null
    try {
        $obj = New-Object System.Security.Principal.NTAccount($checkName)
        $checkSid = $obj.Translate([System.Security.Principal.SecurityIdentifier]).Value
    }
    catch {
        try {
            $sName = if ($checkName -match '\\') { ($checkName -split '\\', 2)[1] } else { $checkName }
            $obj = New-Object System.Security.Principal.NTAccount($sName)
            $checkSid = $obj.Translate([System.Security.Principal.SecurityIdentifier]).Value
        }
        catch {}
    }
    
    if ($checkSid) {
        if (Test-ProfileMounted -UserSID $checkSid) {
            $res = Show-ModernDialog -Message "The user '$checkName' appears to be logged in (profile registry hive is mounted).`r`n`r`nImporting, Exporting, or Converting a logged-in user profile WILL result in locked files and data corruption.`r`n`r`nWould you like to force log off this user now and continue?" -Title "User Logged In Warning" -Type Warning -Buttons YesNo
            
            if ($res -eq "Yes") {
                $logoffSuccess = Invoke-ForceUserLogoff -Username $checkName
                if ($logoffSuccess) {
                    if (Test-ProfileMounted -UserSID $checkSid) {
                        Show-ModernDialog -Message "User was logged off but profile hive is still mounted. Please wait a moment and try again." -Title "Profile Still Mounted" -Type Warning -Buttons OK
                        return $false
                    }
                    # Success
                    return $true
                }
                else {
                    Show-ModernDialog -Message "Failed to force log off the user. Please log them off manually." -Title "Logoff Failed" -Type Error -Buttons OK
                    return $false
                }
            }
            else {
                return $false
            }
        }
    }
    
    return $true
}


