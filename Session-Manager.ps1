Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Remote Session Manager"
$form.Size = New-Object System.Drawing.Size(800,600)
$form.StartPosition = "CenterScreen"
$form.Icon = ".\Images\logo.ico"
$form.maximizeBox = $false
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

# Create server input label
$labelServers = New-Object System.Windows.Forms.Label
$labelServers.Location = New-Object System.Drawing.Point(10,20)
$labelServers.Size = New-Object System.Drawing.Size(400,20)
$labelServers.Text = "Enter server names (comma-separated):"
$form.Controls.Add($labelServers)

# Create server input textbox
$textBoxServers = New-Object System.Windows.Forms.TextBox
$textBoxServers.Location = New-Object System.Drawing.Point(10,40)
$textBoxServers.Size = New-Object System.Drawing.Size(400,20)
$form.Controls.Add($textBoxServers)

# Create search label
$labelSearch = New-Object System.Windows.Forms.Label
$labelSearch.Location = New-Object System.Drawing.Point(420,20)
$labelSearch.Size = New-Object System.Drawing.Size(100,20)
$labelSearch.Text = "Search:"
$form.Controls.Add($labelSearch)

# Create search textbox
$textBoxSearch = New-Object System.Windows.Forms.TextBox
$textBoxSearch.Location = New-Object System.Drawing.Point(420,40)
$textBoxSearch.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxSearch)

# Create refresh button
$buttonRefresh = New-Object System.Windows.Forms.Button
$buttonRefresh.Location = New-Object System.Drawing.Point(700,38)
$buttonRefresh.Size = New-Object System.Drawing.Size(75,24)
$buttonRefresh.Text = "Refresh"
$form.Controls.Add($buttonRefresh)

# Create settings button
$buttonSettings = New-Object System.Windows.Forms.Button
$buttonSettings.Location = New-Object System.Drawing.Point(750,10)
$buttonSettings.Size = New-Object System.Drawing.Size(24,24)
$buttonSettings.Image = [System.Drawing.Image]::FromFile(".\Images\settings.ico")
$buttonSettings.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($buttonSettings)

# Global variables to keep track of the current sort column and order
$global:SortColumn = "Username"
$global:SortOrder = "Ascending"
$global:CachedSessions = @()

# Create ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10,70)
$listView.Size = New-Object System.Drawing.Size(765,480)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Scrollable = $true

# Add columns
$listView.Columns.Add("Username", 120)
$listView.Columns.Add("Display Name", 130)
$listView.Columns.Add("Server", 130)
$listView.Columns.Add("Session ID", 100)
$listView.Columns.Add("State", 90)
$listView.Columns.Add("C: Space Free", 100)
$listView.Columns.Add("Sessions Open", 90)
$form.Controls.Add($listView)

# Enable mouse wheel scrolling
$listView.add_MouseEnter({
    $listView.Focus()
})

# Function to get display name from Active Directory
function Get-DisplayName {
    param (
        [string]$Username
    )
    try {
        $searcher = [adsisearcher]"(samaccountname=$Username)"
        $result = $searcher.FindOne()
        if ($result) {
            return $result.Properties['displayname'][0]
        }
        else {
            return $Username
        }
    }
    catch {
        return $Username
    }
}

# Function to show progress form
function Show-ProgressForm {
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Loading..."
    $progressForm.Size = New-Object System.Drawing.Size(300,100)
    $progressForm.StartPosition = "CenterScreen"
    $progressForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $progressForm.ControlBox = $false
    $progressForm.TopMost = $true

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $progressBar.Dock = [System.Windows.Forms.DockStyle]::Top
    $progressForm.Controls.Add($progressBar)

    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $progressLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $progressForm.Controls.Add($progressLabel)

    $progressForm.Add_Shown({
        $progressForm.Activate()
    })

    return $progressForm, $progressBar, $progressLabel
}

# Function to get sessions
function Get-RemoteSessions {
    param (
        [string[]]$ServerList,
        [System.Windows.Forms.ProgressBar]$ProgressBar = $null,
        [System.Windows.Forms.Label]$ProgressLabel = $null
    )
   
    $settings = Read-Settings
    $sessions = @()
    $diskSpace = @{}
    $userCounts = @{}
    $totalServers = $ServerList.Count
    $currentServer = 0

    if ($settings.UseMultithreading -eq 1) {
        $jobs = @()
        foreach ($server in $ServerList) {
            $jobs += Start-Job -ScriptBlock {
                param($server)
                $result = @{}
                try {
                    $diskSpace = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $server -Filter "DeviceID='C:'" -ErrorAction Stop
                    $result.DiskSpace = [math]::Round($diskSpace.FreeSpace / 1GB, 2)
                }
                catch {
                    $result.DiskSpace = "Error"
                }

                try {
                    $query = quser /server:$server 2>&1
                    if ($query -notmatch "ERROR") {
                        $sessions = @()
                        $query | Select-Object -Skip 1 | ForEach-Object {
                            $line = $_.Trim() -replace '\s+', ' ' -split '\s'
                            if ($line.Count -ge 3) {
                                $sessionId = $line | Where-Object { $_ -match '^\d+$' } | Select-Object -First 1
                                $username = $line[0]
                                $state = $line | Where-Object { $_ -match 'Active|Disc' } | Select-Object -First 1
                                $sessions += [PSCustomObject]@{
                                    Username = $username
                                    Server = $server
                                    SessionID = $sessionId
                                    State = $state
                                }
                            }
                        }
                        $result.Sessions = $sessions
                    }
                }
                catch {
                    $result.Sessions = @()
                }
                return $result
            } -ArgumentList $server
        }

        # Wait for jobs to complete and collect results
        $jobs | Wait-Job
        foreach ($job in $jobs) {
            $output = Receive-Job -Job $job
            $diskSpace[$output.Server] = $output.DiskSpace
            $sessions += $output.Sessions
            Remove-Job -Job $job
        }
    }
    else {
        foreach ($server in $ServerList) {
            $currentServer++
            if ($ProgressBar -ne $null) {
                $ProgressBar.Value = [math]::Round(($currentServer / $totalServers) * 100)
            }
            if ($ProgressLabel -ne $null) {
                $ProgressLabel.Text = "Querying server: $server"
            }
            [System.Windows.Forms.Application]::DoEvents()

            # Get disk space once per server
            $diskSpace[$server] = Get-RemoteDiskSpace -Server $server
            $userCounts[$server] = 0
           
            try {
                $query = quser /server:$server 2>&1
                if ($query -match "ERROR") {
                    Write-Warning "Failed to query $server : $query"
                    continue
                }
               
                $query | Select-Object -Skip 1 | ForEach-Object {
                    $line = $_.Trim() -replace '\s+', ' ' -split '\s'
                    if ($line.Count -ge 3) {
                        $sessionId = $line | Where-Object { $_ -match '^\d+$' } | Select-Object -First 1
                        $username = $line[0]
                        $state = $line | Where-Object { $_ -match 'Active|Disc' } | Select-Object -First 1
                        $sessions += [PSCustomObject]@{
                            Username = $username
                            Server = $server
                            SessionID = $sessionId
                            State = $state
                            DiskSpace = $diskSpace[$server]
                        }
                        $userCounts[$server]++
                    }
                }
            }
            catch {
                Write-Warning "Error querying $server : $_"
            }
        }
    }
   
    return [PSCustomObject]@{
        Sessions = $sessions
        UserCounts = $userCounts
    }
}

# Function to update ListView
function Update-SessionList {
    param (
        [string]$SearchText = "",
        [bool]$Refresh = $false,
        [System.Windows.Forms.ProgressBar]$ProgressBar = $null,
        [System.Windows.Forms.Label]$ProgressLabel = $null
    )
   
    if ($Refresh) {
        $listView.Items.Clear()
       
        # Get servers from textbox or file
        if ($textBoxServers.Text.Trim()) {
            $servers = $textBoxServers.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        }
        else {
            # Try to read from .\servers.txt
            $serverFilePath = ".\servers.txt"
            if (Test-Path $serverFilePath) {
                $servers = Get-Content $serverFilePath | Where-Object { $_.Trim() } | ForEach-Object { $_.Trim() }
            }
            else {
                [System.Windows.Forms.MessageBox]::Show(
                    "No servers specified and couldn't find .\servers.txt",
                    "Warning",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                $servers = @()
            }
        }
       
        if ($servers) {
            $result = Get-RemoteSessions -ServerList $servers -ProgressBar $ProgressBar -ProgressLabel $ProgressLabel
            $global:CachedSessions = $result.Sessions
            $global:UserCounts = $result.UserCounts
        }
    }
   
    $sessions = $global:CachedSessions
   
    if ($SearchText) {
        $searchText = $SearchText.ToLower()
        $sessions = $sessions | Where-Object {
            $_.Username.ToLower().Contains($searchText) -or
            $_.DisplayName.ToLower().Contains($searchText) -or
            $_.Server.ToLower().Contains($searchText) -or
            $_.SessionID.ToString().Contains($searchText) -or
            $_.State.ToLower().Contains($searchText)
        }
    }

    # Sort sessions based on the current sort column and order
    if ($global:SortOrder -eq "Ascending") {
        $sessions = $sessions | Sort-Object -Property $global:SortColumn
    }
    else {
        $sessions = $sessions | Sort-Object -Property $global:SortColumn -Descending
    }
   
    $listView.Items.Clear()
    foreach ($session in $sessions) {
        $item = New-Object System.Windows.Forms.ListViewItem($session.Username)
        $item.SubItems.Add($session.DisplayName)
        $item.SubItems.Add($session.Server)
        $item.SubItems.Add($session.SessionID)
        $item.SubItems.Add($session.State)
        $item.SubItems.Add($session.DiskSpace)
        $item.SubItems.Add($global:UserCounts[$session.Server])
        $listView.Items.Add($item)
    }
}

# Handle the ColumnClick event to sort the ListView
$listView.add_ColumnClick({
    param($sender, $e)
    $columns = @("Username", "DisplayName", "Server", "SessionID", "State", "DiskSpace", "UserCounts")
    $global:SortColumn = $columns[$e.Column]
    if ($global:SortOrder -eq "Ascending") {
        $global:SortOrder = "Descending"
    }
    else {
        $global:SortOrder = "Ascending"
    }
    Update-SessionList -SearchText $textBoxSearch.Text
})

# Create context menu
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$shadowMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$shadowMenuItem.Text = "Shadow Session"
$contextMenu.Items.Add($shadowMenuItem)

$logoffMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$logoffMenuItem.Text = "Log Off Session(s)"
$contextMenu.Items.Add($logoffMenuItem)

$messageMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$messageMenuItem.Text = "Send Message"
$contextMenu.Items.Add($messageMenuItem)

$listView.ContextMenuStrip = $contextMenu

# Function to get disk space
function Get-RemoteDiskSpace {
    param (
        [string]$Server
    )
   
    try {
        $drive = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $Server -Filter "DeviceID='C:'" -ErrorAction Stop
        $freeSpace = [math]::Round($drive.FreeSpace / 1GB, 2)
        return "$freeSpace GB"
    }
    catch {
        return "Error"
    }
}

# Function to read settings from ini file
function Read-Settings {
    $settings = @{
        RefreshOnStartup = 0
        DefaultShadowOptions = 0
        UseMultithreading = 0
    }
    if (Test-Path ".\settings.ini") {
        $ini = Get-Content ".\settings.ini" | Out-String
        $ini -split "`r`n" | ForEach-Object {
            if ($_ -match "RefreshOnStartup=(\d)") {
                $settings.RefreshOnStartup = [int]$matches[1]
            }
            elseif ($_ -match "DefaultShadowOptions=(\d)") {
                $settings.DefaultShadowOptions = [int]$matches[1]
            }
            elseif ($_ -match "UseMultithreading=(\d)") {
                $settings.UseMultithreading = [int]$matches[1]
            }
        }
    }
    return $settings
}

# Function to write settings to ini file
function Write-Settings {
    param (
        [hashtable]$settings
    )
    $content = "[Settings]`nRefreshOnStartup=$($settings.RefreshOnStartup)`nDefaultShadowOptions=$($settings.DefaultShadowOptions)`nUseMultithreading=$($settings.UseMultithreading)"
    Set-Content ".\settings.ini" -Value $content
}

# Function to create settings form
function Show-SettingsForm {
    $settings = Read-Settings

    $settingsForm = New-Object System.Windows.Forms.Form
    $settingsForm.Text = "Settings"
    $settingsForm.Size = New-Object System.Drawing.Size(280,200)
    $settingsForm.StartPosition = "CenterScreen"
    $settingsForm.icon = ".\Images\settings.ico"
    $settingsForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $settingsForm.maximizeBox = $false
    $settingsForm.minimizeBox = $false
    $settingsForm.TopMost = $true

    $checkBoxRefreshOnStartup = New-Object System.Windows.Forms.CheckBox
    $checkBoxRefreshOnStartup.Location = New-Object System.Drawing.Point(10,20)
    $checkBoxRefreshOnStartup.Size = New-Object System.Drawing.Size(260,20)
    $checkBoxRefreshOnStartup.Text = "Refresh on startup"
    $checkBoxRefreshOnStartup.Checked = [bool]$settings.RefreshOnStartup
    $settingsForm.Controls.Add($checkBoxRefreshOnStartup)

    $checkBoxDefaultShadowOptions = New-Object System.Windows.Forms.CheckBox
    $checkBoxDefaultShadowOptions.Location = New-Object System.Drawing.Point(10,50)
    $checkBoxDefaultShadowOptions.Size = New-Object System.Drawing.Size(260,20)
    $checkBoxDefaultShadowOptions.Text = "Default shadow options"
    $checkBoxDefaultShadowOptions.Checked = [bool]$settings.DefaultShadowOptions
    $settingsForm.Controls.Add($checkBoxDefaultShadowOptions)

    $checkBoxUseMultithreading = New-Object System.Windows.Forms.CheckBox
    $checkBoxUseMultithreading.Location = New-Object System.Drawing.Point(10,80)
    $checkBoxUseMultithreading.Size = New-Object System.Drawing.Size(260,20)
    $checkBoxUseMultithreading.Text = "Use multithreading for server scans"
    $checkBoxUseMultithreading.Checked = [bool]$settings.UseMultithreading
    $settingsForm.Controls.Add($checkBoxUseMultithreading)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(50,120)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $settingsForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(150,120)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $settingsForm.Controls.Add($cancelButton)

    $settingsForm.AcceptButton = $okButton
    $settingsForm.CancelButton = $cancelButton

    $result = $settingsForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $settings.RefreshOnStartup = [int]$checkBoxRefreshOnStartup.Checked
        $settings.DefaultShadowOptions = [int]$checkBoxDefaultShadowOptions.Checked
        $settings.UseMultithreading = [int]$checkBoxUseMultithreading.Checked
        Write-Settings -settings $settings
    }
}

# Event handler for settings button click
$buttonSettings.Add_Click({
    Show-SettingsForm
})

# Event handlers
$buttonRefresh.Add_Click({
    $progressForm, $progressBar, $progressLabel = Show-ProgressForm
    $progressForm.Show()
    Update-SessionList -SearchText $textBoxSearch.Text -Refresh $true -ProgressBar $progressBar -ProgressLabel $progressLabel
    $progressForm.Close()
})

$textBoxSearch.Add_TextChanged({
    Update-SessionList -SearchText $textBoxSearch.Text
})

$shadowMenuItem.Add_Click({
    if ($listView.SelectedItems.Count -eq 1) {
        $server = $listView.SelectedItems[0].SubItems[2].Text
        $sessionId = $listView.SelectedItems[0].SubItems[3].Text
        $username = $listView.SelectedItems[0].SubItems[0].Text

        $settings = Read-Settings

        # Create shadow options form
        $optionsForm = New-Object System.Windows.Forms.Form
        $optionsForm.Text = "Shadow Options"
        $optionsForm.Size = New-Object System.Drawing.Size(280,160)
        $optionsForm.StartPosition = "CenterScreen"
        $optionsForm.icon = ".\Images\shadow.ico"
        $optionsForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $optionsForm.maximizeBox = $false
        $optionsForm.minimizeBox = $false
        $optionsForm.TopMost = $true

        $consentCheck = New-Object System.Windows.Forms.CheckBox
        $consentCheck.Location = New-Object System.Drawing.Point(10,20)
        $consentCheck.Size = New-Object System.Drawing.Size(260,20)
        $consentCheck.Text = "No consent prompt"
        $consentCheck.Checked = [bool]$settings.DefaultShadowOptions
        $optionsForm.Controls.Add($consentCheck)

        $controlCheck = New-Object System.Windows.Forms.CheckBox
        $controlCheck.Location = New-Object System.Drawing.Point(10,50)
        $controlCheck.Size = New-Object System.Drawing.Size(260,20)
        $controlCheck.Text = "Enable control (unchecked = view only)"
        $controlCheck.Checked = [bool]$settings.DefaultShadowOptions
        $optionsForm.Controls.Add($controlCheck)

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(50,90)
        $okButton.Size = New-Object System.Drawing.Size(75,23)
        $okButton.Text = "Connect"
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $optionsForm.Controls.Add($okButton)

        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(150,90)
        $cancelButton.Size = New-Object System.Drawing.Size(75,23)
        $cancelButton.Text = "Cancel"
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $optionsForm.Controls.Add($cancelButton)

        $optionsForm.AcceptButton = $okButton
        $optionsForm.CancelButton = $cancelButton

        $result = $optionsForm.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $shadowArgs = "/v:$server /shadow:$sessionId"

            if ($controlCheck.Checked) {
                $shadowArgs += " /control"
            }

            if ($consentCheck.Checked) {
                $shadowArgs += " /noconsentprompt"
            }

            Start-Process "mstsc.exe" -ArgumentList $shadowArgs
        }
    }
})

$logoffMenuItem.Add_Click({
    if ($listView.SelectedItems.Count -ge 1) {
        $selectedSessions = @()
        foreach ($item in $listView.SelectedItems) {
            $selectedSessions += [PSCustomObject]@{
                Username = $item.Text
                Server = $item.SubItems[2].Text
                SessionID = $item.SubItems[3].Text
            }
        }
       
        $message = "Log off the following sessions?`n`n"
        $selectedSessions | ForEach-Object {
            $message += "Server: $($_.Server), User: $($_.Username)`n"
        }
       
        $result = [System.Windows.Forms.MessageBox]::Show(
            $message,
            "Confirm Logoff",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
           
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            foreach ($session in $selectedSessions) {
                try {
                    logoff $session.SessionID /server:$($session.Server)
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Error logging off session for $($session.Username) on $($session.Server): $_",
                        "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
            Update-SessionList -SearchText $textBoxSearch.Text
        }
    }
})

$messageMenuItem.Add_Click({
    if ($listView.SelectedItems.Count -ge 1) {
        # Create message input form
        $msgForm = New-Object System.Windows.Forms.Form
        $msgForm.Text = "Send Message"
        $msgForm.Size = New-Object System.Drawing.Size(400,200)
        $msgForm.StartPosition = "CenterScreen"
        $msgForm.icon = ".\Images\message.ico"
        $msgForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $msgForm.maximizeBox = $false
        $msgForm.minimizeBox = $false
        $msgForm.TopMost = $true
       
        $msgLabel = New-Object System.Windows.Forms.Label
        $msgLabel.Location = New-Object System.Drawing.Point(10,20)
        $msgLabel.Size = New-Object System.Drawing.Size(380,20)
        $msgLabel.Text = "Enter message to send:"
        $msgForm.Controls.Add($msgLabel)
       
        $msgTextBox = New-Object System.Windows.Forms.TextBox
        $msgTextBox.Location = New-Object System.Drawing.Point(10,40)
        $msgTextBox.Size = New-Object System.Drawing.Size(360,80)
        $msgTextBox.Multiline = $true
        $msgForm.Controls.Add($msgTextBox)
       
        $sendButton = New-Object System.Windows.Forms.Button
        $sendButton.Location = New-Object System.Drawing.Point(200,130)
        $sendButton.Size = New-Object System.Drawing.Size(75,23)
        $sendButton.Text = "Send"
        $sendButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $msgForm.Controls.Add($sendButton)
       
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(290,130)
        $cancelButton.Size = New-Object System.Drawing.Size(75,23)
        $cancelButton.Text = "Cancel"
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $msgForm.Controls.Add($cancelButton)
       
        $msgForm.AcceptButton = $sendButton
        $msgForm.CancelButton = $cancelButton
       
        $result = $msgForm.ShowDialog()
       
        if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $msgTextBox.Text.Trim()) {
            $message = $msgTextBox.Text.Trim()
           
            foreach ($item in $listView.SelectedItems) {
                $server = $item.SubItems[2].Text
                $username = $item.SubItems[0].Text
                try {
                    $fullMessage = "Message from IT Support:`n`n$message"
                    msg $username /server:$server $fullMessage
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Error sending message to $username on $server : $_",
                        "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        }
    }
})

# Read settings on startup
$settings = Read-Settings

# Check RefreshOnStartup setting
if ($settings.RefreshOnStartup -eq 1) {
    $progressForm, $progressBar, $progressLabel = Show-ProgressForm
    $progressForm.Show()
    Update-SessionList -SearchText $textBoxSearch.Text -Refresh $true -ProgressBar $progressBar -ProgressLabel $progressLabel
    $progressForm.Close()
}

# Event handler for form key down
$form.add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $buttonRefresh.PerformClick()
    }
})

# Ensure the form can capture key events
$form.KeyPreview = $true

# Show the form
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()

# Release the form
$form.Dispose()