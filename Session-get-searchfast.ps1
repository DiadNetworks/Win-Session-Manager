Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Global variables for caching
$global:SessionCache = @()
$global:FullNameCache = @{}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Remote Session Manager"
$form.Size = New-Object System.Drawing.Size(900,600)
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
$buttonRefresh.Location = New-Object System.Drawing.Point(630,40)
$buttonRefresh.Size = New-Object System.Drawing.Size(75,20)
$buttonRefresh.Text = "Refresh"
$form.Controls.Add($buttonRefresh)

# Create ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10,70)
$listView.Size = New-Object System.Drawing.Size(865,480)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true

# Add columns
$listView.Columns.Add("Username", 120)
$listView.Columns.Add("Full Name", 150)
$listView.Columns.Add("Server", 120)
$listView.Columns.Add("Session ID", 80)
$listView.Columns.Add("State", 80)
$listView.Columns.Add("C: Space Free", 100)
$form.Controls.Add($listView)

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

# Function to get full name using ADSI Searcher with caching
function Get-UserFullName {
    param (
        [string]$Username
    )
    if ($global:FullNameCache.ContainsKey($Username)) {
        return $global:FullNameCache[$Username]
    }
    try {
        $searcher = [adsisearcher]"(&(objectCategory=user)(samaccountname=$Username))"
        $result = $searcher.FindOne()
        $fullName = if ($result) { $result.Properties['displayname'] | Select-Object -First 1 } else { "Not Found" }
        $global:FullNameCache[$Username] = $fullName
        return $fullName
    }
    catch {
        $global:FullNameCache[$Username] = "Not Found"
        return "Not Found"
    }
}

# Function to get sessions (called only when refreshing)
function Get-RemoteSessions {
    param (
        [string[]]$ServerList
    )
   
    $sessions = @()
    $diskSpace = @{}
   
    foreach ($server in $ServerList) {
        $diskSpace[$server] = Get-RemoteDiskSpace -Server $server
       
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
                    $fullName = Get-UserFullName -Username $username
                   
                    $sessions += [PSCustomObject]@{
                        Username = $username
                        FullName = $fullName
                        Server = $server
                        SessionID = $sessionId
                        State = $state
                        DiskSpace = $diskSpace[$server]
                    }
                }
            }
        }
        catch {
            Write-Warning "Error querying $server : $_"
        }
    }
   
    return $sessions | Sort-Object Username
}

# Function to update ListView from cache
function Update-SessionList {
    param (
        [string]$SearchText = ""
    )
   
    $listView.Items.Clear()
   
    if (-not $global:SessionCache) {
        # Initial population of cache if empty
        if ($textBoxServers.Text.Trim()) {
            $servers = $textBoxServers.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        }
        else {
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
            $global:SessionCache = Get-RemoteSessions -ServerList $servers
        }
    }
   
    $filteredSessions = $global:SessionCache
    if ($SearchText) {
        $searchText = $SearchText.ToLower()
        $filteredSessions = $filteredSessions | Where-Object {
            $_.Username.ToLower().Contains($searchText) -or
            $_.FullName.ToLower().Contains($searchText) -or
            $_.Server.ToLower().Contains($searchText) -or
            $_.SessionID.ToString().Contains($searchText) -or
            $_.State.ToLower().Contains($searchText)
        }
    }
   
    foreach ($session in $filteredSessions) {
        $item = New-Object System.Windows.Forms.ListViewItem($session.Username)
        $item.SubItems.Add($session.FullName)
        $item.SubItems.Add($session.Server)
        $item.SubItems.Add($session.SessionID)
        $item.SubItems.Add($session.State)
        $item.SubItems.Add($session.DiskSpace)
        $listView.Items.Add($item)
    }
}

# Debounce timer for search
$debounceTimer = New-Object System.Windows.Forms.Timer
$debounceTimer.Interval = 300 # 300ms delay
$debounceTimer.Add_Tick({
    $debounceTimer.Stop()
    Update-SessionList -SearchText $textBoxSearch.Text
})

# Event handlers
$buttonRefresh.Add_Click({
    $global:SessionCache = @() # Clear cache to force refresh
    $listView.Items.Clear()
    Update-SessionList -SearchText $textBoxSearch.Text
})

$textBoxSearch.Add_TextChanged({
    $debounceTimer.Stop()
    $debounceTimer.Start()
})

$shadowMenuItem.Add_Click({
    if ($listView.SelectedItems.Count -eq 1) {
        $server = $listView.SelectedItems[0].SubItems[2].Text
        $sessionId = $listView.SelectedItems[0].SubItems[3].Text
        $username = $listView.SelectedItems[0].SubItems[0].Text
        
        $optionsForm = New-Object System.Windows.Forms.Form
        $optionsForm.Text = "Shadow Options"
        $optionsForm.Size = New-Object System.Drawing.Size(300,200)
        $optionsForm.StartPosition = "CenterScreen"
       
        $consentCheck = New-Object System.Windows.Forms.CheckBox
        $consentCheck.Location = New-Object System.Drawing.Point(10,20)
        $consentCheck.Size = New-Object System.Drawing.Size(260,20)
        $consentCheck.Text = "No consent prompt"
        $optionsForm.Controls.Add($consentCheck)
       
        $controlCheck = New-Object System.Windows.Forms.CheckBox
        $controlCheck.Location = New-Object System.Drawing.Point(10,50)
        $controlCheck.Size = New-Object System.Drawing.Size(260,20)
        $controlCheck.Text = "Enable control (unchecked = view only)"
        $optionsForm.Controls.Add($controlCheck)
       
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(100,120)
        $okButton.Size = New-Object System.Drawing.Size(75,23)
        $okButton.Text = "Connect"
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $optionsForm.Controls.Add($okButton)
       
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(190,120)
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
                Username = $item.SubItems[0].Text
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
            $global:SessionCache = @() # Clear cache after logoff
            Update-SessionList -SearchText $textBoxSearch.Text
        }
    }
})

$messageMenuItem.Add_Click({
    if ($listView.SelectedItems.Count -ge 1) {
        $msgForm = New-Object System.Windows.Forms.Form
        $msgForm.Text = "Send Message"
        $msgForm.Size = New-Object System.Drawing.Size(400,200)
        $msgForm.StartPosition = "CenterScreen"
        $msgForm.icon = ".\Images\message.ico"
       
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

# Show the form
$form.Add_Shown({
    $form.Activate()
    Update-SessionList # Initial population
})
$form.ShowDialog()

# Release the form
$form.Dispose()
