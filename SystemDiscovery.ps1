Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Clear-Host

# === Setup output and logging directories ===
$scriptRoot = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$parentDir  = Split-Path -Path $scriptRoot -Parent
$outputFolder = Join-Path $parentDir "Discovery"
$logDir = Join-Path $parentDir "logs"
New-Item -Path $outputFolder, $logDir -ItemType Directory -Force | Out-Null
$errorLog = Join-Path $logDir "error.log"
$eventLog = Join-Path $logDir "event.log"

function Log-Error {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [ERROR] $message" | Out-File -Append -FilePath $errorLog
}

function Log-Event {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [INFO] $message" | Out-File -Append -FilePath $eventLog
}

function Check-Cancel {
    if ($cancelRequested) {
        Log-Event "Operation cancelled by user."
        $progressForm.Close()
        exit
    }
}

# === Select Users UI ===
$userFolders = Get-ChildItem -Path "C:\Users" -Directory | Where-Object { $_.Name -notlike "*Default*" -and $_.Name -ne "All Users" }

$formSelect = New-Object Windows.Forms.Form
$formSelect.Text = "Select User Profiles"
$formSelect.Size = '400,350'
$formSelect.StartPosition = 'CenterScreen'
$formSelect.TopMost = $true
$formSelect.FormBorderStyle = 'FixedDialog'
$formSelect.MaximizeBox = $false
$formSelect.MinimizeBox = $false

$clb = New-Object Windows.Forms.CheckedListBox
$clb.Location = '20,20'
$clb.Size = '340,230'
$clb.CheckOnClick = $true
$userFolders.ForEach({ [void]$clb.Items.Add($_.Name, $true) })
$formSelect.Controls.Add($clb)

$btnOK = New-Object Windows.Forms.Button
$btnOK.Text = "OK"
$btnOK.Size = '100,30'
$btnOK.Location = '140,260'
$btnOK.Add_Click({ $formSelect.Close() })
$formSelect.Controls.Add($btnOK)

$formSelect.ShowDialog()
$selectedUsers = $clb.CheckedItems
if ($selectedUsers.Count -eq 0) {
    exit
}

# === Progress UI ===
$progressForm = New-Object Windows.Forms.Form
$progressForm.Text = "Discovery In Progress"
$progressForm.Size = New-Object Drawing.Size(400,160)
$progressForm.StartPosition = "CenterScreen"
$progressForm.TopMost = $true
$progressForm.FormBorderStyle = 'FixedDialog'
$progressForm.MaximizeBox = $false
$progressForm.MinimizeBox = $false

$progressBar = New-Object Windows.Forms.ProgressBar
$progressBar.Location = '20,30'
$progressBar.Size = '350,20'
$progressBar.Minimum = 0
$progressBar.Maximum = 8
$progressForm.Controls.Add($progressBar)

$lblStatus = New-Object Windows.Forms.Label
$lblStatus.AutoSize = $true
$lblStatus.Location = '20, 60'
$lblStatus.Text = "Starting..."
$progressForm.Controls.Add($lblStatus)

$btnCancel = New-Object Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Size = '80,30'
$btnCancel.Location = '150,90'
$progressForm.Controls.Add($btnCancel)

$cancelRequested = $false
$btnCancel.Add_Click({
    $cancelRequested = $true
    $lblStatus.Text = "Cancelling..."
})

$null = $progressForm.Show()

# === Setup paths ===
$modulePath = Join-Path (Split-Path $scriptRoot -Parent) "Modules\ImportExcel"
$pcName = $env:COMPUTERNAME
$timestamp = Get-Date -Format "MM_dd_yy_HHmm"
$excelFile = Join-Path $outputFolder "System_Discovery_${pcName}_$timestamp.xlsx"

# === ImportExcel Check ===
$importExcelAvailable = $false
if (Test-Path "$modulePath\ImportExcel.psd1") {
    Import-Module $modulePath -Force -ErrorAction SilentlyContinue
    $importExcelAvailable = $true
    Log-Event "ImportExcel module loaded."
} else {
    Log-Event "ImportExcel not found. Exporting to CSV."
}

# === SYSTEM INFO ===
Check-Cancel
$lblStatus.Text = "Collecting system info..."
$progressBar.Value = 1
$progressForm.Refresh()
try {
    $sysInfo = [PSCustomObject]@{
        ComputerName = $env:COMPUTERNAME
        Domain       = (Get-CimInstance Win32_ComputerSystem).Domain
        CurrentUser  = $env:USERNAME
        OS           = (Get-CimInstance Win32_OperatingSystem).Caption
        CPU          = (Get-CimInstance Win32_Processor).Name
        RAM_GB       = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 1)
        TimeZone     = (Get-TimeZone).DisplayName
        Uptime       = ((Get-CimInstance Win32_OperatingSystem).LastBootUpTime).ToLocalTime()
    }
    Log-Event "System info collected."
} catch { Log-Error "System Info: $_" }
Start-Sleep -Milliseconds 300

# === NETWORK ADAPTERS ===
Check-Cancel
$lblStatus.Text = "Collecting network adapters..."
$progressBar.Value = 2
$progressForm.Refresh()
try {
    $netAdapters = Get-NetIPConfiguration | Where-Object { $_.NetAdapter.Status -eq 'Up' } | ForEach-Object {
        [PSCustomObject]@{
            Name         = $_.NetAdapter.InterfaceAlias
            Status       = $_.NetAdapter.Status
            MACAddress   = $_.NetAdapter.MacAddress
            IPv4         = ($_.IPv4Address | ForEach-Object { $_.IPAddress }) -join ', '
            IPv6         = ($_.IPv6Address | ForEach-Object { $_.IPAddress }) -join ', '
            PrefixLength = ($_.IPv4Address | ForEach-Object { $_.PrefixLength }) -join ', '
            Gateway      = ($_.IPv4DefaultGateway | ForEach-Object { $_.NextHop }) -join ', '
            DNS          = ($_.DNSServer.ServerAddresses) -join ', '
            DHCPEnabled  = $_.NetAdapter.DhcpEnabled
            IPAssignment = if ($_.NetAdapter.DhcpEnabled) { "DHCP" } else { "Static" }
            Interface    = $_.NetAdapter.InterfaceDescription
        }
    }
    Log-Event "Network adapters collected."
} catch { Log-Error "Network Info: $_" }
Start-Sleep -Milliseconds 300

# === MAPPED DRIVES ===
Check-Cancel
$lblStatus.Text = "Collecting mapped drives..."
$progressBar.Value = 3
$progressForm.Refresh()
$mappedDrives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot } | ForEach-Object {
    [PSCustomObject]@{
        DriveLetter = $_.Name
        RemotePath  = $_.DisplayRoot
        Used_MB     = [math]::Round($_.Used / 1MB, 1)
        Free_MB     = [math]::Round($_.Free / 1MB, 1)
    }
}
Log-Event "Mapped drives collected."
Start-Sleep -Milliseconds 300

# === PRINTERS ===
Check-Cancel
$lblStatus.Text = "Collecting printers..."
$progressBar.Value = 4
$progressForm.Refresh()
$printers = Get-CimInstance Win32_Printer | ForEach-Object {
    [PSCustomObject]@{
        Name     = $_.Name
        Port     = $_.PortName
        Default  = $_.Default
        Network  = $_.Network
        Status   = $_.PrinterStatus
    }
}
Log-Event "Printers collected."
Start-Sleep -Milliseconds 300

# === INSTALLED APPS ===
Check-Cancel
$lblStatus.Text = "Collecting installed apps..."
$progressBar.Value = 5
$progressForm.Refresh()
$excludePublishers = @("Microsoft Corporation", "Microsoft", "Microsoft Windows", "Windows Defender")
$apps = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*, `
                       HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* `
    -ErrorAction SilentlyContinue | Where-Object {
        $_.DisplayName -and ($_.Publisher -notin $excludePublishers)
    } | Sort-Object DisplayName | ForEach-Object {
        [PSCustomObject]@{
            AppName   = $_.DisplayName
            Version   = $_.DisplayVersion
            Publisher = $_.Publisher
        }
    }
Log-Event "Installed apps collected."
Start-Sleep -Milliseconds 300

# === USER PROFILE SIZES ===
Check-Cancel
$lblStatus.Text = "Calculating user profile sizes..."
$progressBar.Value = 6
$progressForm.Refresh()
$userProfiles = foreach ($user in $selectedUsers) {
    $profilePath = Join-Path "C:\Users" $user
    $size = 0
    try {
        $size = (Get-ChildItem -Path $profilePath -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    } catch {
        Log-Error ("Profile Size " + $user + ": " + $_)
    }

    [PSCustomObject]@{
        Profile = $user
        Path    = $profilePath
        Size_GB = [math]::Round($size / 1GB, 2)
    }
}
Log-Event "User profile sizes calculated."
Start-Sleep -Milliseconds 300

# === OUTLOOK CONFIG (COM method) ===
Check-Cancel
$lblStatus.Text = "Retrieving Outlook account config..."
$progressBar.Value = 7
$progressForm.Refresh()
$outlookAccounts = @()
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $namespace.Accounts | ForEach-Object {
        $outlookAccounts += [PSCustomObject]@{
            DisplayName   = $_.DisplayName
            SmtpAddress   = $_.SmtpAddress
            AccountType   = $_.AccountType
            UserName      = $_.UserName
            Server        = $_.DeliveryStore.DisplayName
        }
    }
    Log-Event "Outlook profiles collected."
} catch {
    $outlookAccounts += [PSCustomObject]@{ DisplayName = "Outlook not installed or no profiles"; SmtpAddress = ""; AccountType = ""; UserName = ""; Server = "" }
    Log-Error "Outlook Info: $_"
}
Start-Sleep -Milliseconds 300

# === EXPORT ===
Check-Cancel
$lblStatus.Text = "Exporting data..."
$progressBar.Value = 8
$progressForm.Refresh()
try {
    if ($importExcelAvailable) {
        Export-Excel -Path $excelFile -WorksheetName "SystemInfo"      -AutoSize -FreezeTopRow -BoldTopRow -ClearSheet -InputObject $sysInfo
        Export-Excel -Path $excelFile -WorksheetName "NetworkAdapters" -AutoSize -FreezeTopRow -BoldTopRow -Append -InputObject $netAdapters
        Export-Excel -Path $excelFile -WorksheetName "MappedDrives"    -AutoSize -FreezeTopRow -BoldTopRow -Append -InputObject $mappedDrives
        Export-Excel -Path $excelFile -WorksheetName "Printers"        -AutoSize -FreezeTopRow -BoldTopRow -Append -InputObject $printers
        Export-Excel -Path $excelFile -WorksheetName "InstalledApps"   -AutoSize -FreezeTopRow -BoldTopRow -Append -InputObject $apps
        Export-Excel -Path $excelFile -WorksheetName "UserProfiles"    -AutoSize -FreezeTopRow -BoldTopRow -Append -InputObject $userProfiles
        Export-Excel -Path $excelFile -WorksheetName "OutlookAccounts" -AutoSize -FreezeTopRow -BoldTopRow -Append -InputObject $outlookAccounts
        Log-Event "Exported to Excel: $excelFile"
    } else {
        $csvBase = Join-Path $outputFolder "SystemDiscovery_$pcName"
        $sysInfo        | Export-Csv "$csvBase`_SystemInfo.csv" -NoTypeInformation
        $netAdapters    | Export-Csv "$csvBase`_Network.csv" -NoTypeInformation
        $mappedDrives   | Export-Csv "$csvBase`_MappedDrives.csv" -NoTypeInformation
        $printers       | Export-Csv "$csvBase`_Printers.csv" -NoTypeInformation
        $apps           | Export-Csv "$csvBase`_InstalledApps.csv" -NoTypeInformation
        $userProfiles   | Export-Csv "$csvBase`_UserProfiles.csv" -NoTypeInformation
        $outlookAccounts| Export-Csv "$csvBase`_OutlookAccounts.csv" -NoTypeInformation
        Log-Event "Exported to CSV files."
    }
} catch { Log-Error "Export failed: $_" }

# === Finish ===
$lblStatus.Text = "Discovery complete!"
$progressForm.Refresh()
Log-Event "Discovery finished."
[System.Windows.Forms.MessageBox]::Show($progressForm, "Process complete. Check Output and Logs folder for details.") | Out-Null
Start-Sleep -Milliseconds 1000
$progressForm.Close()
