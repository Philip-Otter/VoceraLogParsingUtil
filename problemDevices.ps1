<#
Locate Problem Devices in Vocera VMP logs
Philip Otter 2024
#>

Add-Type -AssemblyName System.Windows.Forms

$logPathList = @("./Logs")
$unknownFiles = [System.Collections.ArrayList]@()
$logFiles = [System.Collections.ArrayList]@()
$errorFiles = [System.Collections.ArrayList]@()
$deviceDetailsList = [System.Collections.ArrayList]@()
$Form = [System.Windows.Forms.Form]::new()
$applicationVersion = '0.1.0'
$authorStamp = 'Philip Otter 2024'

class VoceraDeviceList{
    $deviceList = [System.Collections.ArrayList]@()
    $macList = [System.Collections.ArrayList]@()
}

class VoceraClientDevice {
    [string]    $ClientType
    [string]    $ClientProto
    [string]    $ClientVersion
    [string]    $MAC
    [string]    $ClientLocalTime
    [int]       $LogFrequency = 1
}


function Open-Window{
    Write-Host "Starting GUI" -backgroundColor gray
    $voceraDeviceTraits = @('Client Type', 'Client Version', 'MAC Address', 'Log Frequency')

    function Draw{
        $form.ShowDialog()
    }

    # Main Form
    $form.Text = "Vocera Log Parser - Version $applicationVersion"
    $form.Width = 600
    $form.Height = 600
    $form.AutoScale = $true

    # Footer
    $footerLabel = New-Object System.Windows.Forms.Label
    $footerLabel.Location = New-Object System.Drawing.Point(230,530)
    $footerLabel.Text = $authorStamp

    # Tabs
    Write-Host "BUILDING TABS" -BackgroundColor Gray
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Width = 560
    $tabControl.Height = 500
    $tabControl.Location = '15,15'
    $tabControl.AutoSize = $true
    $tabControl.Anchor = 'Top,Left,Bottom,Right'

    $deviceTab = New-Object System.Windows.Forms.TabPage
    $deviceTab.TabIndex = 1
    $deviceTab.Text = 'Devices'

    $clientMACBox = New-Object System.Windows.Forms.ListBox
    $clientMACBox.Location = New-Object System.Drawing.Point(10,40)
    #$clientMACBox.Size = New-Object System.Drawing.Size(260,20)
    $clientMACBox.Height = 300
    $clientMACBox.Width = 125
    $clientMACBox.Sorted = $true

    foreach($mac in $AllDevices.macList){
        Write-Host $mac
        $clientMACBox.Items.Add($mac)
    }

    # Button to load Device
    $loadDeviceButton = New-Object System.Windows.Forms.Button
    $loadDeviceButton.Location = New-Object System.Drawing.Point(10,340)
    $loadDeviceButton.Height = 20
    $loadDeviceButton.Width = 125
    $loadDeviceButton.TextAlign = $true
    $loadDeviceButton.Text = 'Load Device'


    # Labels for device traits
    $deviceDetailsLabelOffsetValue = 100
    foreach($trait in $voceraDeviceTraits){
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $trait+":"
        $label.Location = New-Object System.Drawing.Point(225,$deviceDetailsLabelOffsetValue)
        $deviceTab.Controls.Add($label)
        $deviceDetailsLabelOffsetValue = $deviceDetailsLabelOffsetValue + 50
    }

    # Add Elements to Device Tab
    $deviceTab.Controls.AddRange(@($clientMACBox,$loadDeviceButton))

    $foundFilesTab = New-Object System.Windows.Forms.TabPage
    $foundFilesTab.TabIndex = 2
    $foundFilesTab.Text = 'Found Files'

    $foundFilesBox = New-Object System.Windows.Forms.ListBox
    $foundFilesBox.Location = New-Object System.Drawing.Point(10,40)
    $foundFilesBox.Size = New-Object System.Drawing.Size(260,20)
    $foundFilesBox.Height = 300
    $foundFilesBox.Width = 350
    $foundFilesBox.HorizontalScrollbar = $true

    foreach($file in $logFiles){
        $foundFilesBox.Items.Add($file)
    }
    foreach($file in $errorFiles){
        $foundFilesBox.Items.Add($file)
    }

    # Add elements to Found Files tab
    $foundFilesTab.Controls.Add($foundFilesBox)

    # Add tabpages to tab control
    $tabControl.Controls.AddRange(@($deviceTab, $foundFilesTab))
    
    # Add items to main form
    $form.Controls.Add($footerLabel)
    $form.Controls.Add($tabControl)

    Draw

}


function Get-LogFiles($path){
    $list = Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue
    foreach($file in $list){
        # Remove Zip Files
        if ($file -match "[.]zip$"){
            continue
        }elseif ($file -match "[.]log$") {
            $logFiles.Add(($path + "/" + $file))
        }elseif ($file -match "[.]error$"){
            $errorFiles.Add(($path + "/" + $file))
        }else{
            $unknownFiles.Add(($path + "/" + $file))
        }
    }
}


function Get-VoceraDevices(){
    foreach($file in $logFiles){
        Get-Content $file | Select-String "Client details" -CaseSensitive | ForEach-Object{$deviceDetailsList.Add($_)}
    }
    foreach($deviceDetails in $deviceDetailsList){
        [regex]$clientTypeRegex = "(?<=clientType[:]).+?(?=[,])"
        [regex]$clientProtoRegex = "(?<=clientProto[:]).+?(?=[,])"
        [regex]$clientVersionRegex = "(?<=clientVersion[:]).+?(?=[,])"
        [regex]$MACRegex = "(?<=mac[:]).+?(?=[,])"
        [regex]$clientLocalTimeRegex = "(?<=clientLocalTime[:]).+"

        $mac = $MACRegex.Matches($deviceDetails) | ForEach-Object {$_.Value}
        
        if($AllDevices.macList -contains $mac){
            Write-Host -ForegroundColor Yellow $mac " - Appeared Again"
        }else{
            $device = [VoceraClientDevice]::new()
            $device.ClientType = $clientTypeRegex.Matches($deviceDetails) | ForEach-Object {$_.Value}
            $device.ClientProto = $clientProtoRegex.Matches($deviceDetails) | ForEach-Object {$_.Value}
            $device.ClientVersion = $clientVersionRegex.Matches($deviceDetails) | ForEach-Object {$_.Value}
            $device.MAC = $mac
            $device.ClientLocalTime = $clientLocalTimeRegex.Matches($deviceDetails) | ForEach-Object {$_.Value}

            $AllDevices.deviceList.Add($device)
            $AllDevices.macList.Add($device.MAC)
        }
    }
}

$AllDevices = [VoceraDeviceList]::new()
foreach($logPath in $logPathList){
    Get-LogFiles($logPath)
}

# Write-LogFileList
Get-VoceraDevices
Write-Host "Done" -ForegroundColor Red
$AllDevices | Select-Object * | Format-Table -AutoSize
$AllDevices.deviceList[0]
Open-Window