<#
Locate Problem Devices in Vocera VMP logs
Philip Otter 2024
#>

Add-Type -AssemblyName System.Windows.Forms

$logPathList = @("./Logs")
$unknownFiles = [System.Collections.ArrayList]::new()
$logFiles = [System.Collections.ArrayList]::new()
$errorFiles = [System.Collections.ArrayList]::new()
$Form = [System.Windows.Forms.Form]::new()
$applicationVersion = '0.1.0'
$authorStamp = 'Philip Otter 2024'


class VoceraUserAuthentications{
    [string]    $RequestDate
    [string]    $RequestTime
    [string]    $requestID
    [string]    $UserID
    [bool]      $UserAuthenticated
}


class VoceraUser{
    [string]    $UserName
    [string]    $VoceraID = "N/A"
    [string]    $UserID

    $authentications = [System.Collections.ArrayList]@()
}


class VoceraDeviceList{
    $deviceList = [System.Collections.ArrayList]::new()
    $macList = [System.Collections.ArrayList]::new()
    $userList = [System.Collections.ArrayList]::new()
    $userNameList = [System.Collections.ArrayList]::new()
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
    $voceraUserTraits = @('Name','Voice ID','VMP ID','Auth Count')

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

    # Found Files Tab
    $foundFilesTab = New-Object System.Windows.Forms.TabPage
    $foundFilesTab.TabIndex = 1
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

    # Device Tab
    $deviceTab = New-Object System.Windows.Forms.TabPage
    $deviceTab.TabIndex = 2
    $deviceTab.Text = 'Devices'

    $clientMACBox = New-Object System.Windows.Forms.ListBox
    $clientMACBox.Location = New-Object System.Drawing.Point(10,40)
    $clientMACBox.Height = 300
    $clientMACBox.Width = 125
    $clientMACBox.Sorted = $true

    foreach($mac in $AllDevices.macList){
        $clientMACBox.Items.Add($mac)
    }

    # Button to load Device
    $loadDeviceButton = New-Object System.Windows.Forms.Button
    $loadDeviceButton.Location = New-Object System.Drawing.Point(10,340)
    $loadDeviceButton.Height = 20
    $loadDeviceButton.Width = 125
    $loadDeviceButton.TextAlign
    $loadDeviceButton.Text = 'Load Device'
    $loadDeviceButton.Add_Click({
        $propertyArray= @('ClientType','ClientVersion','MAC','LogFrequency')
        Write-Host "Clicked"
        $selectedMAC = $clientMACBox.SelectedItem
        Write-Host "Selected MAC:  $selectedMAC"
        foreach($device in $AllDevices.deviceList){
            Write-Host "Comparing $selectedMac to " $device.MAC
            if($selectedMac -eq $device.MAC){
                Write-Host "Match Found!" $device.MAC
                $deviceDetailsLabelOffsetValue = 100
                $device.PSObject.Properties| ForEach-Object{
                    if($propertyArray -Contains $_.Name){
                        Write-Host "Writing Trait:  " $_.Name
                        # Check for and remove old information already loaded for a different device
                        if($deviceTab.Controls.ContainsKey($_.Name)){
                            Write-Host "Removing old label data" -ForegroundColor Green
                            $deviceTab.Controls.RemoveByKey($_.Name)
                        }
                        else{
                            $key = $_.Name
                            Write-Host "No Key Found $key" -BackgroundColor Red
                        }
                        $label = New-Object System.Windows.Forms.Label
                        $label.Text = $_.Value
                        $label.Name = $_.Name
                        $label.Location = New-Object System.Drawing.Point(350,$deviceDetailsLabelOffsetValue)
                        $deviceTab.Controls.Add($label)
                        $deviceDetailsLabelOffsetValue = $deviceDetailsLabelOffsetValue + 50
                        $form.Refresh()
                    }
                }
                break
            }
        }
    })

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

    # Users Tab
    $usersTab = New-Object System.Windows.Forms.tabPage
    $usersTab.TabIndex = 3
    $usersTab.Text = 'Users'

    $usersBox = New-Object System.Windows.Forms.ListBox
    $usersBox.Location = New-Object System.Drawing.Point(10,40)
    $usersBox.Height = 300
    $usersBox.Width = 175
    $usersBox.Sorted = $true
    

    foreach($userName in $AllDevices.userNameList){
        $usersBox.Items.Add($userName)
    }

     # Button to load User
     $loadUserButton = New-Object System.Windows.Forms.Button
     $loadUserButton.Location = New-Object System.Drawing.Point(10,340)
     $loadUserButton.Height = 20
     $loadUserButton.Width = 175
     $loadUserButton.TextAlign
     $loadUserButton.Text = 'Load User'
     $loadUserButton.Add_Click({
         $propertyArray= @('UserName','VoceraID','UserID','authentications')
         Write-Host "Clicked"
         $selectedUser = $usersBox.SelectedItem
         Write-Host "Selected User ID:  $selectedUser"
         foreach($voceraUser in $AllDevices.userList){
             Write-Host "Comparing $selectedUser to " $voceraUser.UserName
             if($selectedUser -eq $voceraUser.UserName){
                 Write-Host "Match Found!" $voceraUser.UserName
                 $deviceDetailsLabelOffsetValue = 100
                 $voceraUser.PSObject.Properties| ForEach-Object{
                     if($propertyArray -Contains $_.Name){
                         Write-Host "Writing Trait:  " $_.Name
                         # Check for and remove old information already loaded for a different device
                         if($usersTab.Controls.ContainsKey($_.Name)){
                             Write-Host "Removing old label data" -ForegroundColor Green
                             $usersTab.Controls.RemoveByKey($_.Name)
                         }
                         else{
                             $key = $_.Name
                             Write-Host "No Key Found $key" -BackgroundColor Red
                         }
                         $label = New-Object System.Windows.Forms.Label
                         $label.Name = $_.Name
                         $label.Location = New-Object System.Drawing.Point(350,$deviceDetailsLabelOffsetValue)
                         $deviceDetailsLabelOffsetValue = $deviceDetailsLabelOffsetValue + 50
                         if($_.Name -eq 'authentications'){
                            $label.Text = $_.Length
                         }else{
                            $label.Text = $_.Value
                         }
                         $usersTab.Controls.Add($label)
                         $form.Refresh()
                     }
                 }
                break
             }
         }
     })

    # Labels for User traits
    $userDetailsLabelOffsetValue = 100
    foreach($trait in $voceraUserTraits){
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $trait+":"
        $label.Location = New-Object System.Drawing.Point(250,$userDetailsLabelOffsetValue)
        $usersTab.Controls.Add($label)
        $userDetailsLabelOffsetValue = $userDetailsLabelOffsetValue + 50
    }

    # Add Elements to Users Tab
    $usersTab.Controls.AddRange(@($usersBox,$loadUserButton))

    # Add tabpages to tab control
    $tabControl.Controls.AddRange(@($foundFilesTab, $deviceTab, $usersTab))
    
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
    $deviceDetailsList = [System.Collections.ArrayList]::new()
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
        
        if($AllDevices.macList -notcontains $mac){
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


function Get-VocerLogUsers(){
    $userLineList = [System.Collections.ArrayList]::new()
    foreach($file in $logFiles){
        Get-Content $file | Select-String "user Found," | ForEach-Object{$userLineList.Add($_)}
    }foreach($userLine in $userLineList){
        [regex]$usersNameRegex = "(?<=found[,][ ]name:[ ]).+?(?=[,])"
        [regex]$usersIDRegex = "(?<=[,][ ][ ]id[:][ ]).+?(?=[,])"
        [regex]$usersVoiceIDRegex = "(?<=[,][ ]voice[ ]id[:][ ]).+"

        $voceraUser = [VoceraUser]::new()
        $voceraUser.UserName = $usersNameRegex.Matches($userLine) | ForEach-Object {$_.Value}
        $voceraUser.UserID = $usersIDRegex.Matches($userLine) | ForEach-Object {$_.Value}
        $voceraUser.VoceraID = $usersVoiceIDRegex.Matches($userLine) | ForEach-Object {$_.Value}

        if($AllDevices.userNameList -notcontains $voceraUser.UserName){
            $AllDevices.userNameList.Add($voceraUser.UserName)
            $AllDevices.userList.Add($voceraUser)
        }
    }
}


function Get-UserAuthentications(){
    $authLineList = [System.Collections.ArrayList]::new()
    foreach($file in $logFiles){
        Get-Content $file | Select-String "Request ID:" | ForEach-Object{$authLineList.Add($_)}
    }
    foreach($authLine in $authLineList){
        [regex]$authDateStampRegex = "[0-3][0-9]\/[0-1][0-9]\/[0-9][0-9]"
        [regex]$authTimeStampRegex = "[0-2][0-9]\:[0-6][0-9]\:[0-6][0-9]\..+?(?=[ ])"
        [regex]$requestIDRegex = "(?<=Request[ ]ID[:]).+?(?=[,])"
        [regex]$vmpIDRegex = "(?<=VMP[ ]ID[:]).+?(?=[,])"

        $authenticationInstance = [VoceraUserAuthentications]::new()
        $authenticationInstance.RequestDate = $authDateStampRegex.Matches($authLine) | ForEach-Object {$_.Value}
        $authenticationInstance.RequestTime = $authTimeStampRegex.Matches($authLine) | ForEach-Object {$_.Value}
        $authenticationInstance.requestID = $requestIDRegex.Matches($authLine) | ForEach-Object {$_.Value}
        $authenticationInstance.UserID = $vmpIDRegex.Matches($authLine) | ForEach-Object {$_.Value}

        foreach($user in $AllDevices.userList){
            if($user.UserID -eq $authenticationInstance.UserID){
                $user.authentications.Add($authenticationInstance)
            }
        }
    }

}


$AllDevices = [VoceraDeviceList]::new()
foreach($logPath in $logPathList){
    Get-LogFiles($logPath)
}


Write-Host "Building Devices" -ForegroundColor Cyan
Get-VoceraDevices
Write-Host "Building Users" -ForegroundColor Cyan
Get-VocerLogUsers
Write-Host "Finding Authentications" -ForegroundColor Cyan
Get-UserAuthentications
Write-Host "LAUNCH WORK DONE" -ForegroundColor Green
Open-Window