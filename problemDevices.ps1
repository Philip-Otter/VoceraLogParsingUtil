<#
A script to make looking through VMP logs easier
Philip Otter 2024
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Web

$logPathList = @("./Logs")
$unknownFiles = [System.Collections.ArrayList]::new()
$logFiles = [System.Collections.ArrayList]::new()
$errorFiles = [System.Collections.ArrayList]::new()
$Form = [System.Windows.Forms.Form]::new()
$applicationVersion = '0.1.0'
$authorStamp = 'Philip Otter 2024'

class VMPServer{
    [string]    $LastUserCache = "N/A"
    [string]    $LastUserCacheDate = "-/-/-"
    [string]    $LastUserCacheTime = "-:-:-"
    [string]    $LastUserCacheTimeDuration = "--"
    [string]    $LastDistCache = "N/A"
    [string]    $LastDistCacheDate = "-/-/-"
    [string]    $LastDistCacheTime = "-:-:-"
    [string]    $LastDistCacheTimeDuration = "--"
    [string]    $LastHTTPServerStartDate = "-/-/-"
    [string]    $LastHTTPServerStartTime = "-:-:-"
    [string]    $LastHTTPPort = "<-->"
    [string]    $LastHTTPSPort = "<-->"
}

class VoceraUserAuthentications{
    [string]    $RequestDate
    [string]    $RequestTime
    [string]    $requestID
    [string]    $UserID
    [bool]      $UserAuthenticated
}


class VoceraUser{
    [string]    $UserName = "--"
    [string]    $VoceraID = "N/A"
    [string]    $UserID = "--"

    $siteID = [System.Collections.ArrayList]::new()
    $authentications = [System.Collections.ArrayList]::new()
    $associatedDevices = [System.Collections.ArrayList]::new()

    [string]    $UserEmail = "--"
    $deviceIDs = [System.Collections.ArrayList]::new()
}


class ErrorMesssage{
    [string]    $SourceType = "--" # This is either a line or a File
    [string]    $LogType = "--" # INFO, VERBOSE, WARNING
    [string]    $Date = "--/--/--"
    [string]    $Time = "--:--:--"
    [string]    $ErrorCode = "--"
    [string]    $Message = "--"
    [string]    $FullError = "-----"
}


class VoceraDeviceList{
    $deviceList = [System.Collections.ArrayList]::new()
    $macList = [System.Collections.ArrayList]::new()
    $userList = [System.Collections.ArrayList]::new()
    $userNameList = [System.Collections.ArrayList]::new()
    $errorObjects = [System.Collections.ArrayList]::new()
}

class VoceraClientDevice {
    # Get drawn on the form
    [string]    $ClientType = "--" # Houses mobile client OS as well
    [string]    $MobileClientOSVersion = "-.-.-"
    [string]    $ClientProto = "--"
    [string]    $ClientVersion = "-.-.-"
    [string]    $MAC = "--"
    [string]    $MobileClientModel = "--"
    [string]    $SSID = "--"
    [string]    $MobileClientCarrier = "--"
    [int]       $LogFrequency = 1

    # Behind the scenes
    [string]    $MobileClientPIN = "----"
    [bool]      $IsMobileDevice = $false

    $connectionIDList = [System.Collections.ArrayList]::new()
}


function Open-Window{
    Write-Host "Starting GUI" -backgroundColor gray
    $voceraDeviceTraits = @('Client Type', 'Client Version', 'MAC Address', 'Log Frequency')
    $voceraMobileDeviceTraits = @('Client Type', 'OS Version', 'VCS Version', 'MAC Address', 'Model', 'SSID', 'Carrier', 'Log Frequency')
    $voceraUserTraits = @('Name','Voice ID','VMP ID', 'Site ID', 'Auth Count', 'Associated Devices')
    $VMPServerTraitsA = @('Users Cached','User Cache Date','User Cache Time','Caching Time (ms)', 'HTTP Start Date', 'HTTP Start Time')
    $VMPServerTraitsB = @('DLs Cached','DL Cache Date','DL Cache Time','Caching Time (ms)', 'HTTP Port', 'HTTPS Port')

    function Draw{
        $form.ShowDialog()
    }

    # Main Form
    $form.Text = "Vocera VMP Log Parser - Version $applicationVersion"
    $form.Width = 600
    $form.Height = 600
    $form.AutoScale = $true
    $form.StartPosition = "CenterScreen"

    # Footer
    $footerLabel = [System.Windows.Forms.Label]::new()
    $footerLabel.Location = [System.Drawing.Point]::new(230,530)
    $footerLabel.Text = $authorStamp

    # Tabs
    Write-Host "BUILDING TABS" -BackgroundColor Gray
    $tabControl = [System.Windows.Forms.TabControl]::new()
    $tabControl.Width = 555
    $tabControl.Height = 500
    $tabControl.Location = '15,15'
    $tabControl.AutoSize = $true
    $tabControl.Anchor = 'Top,Left,Bottom,Right'

    # Home Tab
    $homeTab = [System.Windows.Forms.TabPage]::new()
    $homeTab.TabIndex = 1
    $homeTab.Text = 'Home'

    $foundFilesBox = [System.Windows.Forms.ListBox]::new()
    $foundFilesBox.Location = [System.Drawing.Point]::new(10,140) 
    $foundFilesBox.Size = [System.Drawing.Size]::new(260,20)
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
    $homeTab.Controls.Add($foundFilesBox)

    # Device Tab
    $deviceTab = [System.Windows.Forms.TabPage]::new()
    $deviceTab.TabIndex = 2
    $deviceTab.Text = 'Devices'

    $clientMACBox = [System.Windows.Forms.ListBox]::new()
    $clientMACBox.Location = [System.Drawing.Point]::new(10,40)
    $clientMACBox.Height = 300
    $clientMACBox.Width = 125
    $clientMACBox.Sorted = $true

    foreach($mac in $AllDevices.macList){
        $clientMACBox.Items.Add($mac)
    }

    # Button to load Device
    $loadDeviceButton = [System.Windows.Forms.Button]::new()
    $loadDeviceButton.Location = [System.Drawing.Point]::new(10,340)
    $loadDeviceButton.Height = 20
    $loadDeviceButton.Width = 125
    $loadDeviceButton.TextAlign
    $loadDeviceButton.Text = 'Load Device'
    $loadDeviceButton.Add_Click({
        $propertyArray = @('ClientType', 'ClientVersion', 'MAC', 'LogFrequency')
        $mobilePropertyArray = @('ClientType', 'MobileClientOSVersion', 'ClientVersion', 'MAC', 'MobileClientModel', 'SSID', 'MobileClientCarrier', 'LogFrequency')
        Write-Host "Clicked"
        $selectedMAC = $clientMACBox.SelectedItem
        Write-Host "Selected MAC:  $selectedMAC"
        foreach($device in $AllDevices.deviceList){
            Write-Host "Comparing $selectedMac to " $device.MAC
            if($selectedMac -eq $device.MAC){
                Write-Host "Match Found!" $device.MAC
                
                # Clean up the labels on the devices tab
                foreach($item in $propertyArray) {
                    $deviceTab.Controls.RemoveByKey($item)
                }
                foreach($item in $voceraDeviceTraits){
                    $deviceTab.Controls.RemoveByKey($item)
                }
                foreach($item in $mobilePropertyArray){
                    $deviceTab.Controls.RemoveByKey($item)
                }
                foreach($item in $voceraMobileDeviceTraits){
                    $deviceTab.Controls.RemoveByKey($item)
                }

                    if($device.IsMobileDevice){
                        # Labels for device traits
                        $deviceDetailsLabelOffsetValue = 50
                        foreach($trait in $voceraMobileDeviceTraits){
                            $label = [System.Windows.Forms.Label]::new()
                            $label.Name = $trait
                            $label.Text = $trait+":"
                            $label.Location = [System.Drawing.Point]::new(225,$deviceDetailsLabelOffsetValue)
                            $deviceTab.Controls.Add($label)
                            $deviceDetailsLabelOffsetValue = $deviceDetailsLabelOffsetValue + 50
                        }

                        $traitLabelOffsetValue = 50
                        $device.PSObject.Properties| ForEach-Object{
                            if($mobilePropertyArray -Contains $_.Name){
                                $textBox = [System.Windows.Forms.TextBox]::new()
                                $textBox.Text = [System.Web.HttpUtility]::UrlDecode($_.Value)
                                $textBox.Name = $_.Name
                                $textBox.BorderStyle = 0
                                $textBox.BackColor = $form.BackColor
                                $textBox.Location = [System.Drawing.Point]::new(350,$traitLabelOffsetValue)
                                $deviceTab.Controls.Add($textBox)
                                $traitLabelOffsetValue = $traitLabelOffsetValue + 50
                                $form.Refresh()
                            }
                        }
                    }
                    else{
                        $traitLabelOffsetValue = 100

                        # Labels for device traits
                        $deviceDetailsLabelOffsetValue = 100
                        foreach($trait in $voceraDeviceTraits){
                            $label = [System.Windows.Forms.Label]::new()
                            $label.Name = $trait
                            $label.Text = $trait+":"
                            $label.Location = [System.Drawing.Point]::new(225,$deviceDetailsLabelOffsetValue)
                            $deviceTab.Controls.Add($label)
                            $deviceDetailsLabelOffsetValue = $deviceDetailsLabelOffsetValue + 50
                        }
                        $device.PSObject.Properties| ForEach-Object{
                            if($propertyArray -Contains $_.Name){
                                $textBox = [System.Windows.Forms.TextBox]::new()
                                $textBox.Text = $_.Value
                                $textBox.Name = $_.Name
                                $textBox.BorderStyle = 0
                                $textBox.BackColor = $form.BackColor
                                $textBox.Location = [System.Drawing.Point]::new(350,$traitLabelOffsetValue)
                                $deviceTab.Controls.Add($textBox)
                                $traitLabelOffsetValue = $traitLabelOffsetValue + 50
                                $form.Refresh()
                            }
                        }
                    }
                break
            }
        }
    })

    # Button to get all object (device) properties
    $deviceObjectPropertiesButton = [System.Windows.Forms.Button]::new()
    $deviceObjectPropertiesButton.Location = [System.Drawing.Point]::new(10,360)
    $deviceObjectPropertiesButton.Height = 20
    $deviceObjectPropertiesButton.Width = 125
    $deviceObjectPropertiesButton.TextAlign
    $deviceObjectPropertiesButton.Text = 'Object Properties'
    $deviceObjectPropertiesButton.Add_Click({
        $selectedMAC = $clientMACBox.SelectedItem
        foreach($device in $AllDevices.deviceList){
            if($device.MAC -eq $selectedMAC){
                $device | Select-Object * | Out-GridView
                break
            }
        }
    })

    # Add Elements to Device Tab
    $deviceTab.Controls.AddRange(@($clientMACBox, $loadDeviceButton, $deviceObjectPropertiesButton))

    # Users Tab
    $usersTab = [System.Windows.Forms.tabPage]::new()
    $usersTab.TabIndex = 3
    $usersTab.Text = 'Users'

    $usersBox = [System.Windows.Forms.ListBox]::new()
    $usersBox.Location = [System.Drawing.Point]::new(10,40)
    $usersBox.Height = 300
    $usersBox.Width = 175
    $usersBox.Sorted = $true
    

    foreach($userName in $AllDevices.userNameList){
        $usersBox.Items.Add($userName)
    }

     # Button to load User
     $loadUserButton = [System.Windows.Forms.Button]::new()
     $loadUserButton.Location = [System.Drawing.Point]::new(10,340)
     $loadUserButton.Height = 20
     $loadUserButton.Width = 175
     $loadUserButton.TextAlign
     $loadUserButton.Text = 'Load User'
     $loadUserButton.Add_Click({
         $propertyArray= @('UserName','VoceraID','UserID', 'siteID', 'authentications', 'associatedDevices')
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
                         # Check for and remove old information already loaded for a different device
                         if($usersTab.Controls.ContainsKey($_.Name)){
                             $usersTab.Controls.RemoveByKey($_.Name)
                         }
                         $textBox = [System.Windows.Forms.TextBox]::new()
                         $textBox.Name = $_.Name
                         $textBox.BorderStyle = 0
                         $textBox.BackColor = $form.BackColor
                         $textBox.Location = [System.Drawing.Point]::new(350,$deviceDetailsLabelOffsetValue)
                         $deviceDetailsLabelOffsetValue = $deviceDetailsLabelOffsetValue + 50
                         if($_.Name -eq 'authentications'){
                            $textBox.Text = $_.Value.Count
                         }else{
                            $textBox.Text = $_.Value
                         }
                         $usersTab.Controls.Add($textBox)
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
        $label = [System.Windows.Forms.Label]::new()
        $label.Text = $trait+":"
        $label.Location = [System.Drawing.Point]::new(250,$userDetailsLabelOffsetValue)
        $usersTab.Controls.Add($label)
        $userDetailsLabelOffsetValue = $userDetailsLabelOffsetValue + 50
    }

    # Button to display all user object properties
    $userObjectPropertiesButton = [System.Windows.Forms.Button]::new()
    $userObjectPropertiesButton.Location = [System.Drawing.Point]::new(10,360)
    $userObjectPropertiesButton.Height = 20
    $userObjectPropertiesButton.Width = 175
    $userObjectPropertiesButton.TextAlign
    $userObjectPropertiesButton.Text = 'Object Properties'
    $userObjectPropertiesButton.Add_Click({
        $selectedUser = $usersBox.SelectedItem
        foreach($userObject in $AllDevices.userList){
            if($userObject.UserName -eq $selectedUser){
                $userObject | Select-Object * | Out-GridView
                break
            }
        }
    })

    # Button to load user site ID and device associations
    $userAssociationsButton = [System.Windows.Forms.Button]::new()
    $userAssociationsButton.Location = [System.Drawing.Point]::new(10,380)
    $userAssociationsButton.Height = 20
    $userAssociationsButton.Width = 175
    $userAssociationsButton.TextAlign
    $userAssociationsButton.Text = "Find Associations"
    $userAssociationsButton.Add_Click({
        $selectedUser = $usersBox.SelectedItem
        foreach($userObject in $AllDevices.userList){
            if($userObject.UserName -eq $selectedUser){
                Get-UserIDAssociations($userObject)
            }
        }
    })

    # Add Elements to Users Tab
    $usersTab.Controls.AddRange(@($usersBox,$loadUserButton,$userObjectPropertiesButton,$userAssociationsButton))

    # VMP Server Tab
    $VMPServerTab = [System.Windows.Forms.TabPage]::new()
    $VMPServerTab.TabIndex = 5
    $VMPServerTab.Text = 'VMP Server'

    # Labels for VMP Server Traits
    $VMPServerTraitArray = @($VMPServerTraitsA, $VMPServerTraitsB)
    foreach($traitList in $VMPServerTraitArray){
        $deviceDetailsLabelOffsetValue = 50
        foreach($trait in $traitList){
            if($VMPServerTraitsA -contains $trait){
                $horrizontalOffset = 10
            }elseif($VMPServerTraitsB -contains $trait){
                $horrizontalOffset = 250
            }
            $label = [System.Windows.Forms.Label]::new()
            $label.Text = $trait+":"
            $label.Location = [System.Drawing.Point]::new($horrizontalOffset,$deviceDetailsLabelOffsetValue)
            $VMPServerTab.Controls.Add($label)
            $deviceDetailsLabelOffsetValue = $deviceDetailsLabelOffsetValue + 50
        }
    }

    # VMP Server traits
    $deviceDetailsLabelOffsetValueA = 50
    $deviceDetailsLabelOffsetValueB = 50
    $VMPServer.PSObject.Properties| ForEach-Object{
        $sideA = @('LastUserCache', 'LastUserCacheDate', 'LastUserCacheTime', 'LastUserCacheTimeDuration', 'LastHTTPServerStartDate', 'LastHTTPServerStartTime')
        $sideB = @('LastDistCache', 'LastDistCacheDate', 'LastDistCacheTime', 'LastDistCacheTimeDuration', 'LastDistCacheTimeDuration', 'LastHTTPPort', 'LastHTTPSPort')
        if($sideA -contains $_.Name){
            $horrizontalOffset = 130
            $label = [System.Windows.Forms.Label]::new()
            $label.Text = $_.Value
            $label.Location = [System.Drawing.Point]::new($horrizontalOffset,$deviceDetailsLabelOffsetValueA)
            $VMPServerTab.Controls.Add($label)
            $deviceDetailsLabelOffsetValueA = $deviceDetailsLabelOffsetValueA + 50
        }elseif($sideB -contains $_.Name){
            $horrizontalOffset = 370
            $label = [System.Windows.Forms.Label]::new()
            $label.Text = $_.Value
            $label.Location = [System.Drawing.Point]::new($horrizontalOffset,$deviceDetailsLabelOffsetValueB)
            $VMPServerTab.Controls.Add($label)
            $deviceDetailsLabelOffsetValueB = $deviceDetailsLabelOffsetValueB + 50
        }else{
            Write-Host "FAILED" -BackgroundColor Red
        }
    }

    # Add Error Tab
    $errorTab = [System.Windows.Forms.TabPage]::new()
    $errorTab.TabIndex = 4
    $errorTab.Text = "Errors"

    $errorsBox = [System.Windows.Forms.ListBox]::new()
    $errorsBox.Location = [System.Drawing.Point]::new(10,40)
    $errorsBox.Height = 300
    $errorsBox.Width = 525
    $errorsBox.Sorted = $true
    $errorsBox.HorizontalScrollbar = $true

    foreach($errorTypeInstance in $AllDevices.errorObjects){
        $errorsBox.Items.Add($errorTypeInstance.Date + " - " + $errorTypeInstance.Time + " -  " + $errorTypeInstance.Message + " - " + $errorTypeInstance.LogType)
    }

    $errorObjectPropertiesButton = [System.Windows.Forms.Button]::new()
    $errorObjectPropertiesButton.Location = [System.Drawing.Point]::new(10,360)
    $errorObjectPropertiesButton.Height = 20
    $errorObjectPropertiesButton.Width = 175
    $errorObjectPropertiesButton.TextAlign
    $errorObjectPropertiesButton.Text = 'Object Properties'
    $errorObjectPropertiesButton.Add_Click({
        $selectedError = $errorsBox.SelectedItem
        foreach($errorObject in $AllDevices.errorObjects){
            [regex] $timeRegex = "[0-2][0-9]\:[0-6][0-9]\:[0-6][0-9]\..+?(?=[ ])"
            if($errorObject.Time -eq $timeRegex.Match($selectedError)){
                $errorObject | Select-Object * | Out-GridView
                break
            }
        }
    })

    $errorTab.Controls.AddRange(@($errorsBox, $errorObjectPropertiesButton))

    # Add tabpages to tab control
    $tabControl.Controls.AddRange(@($homeTab, $deviceTab, $usersTab, $errorTab, $VMPServerTab))
    
    # Add items to main form
    $form.Controls.Add($footerLabel)
    $form.Controls.Add($tabControl)

    Draw

}


function Get-WarningInformation($errorObject){
    [regex]$warnMessageRegex = "(?<=[0-2][0-9]\:[0-6][0-9]\:[0-6][0-9]\.[0-9][0-9][0-9][ ]).+"
    $errorObject.Message = $warnMessageRegex.Matches($errorObject.FullError) | ForEach-Object {$_.Value}
}


function Get-ErrorLines{
    $errorStringList = [System.Collections.ArrayList]::new()
    foreach($file in $logFiles){
        Get-Content $file | Select-String "error:" | ForEach-Object{$errorStringList.Add($_)}
    }
    [regex]$logTypeRegex = "[A-Z]{3,}"
    [regex]$dateRegex = "[0-3][0-9]\/[0-1][0-9]\/[0-9][0-9]"
    [regex]$timeRegex = "[0-2][0-9]\:[0-6][0-9]\:[0-6][0-9]\..+?(?=[ ])"
    [regex]$errorCodeRegex = "((?<=error\:[ ])|(?<=error:[ ]\())\d+"
    [regex]$messageRegex = "(((?<=error[ ]message\:[ ])|(?<=)Execution error\:)|(?<=Curl Error\:[ ])).+"

    foreach($errorLine in $errorStringList){
        $logType = $logTypeRegex.Matches($errorLine) | ForEach-Object {$_.Value}
        $date = $dateRegex.Matches($errorLine) | ForEach-Object {$_.Value}
        $time = $timeRegex.Matches($errorLine) | ForEach-Object {$_.Value}
        $errorCode = $errorCodeRegex.Matches($errorLine) | ForEach-Object {$_.Value}
        $message = $messageRegex.Matches($errorLine) | ForEach-Object {$_.Value}

        $errorObject = [ErrorMesssage]::new()
        $errorObject.SourceType = "Log Line"
        $errorObject.LogType = $logType
        $errorObject.Date = $date
        $errorObject.Time = $time
        $errorObject.ErrorCode = $errorCode
        $errorObject.Message = $message
        $errorObject.FullError = $errorLine

        if($errorObject.LogType -eq "WARNING"){
            Get-WarningInformation($errorObject)
        }

        if($errorObject.Message -eq ""){
            $errorObject.Message = $errorObject.FullError -replace "INFO|VERBOSE|Warning", "" -replace "[0-3][0-9]\/[0-1][0-9]\/[0-9][0-9]", "" -replace "[0-2][0-9]\:[0-6][0-9]\:[0-6][0-9]\..+?(?=[ ])", ""
        }

        $AllDevices.errorObjects.Add($errorObject)
    }
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
    $mobileDetailsList = [System.Collections.ArrayList]::new()
    foreach($file in $logFiles){
        Get-Content $file | Select-String "Client details" -CaseSensitive | ForEach-Object{$deviceDetailsList.Add($_)}
        Get-Content $file | Select-String "service\: \/WIC\, querystring\: PIN" | ForEach-Object{$mobileDetailsList.Add($_)}
    }
    foreach($deviceDetails in $deviceDetailsList){
        [regex]$clientTypeRegex = "(?<=clientType[:]).+?(?=[,])"
        [regex]$clientProtoRegex = "(?<=clientProto[:]).+?(?=[,])"
        [regex]$clientVersionRegex = "(?<=clientVersion[:]).+?(?=[,])"
        [regex]$MACRegex = "(?<=mac[:]).+?(?=[,])"

        $mac = $MACRegex.Matches($deviceDetails) | ForEach-Object {$_.Value}
        
        if($AllDevices.macList -notcontains $mac){
            $device = [VoceraClientDevice]::new()
            $device.ClientType = $clientTypeRegex.Matches($deviceDetails) | ForEach-Object {$_.Value}
            $device.ClientProto = $clientProtoRegex.Matches($deviceDetails) | ForEach-Object {$_.Value}
            $device.ClientVersion = $clientVersionRegex.Matches($deviceDetails) | ForEach-Object {$_.Value}
            $device.MAC = $mac

            $AllDevices.deviceList.Add($device)
            $AllDevices.macList.Add($device.MAC)
        }
    }
    foreach($mobileDetails in $mobileDetailsList){
        [regex]$mobilePIN = "(?<=WIC\,[ ]querystring\:[ ]PIN\=).+?(\&=?)"  # VCS PIN. THIS VALUE SHOULD ALMOST ALWAYS BE OBSCURED ex. 'PIN=*****'
        [regex]$mobileProto = "(?<=proto\=).+?(?=\&)"
        [regex]$mobileVersion = "(?<=\&ver\=).+?(?=\&)"
        [regex]$mobileDeviceOS = "(?<=[0-9][0-9]\%20).+?(?=\&)"
        [regex]$mobileDeviceOSVersion = "(?<=osVersion\=).+?(?=\&)"
        [regex]$mobileDeviceModel = "(?<=deviceModel\=).+?(?=\&)"
        [regex]$mobileDeviceMAC = "(?<=MAC\=).+?(?=\&)"
        [regex]$mobileDeviceSSID = "(?<=SSID\=).+?(?=\&)"
        [regex]$mobileDeviceCarrier = "(?<=carrier\=).+?(?=\&)"
        [regex]$mobileDeviceConnectionID = "(?<=\{)[0-9].+?(?=})"

        $mac = $mobileDeviceMAC.Match($mobileDetails) | ForEach-Object {$_.Value}
        $tempConnectionID = $mobileDeviceConnectionID.Matches($mobileDetails) | ForEach-Object {$_.value}

        if($AllDevices.macList -notcontains $mac){
            $mobileDevice = [VoceraClientDevice]::new()
            $mobileDevice.MAC = $mac
            $mobileDevice.MobileClientPIN = $mobilePIN.Matches($mobileDetails) | ForEach-Object {$_.Value}
            $mobileDevice.ClientProto = $mobileProto.Matches($mobileDetails) | ForEach-Object {$_.value}
            $mobileDevice.ClientVersion = $mobileVersion.Matches($mobileDetails) | ForEach-Object {$_.value}
            $mobileDevice.MobileClientOSVersion = $mobileDeviceOSVersion.Matches($mobileDetails) | ForEach-Object {$_.Value}
            $mobileDevice.ClientType = $mobileDeviceOS.Matches($mobileDetails) | ForEach-Object {$_.Value}
            $mobileDevice.MobileClientModel = $mobileDeviceModel.Matches($mobileDetails) | ForEach-Object {$_.Value}
            $mobileDevice.SSID = $mobileDeviceSSID.Matches($mobileDetails) | ForEach-Object {$_.value}
            $mobileDevice.MobileClientCarrier = $mobileDeviceCarrier.Matches($mobileDetails) | ForEach-Object {$_.Value}
            $mobileDevice.connectionIDList.Add($tempConnectionID)
            $mobileDevice.IsMobileDevice = $true
            
            $AllDevices.deviceList.Add($mobileDevice)
            $AllDevices.macList.Add($mobileDevice.MAC)
        }else{
            foreach($device in $AllDevices.deviceList){
                if($device.MAC -eq $mac){
                    $device.LogFrequency++
                    if($device.connectionIDList -notcontains $tempConnectionID){
                        $device.connectionIDList.Add($tempConnectionID)
                    }
                }    
            }
        }
    }
}


function Get-UserEmailAndDeviceID($userObject){
    $syncLines = [System.Collections.ArrayList]::new()
    foreach($file in $logFiles){
        $matchString = $userObject.UserName
        Get-Content $file | Select-String "Sync[:].+Name:$matchString" | ForEach-Object{$syncLines.Add($_)}
    }
    foreach($line in $syncLines){
        [regex]$userEmailRegex = "(?<=Email:).+?(?=\,)"
        [regex]$deviceIDRegex = "(?<=PIN:).+"

        $userObject.UserEmail = $userEmailRegex.Matches($line) | ForEach-Object {$_.Value}
        $userObject.deviceIDs = $deviceIDRegex.Matches($line) | ForEach-Object {$_.Value}
    }
}


function Get-UserIDAssociations($userObject){
    $siteIDList = [System.Collections.ArrayList]::new()
    $macAssociationsList = [System.Collections.ArrayList]::new()  # Also the same line that contains DND and voice forwarding information
    $userID = $userObject.UserID
    foreach($file in $logFiles){
        Get-Content $file | Select-String "in database, UserID: $userID" | ForEach-Object{$siteIDList.Add($_)}
        Get-Content $file | Select-String ", userID=$userID, VoicePresence" | ForEach-Object{$macAssociationsList.Add($_)}
    }
    foreach($siteIDLine in $siteIDList){
        [regex]$siteIDRegex = "(?<=\,[ ]SiteID\:[ ]).+"

        $siteID = $siteIDRegex.Matches($siteIDLine) | ForEach-Object {$_.Value}
        $userObject.siteID.add($siteID)
    }
    foreach($macAssociationLine in $macAssociationsList){
        [regex]$macRegex = "(?<=\,[ ]MACAddress\:).{12}"
        
        $associatedMAC = $macRegex.Matches($macAssociationLine) | ForEach-Object {$_.Value}
        $userObject.associatedDevices.add($associatedMAC)
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


function Get-VMPServerInformation(){
    $userCacheList = [System.Collections.ArrayList]::new()
    $distCacheList = [System.Collections.ArrayList]::new()
    $startHTTPServerList = [System.Collections.ArrayList]::new()
    $portListHTTP = [System.Collections.ArrayList]::new()
    $portListHTTPS = [System.Collections.ArrayList]::new()
    foreach($file in $logFiles){
        Get-Content $file | Select-String "Users cache. [0-9]" | ForEach-Object{$userCacheList.Add($_)}
        Get-Content $file | Select-String "DistLists cache. [0-9]" | ForEach-Object{$distCacheList.Add($_)}
        Get-Content $file | Select-String "Starting HTTP server ..." | ForEach-Object{$startHTTPServerList.Add($_)}
        Get-Content $file | Select-String "HTTP interface activated on" | ForEach-Object{$portListHTTP.Add($_)}
        Get-Content $file | Select-String "HTTPs interface is activated" | ForEach-Object{$portListHTTPS.Add($_)}
    }
    foreach($userLine in $userCacheList){
        [regex]$userCacheRegex = "(?<=Users[ ]cache[.][ ])[0-9]+?(?=[ ]users)"
        [regex]$userCachingDurationRegex = "(?<=\(ms\)\:[ ]).+"
        [regex]$userDateStampRegex = "[0-3][0-9]\/[0-1][0-9]\/[0-9][0-9]"
        [regex]$userTimeStampRegex = "[0-2][0-9]\:[0-6][0-9]\:[0-6][0-9]\..+?(?=[ ])"

        $numberOfCachedUsers = $userCacheRegex.Matches($userLine) | ForEach-Object {$_.Value}

        if($numberOfCachedUsers -ne "0"){ 
            $VMPServer.LastUserCache = $numberOfCachedUsers
            $VMPServer.LastUserCacheTimeDuration = $userCachingDurationRegex.Matches($userLine) | ForEach-Object {$_.Value}
            $VMPServer.LastUserCacheDate = $userDateStampRegex.Matches($userLine) | ForEach-Object {$_.Value}
            $VMPServer.LastUserCacheTime = $userTimeStampRegex.Matches($userLine) | ForEach-Object {$_.Value}
        }
    }
    foreach($distLine in $distCacheList){
        [regex]$distCacheRegex = "(?<=DistLists[ ]cache\.[ ]).+?(?=[ ]dist)"
        [regex]$distCachingDurationRegex = "(?<=\(ms\)\:[ ]).+"
        [regex]$distDateStampRegex = "[0-3][0-9]\/[0-1][0-9]\/[0-9][0-9]"
        [regex]$distTimeStampRegex = "[0-2][0-9]\:[0-6][0-9]\:[0-6][0-9]\..+?(?=[ ])"

        $numberOfCachedDists = $distCacheRegex.Matches($distLine) | ForEach-Object {$_.Value}

        if($numberOfCachedDists -ne "0"){
            $VMPServer.LastDistCache = $numberOfCachedDists
            $VMPServer.LastDistCacheTimeDuration = $distCachingDurationRegex.Matches($distLine) | ForEach-Object {$_.Value}
            $VMPServer.LastDistCacheDate = $distDateStampRegex.Matches($distLine) | ForEach-Object {$_.Value}
            $VMPServer.LastDistCacheTime = $distTimeStampRegex.Matches($distLine) | ForEach-Object {$_.Value}
        }
    }
    foreach($startHTTPLine in $startHTTPServerList){
        [regex]$startHTTPDateStampRegex = "[0-3][0-9]\/[0-1][0-9]\/[0-9][0-9]"
        [regex]$startHTTPTimeStampRegex = "[0-2][0-9]\:[0-6][0-9]\:[0-6][0-9]\..+?(?=[ ])"

        $VMPServer.LastHTTPServerStartDate = $startHTTPDateStampRegex.Matches($startHTTPLine) | ForEach-Object {$_.Value}
        $VMPServer.LastHTTPServerStartTime = $startHTTPTimeStampRegex.Matches($startHTTPLine) | ForEach-Object {$_.value}
    }
    foreach($portLineHTTP in $portListHTTP){
        [regex]$portHTTPRegex = "(?<=interface[ ]activated[ ]on[ ]\<\*\>\,[ ]port[ ]\<).+?(?=\>)"

        $VMPServer.LastHTTPPort = $portHTTPRegex.Matches($portLineHTTP) | ForEach-Object {$_.Value}
    }
    foreach($portLineHTTPS in $portListHTTPS){
        [regex]$portHTTPSRegex = "(?<=interface[ ]is[ ]activated[ ]on[ ]\<\*\>\,[ ]port[ ]\<).+?(?=\>)"

        $VMPServer.LastHTTPSPort = $portHTTPSRegex.Matches($portLineHTTPS) | ForEach-Object {$_.Value}
    }
}


$AllDevices = [VoceraDeviceList]::new()
$VMPServer = [VMPServer]::new()
foreach($logPath in $logPathList){
    Get-LogFiles($logPath)
}


Write-Host "Building Devices" -ForegroundColor Cyan
Get-VoceraDevices
Write-Host "Building Users" -ForegroundColor Cyan
Get-VocerLogUsers
Write-Host "Finding Authentications" -ForegroundColor Cyan
Get-UserAuthentications
Write-Host "Gathering VMP Server Information" -ForegroundColor Cyan
Get-VMPServerInformation
Write-Host "Parsing Errors" -ForegroundColor Cyan
Get-ErrorLines
Write-Host "LAUNCH WORK DONE" -ForegroundColor Green
Open-Window