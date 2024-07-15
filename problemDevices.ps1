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
    [string]    $SiteID = "--"

    $authentications = [System.Collections.ArrayList]@()
}


class VoceraDeviceList{
    $deviceList = [System.Collections.ArrayList]::new()
    $macList = [System.Collections.ArrayList]::new()
    $userList = [System.Collections.ArrayList]::new()
    $userNameList = [System.Collections.ArrayList]::new()
}

class VoceraClientDevice {
    # General Vocera Client Traits
    [string]    $ClientType = "--" # Houses mobile client OS as well
    [string]    $MobileClientOSVersion = "-.-.-"
    [string]    $ClientProto = "--"
    [string]    $ClientVersion = "-.-.-"
    [string]    $MAC = "--"
    [string]    $MobileClientModel = "--"
    [string]    $SSID = "--"
    [string]    $MobileClientCarrier = "--"
    [int]       $LogFrequency = 1

    [bool]      $IsMobileDevice = $false

    [string]    $MobileClientPIN
}


function Open-Window{
    Write-Host "Starting GUI" -backgroundColor gray
    $voceraDeviceTraits = @('Client Type', 'Client Version', 'MAC Address', 'Log Frequency')
    $voceraMobileDeviceTraits = @('Client Type', 'OS Version', 'VCS Version', 'MAC Address', 'Model', 'SSID', 'Carrier', 'Log Frequency')
    $voceraUserTraits = @('Name','Voice ID','VMP ID', 'Site ID', 'Auth Count')
    $VMPServerTraitsA = @('Users Cached','User Cache Date','User Cache Time','Caching Time (ms)', 'HTTP Start Date', 'HTTP Start Time')
    $VMPServerTraitsB = @('DLs Cached','DL Cache Date','DL Cache Time','Caching Time (ms)', 'HTTP Port', 'HTTPS Port')

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
        $propertyArray = @('ClientType','ClientVersion','MAC','LogFrequency')
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
                            $label = New-Object System.Windows.Forms.Label
                            $label.Name = $trait
                            $label.Text = $trait+":"
                            $label.Location = New-Object System.Drawing.Point(225,$deviceDetailsLabelOffsetValue)
                            $deviceTab.Controls.Add($label)
                            $deviceDetailsLabelOffsetValue = $deviceDetailsLabelOffsetValue + 50
                        }

                        $traitLabelOffsetValue = 50
                        $device.PSObject.Properties| ForEach-Object{
                            if($mobilePropertyArray -Contains $_.Name){
                                $label = New-Object System.Windows.Forms.Label
                                $label.Text = [System.Web.HttpUtility]::UrlDecode($_.Value)
                                $label.Name = $_.Name
                                $label.Location = New-Object System.Drawing.Point(350,$traitLabelOffsetValue)
                                $deviceTab.Controls.Add($label)
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
                            $label = New-Object System.Windows.Forms.Label
                            $label.Name = $trait
                            $label.Text = $trait+":"
                            $label.Location = New-Object System.Drawing.Point(225,$deviceDetailsLabelOffsetValue)
                            $deviceTab.Controls.Add($label)
                            $deviceDetailsLabelOffsetValue = $deviceDetailsLabelOffsetValue + 50
                        }
                        $device.PSObject.Properties| ForEach-Object{
                            if($propertyArray -Contains $_.Name){
                                $label = New-Object System.Windows.Forms.Label
                                $label.Text = $_.Value
                                $label.Name = $_.Name
                                $label.Location = New-Object System.Drawing.Point(350,$traitLabelOffsetValue)
                                $deviceTab.Controls.Add($label)
                                $traitLabelOffsetValue = $traitLabelOffsetValue + 50
                                $form.Refresh()
                            }
                        }
                    }
                break
            }
        }
    })

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
         $propertyArray= @('UserName','VoceraID','UserID', 'SiteID', 'authentications')
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

    # VMP Server Tab
    $VMPServerTab = New-Object System.Windows.Forms.TabPage
    $VMPServerTab.TabIndex = 4
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
            $label = New-Object System.Windows.Forms.Label
            $label.Text = $trait+":"
            $label.Location = New-Object System.Drawing.Point($horrizontalOffset,$deviceDetailsLabelOffsetValue)
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
            $label = New-Object System.Windows.Forms.Label
            $label.Text = $_.Value
            $label.Location = New-Object System.Drawing.Point($horrizontalOffset,$deviceDetailsLabelOffsetValueA)
            $VMPServerTab.Controls.Add($label)
            $deviceDetailsLabelOffsetValueA = $deviceDetailsLabelOffsetValueA + 50
        }elseif($sideB -contains $_.Name){
            $horrizontalOffset = 370
            $label = New-Object System.Windows.Forms.Label
            $label.Text = $_.Value
            $label.Location = New-Object System.Drawing.Point($horrizontalOffset,$deviceDetailsLabelOffsetValueB)
            $VMPServerTab.Controls.Add($label)
            $deviceDetailsLabelOffsetValueB = $deviceDetailsLabelOffsetValueB + 50
        }else{
            Write-Host "FAILED" -BackgroundColor Red
        }
    }

    # Add tabpages to tab control
    $tabControl.Controls.AddRange(@($foundFilesTab, $deviceTab, $usersTab, $VMPServerTab))
    
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
        [regex]$mobileVersion = "(?<=\&ver\=).+?(?=\%)"
        [regex]$mobileDeviceOS = "(?<=[0-9][0-9]\%20).+?(?=\&)"
        [regex]$mobileDeviceOSVersion = "(?<=deviceModel\=).+?(?=\&)"
        [regex]$mobileDeviceModel = "(?<=deviceModel\=).+?(?=\&)"
        [regex]$mobileDeviceMAC = "(?<=MAC\=).+?(?=\&)"
        [regex]$mobileDeviceSSID = "(?<=SSID\=).+?(?=\&)"
        [regex]$mobileDeviceCarrier = "(?<=carrier\=).+?(?=\&)"

        $mac = $mobileDeviceMAC.Match($mobileDetails) | ForEach-Object {$_.Value}

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
            $mobileDevice.IsMobileDevice = $true

            $AllDevices.deviceList.Add($mobileDevice)
            $AllDevices.macList.Add($mobileDevice.MAC)
        }
    }
}


function Get-UserIDAssociations($UserID){
    $siteIDList = [System.Collections.ArrayList]::new()
    foreach($file in $logFiles){
        Get-Content $file | Select-String "in database, UserID: $userID" | ForEach-Object{$siteIDList.Add($_)}
    }
    foreach($siteIDLine in $siteIDList){
        [regex]$siteIDRegex = "(?<=\,[ ]SiteID\:[ ]).+"
        Write-Host $siteIDLine -BackgroundColor Gray
    }
    foreach($VMPUser in $AllDevices.userList){
        if($VMPUser.UserID -eq $userID){
            $VMPUser.SiteID = $siteIDRegex.Matches($siteIDLine) | ForEach-Object {$_.Value}
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
Write-Host "LAUNCH WORK DONE" -ForegroundColor Green
Open-Window