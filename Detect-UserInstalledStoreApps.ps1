# Created by:  Scott Fairchild - https://www.scottjfairchild.com
# Modified by JS and RCF 9/11/24

# NOTE: When the WMI class is added to Configuration Manager Hardware Inventory, 
#       Configuration Manager will create a view called v_GS_<Whatever You Put In The $wmiCustomClass Variable>
#       You can then create custom reports against that view.

# Set script variables
$wmiCustomNamespace = "ACG" # Will be created under the ROOT namespace
$wmiCustomClass = "UserInstalledStoreApps" # Will be created in the $wmiCustomNamespace. Will also be used to name the view in Configuration Manager
$DoNotLoadOfflineProfiles = $false # Prevents loading the ntuser.dat file for users that are not logged on

if ($DoNotLoadOfflineProfiles) {
    Write-Host -Value "DoNotLoadOfflineProfiles = True. Only logged in users will be checked"
}
else {
    Write-Host -Value "DoNotLoadOfflineProfiles = False. All user profiles will be checked"
}

# Check if custom WMI Namespace Exists. If not, create it.
$namespaceExists = Get-CimInstance -Namespace root -ClassName __Namespace -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $wmiCustomNamespace }
if (-not $namespaceExists) {
    Write-Host -Value "$wmiCustomNamespace WMI Namespace does not exist. Creating..."
    $ns = [wmiclass]'ROOT:__namespace'
    $sc = $ns.CreateInstance()
    $sc.Name = $wmiCustomNamespace
    $sc.Put() | Out-Null
}

# Check if custom WMI Class Exists. If not, create it.
$classExists = Get-CimClass -Namespace root\$wmiCustomNamespace -ClassName $wmiCustomClass -ErrorAction SilentlyContinue
if (-not $classExists) {
    Write-Host -Value "$wmiCustomClass WMI Class does not exist in the ROOT\$wmiCustomNamespace namespace. Creating..."
    $newClass = New-Object System.Management.ManagementClass ("ROOT\$($wmiCustomNamespace)", [String]::Empty, $null); 
    $newClass["__CLASS"] = $wmiCustomClass; 
    $newClass.Qualifiers.Add("Static", $true)
    $newClass.Properties.Add("UserName", [System.Management.CimType]::String, $false)
    $newClass.Properties["UserName"].Qualifiers.Add("Key", $true)
    $newClass.Properties.Add("ProdID", [System.Management.CimType]::String, $false)
    $newClass.Properties["ProdID"].Qualifiers.Add("Key", $true)
    $newClass.Put() | Out-Null
}

if ($DoNotLoadOfflineProfiles -eq $false) {
    # Remove current inventory records from WMI
    # This is done so Hardware Inventory can pick up applications that have been removed
    Write-Host -Value "Clearing current inventory records"
    Get-CimInstance -Namespace root\$wmiCustomNamespace -Query "Select * from $wmiCustomClass" | Remove-CimInstance
}

# Regex pattern for SIDs
$PatternSID = 'S-1-5-21-\d+-\d+\-\d+\-\d+$'
 
# Get all logged on user SIDs found in HKEY_USERS (ntuser.dat files that are loaded)
Write-Host -Value "Identifying users who are logged on"
$LoadedHives = Get-ChildItem Registry::HKEY_USERS | Where-Object { $_.PSChildname -match $PatternSID } | Select-Object @{name = "SID"; expression = { $_.PSChildName } }
if ($LoadedHives) {
    # Log all logged on users
    foreach ($userSID in $LoadedHives) {
        Write-Host -Value "-> $userSID"
    }
}
else {
    Write-Host -Value "-> None Found"
}

if ($DoNotLoadOfflineProfiles -eq $false) {

    # Get SID and location of ntuser.dat for all users
    Write-Host -Value "All user profiles on machine"
    $ProfileList = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' | Where-Object { $_.PSChildName -match $PatternSID } | 
    Select-Object  @{name = "SID"; expression = { $_.PSChildName } }, 
    @{name = "UserHive"; expression = { "$($_.ProfileImagePath)\ntuser.dat" } }
    # Log All User Profiles
    foreach ($userSID in $ProfileList) {
        Write-Host -Value "-> $userSID"
    }

    # Compare logged on users to all profiles and remove loggon on users from list
    Write-Host -Value "Profiles that have to be loaded from disk"
    # If logged on users found, compare profile list to see which ones are logged off
    if ($LoadedHives) {
        $UnloadedHives = Compare-Object $ProfileList.SID $LoadedHives.SID | Select-Object @{name = "SID"; expression = { $_.InputObject } }
    }
    else { # No logged on users found so lets load all profiles
        $UnloadedHives = $ProfileList | Select-Object -Property SID
    }

    # Log SID's that need to be loaded
    if ($UnloadedHives) {
        foreach ($userSID in $UnloadedHives) {
            Write-Host -Value "-> $userSID"
        }
    }
}

# Determine list of users we will iterate over
$profilesToQuery = $null
if ($DoNotLoadOfflineProfiles) {
    
    if ($LoadedHives) {
        $profilesToQuery = $LoadedHives
    }
    else {
        Write-Host -Value "No users are logged on. Exiting..."
        Write-Host -Value "****************************** Script Finished ******************************"
        Return "True"
        Exit
    }
}
else {
    $profilesToQuery = $ProfileList
}

# Loop through each profile
Foreach ($item in $profilesToQuery) {
    Write-Host -Value "-------------------------------------------------------------------------------------------------------------"
    $userName = ''

    # Get user name associated with profile from SID
    $objSID = New-Object System.Security.Principal.SecurityIdentifier ($item.SID)
    $userName = $objSID.Translate( [System.Security.Principal.NTAccount]).ToString()

    if ($DoNotLoadOfflineProfiles) {
        # Remove current inventory records from WMI
        # This is done so Hardware Inventory can pick up applications that have been removed
        Write-Host -Value "Clearing out current inventory for $userName" -Severity 1
        $escapedUserName = $userName.Replace('\', '\\')
        $delItem = Get-CimInstance -Namespace root\$wmiCustomNamespace -Query "Select * from $wmiCustomClass where UserName = '$escapedUserName'"
        if ($delItem) {
            $delItem | Remove-CimInstance
        }
    }

    # Load ntuser.dat if the user is not logged on
    if ($DoNotLoadOfflineProfiles -eq $false) {
        if ($item.SID -in $UnloadedHives.SID) {
            Write-Host -Value "Loading user hive for $userName from $($Item.UserHive)"
            reg load HKU\$($Item.SID) $($Item.UserHive) | Out-Null
        }
        else {
            Write-Host -Value "$UserName is logged on. No need to load hive from disk"
        }
    }

    Write-Host -Value "Getting installed User applications for $userName"

    # Define x64 apps location
    $userApps = Get-Item -path Registry::HKEY_USERS\$($Item.SID)\Software\RegisteredApplications\PackagedApps -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Property | Where-Object{($_ -notlike "Microsoft*") -and ($_ -notlike "Windows*") -and ($_ -notlike "*Teams*") -and ($_ -notlike "Dell*") -and ($_ -notlike "c5e2524a-ea46-4f67-841f-6a9465d9d515_cw5n1h2txyewy!App")} -ErrorAction SilentlyContinue
    if ($userApps) {
        Write-Host -Value "Found user installed applications"

        # Create array
        $array = @()

        # parse each app
        $userApps | ForEach-Object{
            
            # Clear current values
            $ProdID = ''

            # Select everything in the $userApps string up to the underscore
            $NewItem = $_.IndexOf("_")

            # Split by selecting first letter of the string and going up until the underscore
            $ProdID = $_.Substring(0, $NewItem)

            # Add strings to array
            $array + [string]$ProdID

            Write-Host -Value "-> Adding $ProdID"

            # Create new instance in WMI
            $newRec = New-CimInstance -Namespace root\$wmiCustomNamespace -ClassName $wmiCustomClass -Property @{UserName = "$userName"; ProdID = "$ProdID" }

            # Save to WMI
            $newRec | Set-CimInstance
        }

    }
    else {
        Write-Host -Value "No user applications found"
    }

    if ($DoNotLoadOfflineProfiles -eq $false) {
        # Unload ntuser.dat   
        # Let's do everything possible to make sure we no longer have a hook into the user profile,
        # because if we do, an Access Denied error will be displayed when trying to unload.     
        IF ($item.SID -in $UnloadedHives.SID) {
            # check if we loaded the hive
            Write-Host -Value "Unloading user hive for $userName"

            # Close Handles
            If ($userApps) {
                $userApps.Handle.Close()
            }

            # Set variable to $null
            $userApps = $null

            # Garbage collection
            [gc]::Collect()

            # Sleep for 2 seconds
            Start-Sleep -Seconds 2

            #unload registry hive
            reg unload HKU\$($Item.SID) | Out-Null
        }
    }
}

# Provide a result for MEMCM configuration baseline
Return "True"