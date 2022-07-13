# Check and create logging path

If (!(Test-Path -Path "C:\Temp")) {
    New-Item -Path "C:\Temp" -ItemType Directory | Out-Null
}

# Start Logging

Start-Transcript -Path "C:\Temp\New-OutlookProfile.log" -Append

# Get List of Windows Accounts with Outlook Profiles Created

$userHives = Get-ChildItem "REGISTRY::HKEY_USERS" | Select Name

# Initialize array and object for Outlook Profile objects.
$outlookUsers = $()
$profileObject = [PSCustomObject]@{}

# Check each of the User Hives in the registry to see if there are any valid Outlook profiles

Write-Host "`n////////////////// CHECK FOR EXISTING OUTLOOK PROFILES //////////////////" -ForegroundColor Yellow

ForEach ($user in $userHives) {

    $userHive = "REGISTRY::$($user.Name)"
    $profileRegPath = "REGISTRY::$($user.Name)\Software\Microsoft\Office\16.0\Outlook\Profiles\"

    if (Test-Path -Path "REGISTRY::$($user.Name)\Software\Microsoft\Office\16.0\Outlook\Profiles\") {
        
        # We found an Outlook profile, add the data to an object and array for us to use later.

        $profileUser = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($user.Name.Substring(11))").ProfileImagePath.Substring(9)
        $profileName = (Get-ChildItem -Path $profileRegPath).PSChildName

        # Build Object and add to profiles array

        $profileObject = [PSCustomObject]@{WindowsUser=$profileUser;UserHive=$userHive;RegPath=$profileRegPath;ProfileName=$profileName}
        $outlookUsers += $profileObject

        # Log found Outlook Profile information to host.

        
        Write-Host "`nOutlook profile found in hive: $($user.Name)." -ForegroundColor Green
        Write-Host "Windows User: $($profileUser)" -ForegroundColor Cyan
        Write-Host "Registry Path: $($profileRegPath)" -ForegroundColor Gray
        Write-Host "Profile Name: $($profileName)" -ForegroundColor Yellow

    } Else {

        # There are no Outlook profiles in the hive.

        Write-Host "`nNo Outlook profiles found in hive: $($user.Name)." -ForegroundColor Red

    }
}

Write-Host "`n////////////////// NEW OUTLOOK PROFILE CREATION //////////////////`n" -ForegroundColor Yellow

# Loop through the users with Outlook profiles found and create a new Default O365 profile for that user account

ForEach ($user in $outlookUsers) {

    # Check to see if the new 'O365' profile already exists (eg. the script ran previously). If the profile exists, remove it.
    Write-Host "Checking to see if a new O365 profile has already been created."
    $365ProfileExists = Test-Path -Path "$($user.RegPath)\O365-$($user.WindowsUser)"

    If ($365ProfileExists) {

        # Remove the O365-Username profile in preparation for recreating it.

        Write-Host "New Profile already exists, cleaning it up..." -ForegroundColor Red
        Remove-Item -Path "$($user.RegPath)\O365-$($user.WindowsUser)" -Force -Recurse
    } Else {
        
        # The profile hasn't been created yet.

        Write-Host "Profile doesn't exist yet. Continue with creation..." -ForegroundColor Cyan
    }
    
    # Create a new Default Outlook profile for the Windows user called 'O365-Username'

    $newProfileName = "O365-$($user.WindowsUser)"
    Write-Host "Creating New Default Outlook Profile for $($user.WindowsUser) called $($newProfileName)"

    New-Item -Path $user.RegPath -Name $newProfileName| Out-Null
    Set-ItemProperty -Path "$($user.RegPath.TrimEnd("Profiles\"))" -Name "DefaultProfile" -Value $newProfileName | Out-Null
    Write-Host "New Outlook profile created and set as default"-ForegroundColor Green

    # Check and remove O365DirectConnect Reg Key

    $dcRegKeyExists = if (Get-ItemProperty -Path "$($user.UserHive)\Software\Microsoft\Office\16.0\Outlook\AutoDiscover" -Name "ExcludeExplicitO365Endpoint" -ErrorAction SilentlyContinue) {$true} Else {$false}

    If ($dcRegKeyExists) {
        Write-Host "Office365 DirectConnect is disabled via reg key, re-enabling..." -ForegroundColor Red
        Remove-ItemProperty -Path "$($user.UserHive)\Software\Microsoft\Office\16.0\Outlook\AutoDiscover" -Name "ExcludeExplicitO365Endpoint" -Force
    } Else {
        Write-Host "Nothing preventing DirectConnect." -ForegroundColor Green
    }
}

Write-Host "`n////////////////// SCRIPT COMPLETED //////////////////`n" -ForegroundColor Yellow

Stop-Transcript