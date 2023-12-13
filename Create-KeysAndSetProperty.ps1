$regexSID = 'S-1-[0-59]-\d{2}-\d{8,10}-\d{8,10}-\d{8,10}-[1-9]\d{4}$'

$properties = [System.Collections.ArrayList]@(
    @{
        'key' = 'Software\Policies\Microsoft\Office\16.0\Common\Internet'
        'name' = 'OnlineStorage'
        'type' = 'Dword'
        'value' = 3
    },
    @{
        'key' = 'Software\Microsoft\Office\16.0\Word\Options'
        'name' = 'DOC-PATH'
        'type' = 'ExpandString'
        'value' = 'P:\'
    },
    @{
        'key' = 'Software\Microsoft\Office\16.0\Excel\Options'
        'name' = 'DefaultPath'
        'type' = 'ExpandString'
        'value' = 'P:\'
    },
    @{
        'key' = 'Software\Microsoft\Office\16.0\PowerPoint\RecentFolderList'
        'name' = 'Default'
        'type' = 'ExpandString'
        'value' = 'P:\'
    }
)

function Get-PathToKey {
    [CmdletBinding()]
    param(
        [Parameter()]
        [Microsoft.Win32.RegistryKey]$UserProfileInHKU,
        [Parameter()]
        [System.string]$PathToKey
    )

    return Join-Path $userProfileInHKU.ToString().Replace('HKEY_USERS', 'HKU:') $PathToKey
}

function Create-KeyAndSetProperty {
    [CmdletBinding()]
    param(
        [Parameter()]
        [System.string]$Key,
        [Parameter()]
        [System.string]$Type,
        [Parameter()]
        [System.string]$Name,
        [Parameter()]
        [System.string]$Value
    )

    write-host $Key
   
    $targetKey = try {
        Get-Item $Key -ErrorAction Stop
    }
    catch {
        New-Item $Key -Force
    }

    Write-Host $Key

    try {
        New-ItemProperty -Path $targetKey.PSPath -Name $Name -Type $Type -Value $Value -ErrorAction Stop
    }
    catch {
        Write-Host "Property $Name already exists, updating properties" -ForegroundColor DarkGreen
        Set-ItemProperty -Path $targetKey.PSPath -Name $Name -Type $Type -Value $Value
    }
}

function Set-RegKeys {
    [CmdletBinding()]
    param(
        [Parameter()]
        [System.Object]$Profiles
    )
    ForEach ($Profile in $Profiles) {
        ForEach ($p in $properties) {
            $fullPathToKey = Get-PathToKey -UserProfileInHKU $Profile[0] -PathToKey $p.key
            Create-KeyAndSetProperty -Key $fullPathToKey -Name $p.name -Type $p.type -Value $p.value
        }
    }
}

try {
    New-PSDrive HKU Registry HKEY_USERS -ErrorAction Stop
}
catch {
    Write-Host 'HKU already connected, continuing...' -ForegroundColor Green
}

$loggedInUserProfiles = Get-ChildItem HKU: -ErrorACtion SilentlyContinue | Where-Object {$_.Name -match $regexSID}

Set-RegKeys -Profiles $loggedInUserProfiles

try {
    Remove-PSDrive HKU -ErrorAction Stop
} catch {
    Write-Host 'HKU is already disconneted, nothing to do...' -ForegroundColor Orange
}

Write-Host 'Done!' -ForegroundColor Green
