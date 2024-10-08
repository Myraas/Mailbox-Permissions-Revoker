<#

Name:           Mailbox Permissions Revoker
Version:        2.2
Last Updated:   2024-09-28

Change Log:
- Introduced a global variable $GlobalVerbose to control the verbosity of the script.
- Replaced unnecessary Write-Host Statements with Write-Verbose
- Added the Import-CSVResults function that was missing from the previous version of the script.
- Corrected "Import-Module ExchangeOnline" with "Import-Module ExchangeOnlineManagement"

Run the AdminDroid "GetMailboxPermission.ps1" (version 3.0) script in the same directory as this script to cache mailbox permissions. You can find the script at the following link:
https://github.com/admindroid-community/powershell-scripts/blob/master/Office%20365%20Mailbox%20Permissions%20Report/GetMailboxPermission.ps1

#>

Param
(
    [Parameter(Mandatory = $false)]
    [string]$UPN = $NULL
)

$GlobalVerbose = $false

if ($GlobalVerbose) {
    $VerbosePreference = 'Continue'
} else {
    $VerbosePreference = 'SilentlyContinue'
    Clear-Host
}

$global:RemoveAll = $false
$scriptDirectory = $PSScriptRoot
$csvFiles = Get-ChildItem -Path $scriptDirectory -Filter *.csv

if ($IsWindows) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}

if ($csvFiles.Count -gt 1) {
    $sortedFiles = $csvFiles | Sort-Object LastWriteTime
    for ($i = 0; $i -lt $sortedFiles.Count - 1; $i++) {
        $oldFile = $sortedFiles[$i]
        $newName = "$($oldFile.FullName).old"
        if (-not (Test-Path $newName)) {
            Rename-Item -Path $oldFile.FullName -NewName $newName
        }
    }
    $latestCSV = $sortedFiles[-1].FullName
} elseif ($csvFiles.Count -eq 1) {
    $latestCSV = $csvFiles[0].FullName
} else {
    # Write-Host ""
}

$csvPath = $latestCSV

function Import-CSVResults {
    $global:FormattedResults = @()

    if (-not (Test-Path $csvPath)) {
        Write-Host "CSV file not found at $csvPath. Exiting."
        Exit
    }

    $csvData = Import-Csv $csvPath
    $filteredResults = $csvData | Where-Object { $_.UserWithAccess -match $UPN }

    $combinedResults = @{}

    foreach ($entry in $filteredResults) {
        $delegatedMailbox = $entry.UserPrincipalName
        $displayName = $entry.DisplayName
        $accessType = $entry.AccessType
        $username = $entry.UserPrincipalName

        if ($combinedResults.ContainsKey($delegatedMailbox)) {
            if ($combinedResults[$delegatedMailbox]["Access"] -notlike "*{$accessType}*") {
                $combinedResults[$delegatedMailbox]["Access"] += "{" + $accessType + "}"
            }
        } else {
            $combinedResults[$delegatedMailbox] = @{
                "DisplayName" = $displayName
                "Username" = $username
                "Access" = "{" + $accessType + "}"
            }
        }
    }

    $global:FormattedResults = $combinedResults.GetEnumerator() | ForEach-Object {
        [PSCustomObject]@{
            'Display Name' = $_.Value["DisplayName"]
            'Username'     = $_.Value["Username"]
            'Access'       = $_.Value["Access"]
        }
    } | Sort-Object Username

    if ($global:FormattedResults.Count -eq 0) {   
    } else {
        $global:FormattedResults | Format-Table -AutoSize
    }
}

function Connect_Exo {
    $Module = Get-Module ExchangeOnlineManagement -ListAvailable
    if ($Module.count -ne 0) {
        Write-Host "Connecting to Exchange Online..."
        Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue -Force
        Connect-ExchangeOnline
        
        $mailboxes = Get-Mailbox -ResultSize 10
        if ($mailboxes.count -gt 0) {
            $domains = $mailboxes | ForEach-Object {
                ($_.PrimarySmtpAddress -split '@')[1]
            }
            $mostCommonDomain = $domains | Group-Object | Sort-Object Count -Descending | Select-Object -First 1 -ExpandProperty Name
            Write-Host ""
            Write-Host "Connected to Office 365 Tenant: $mostCommonDomain" -ForegroundColor Green
        } else {
            Write-Host "No mailboxes found to determine the domain." -ForegroundColor yellow
        }
        Write-Host "Exchange Online PowerShell module is connected successfully"
    } else {
        Write-Host "Exchange Online PowerShell module is not available" -ForegroundColor yellow  
        $Confirm = Read-Host "Are you sure you want to install module? [Y] Yes [N] No" 
        if ($Confirm -match "[yY]") {
            Write-host "Installing Exchange Online PowerShell module"
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
        } else { 
            Write-Host "EXO V2 module is required to connect Exchange Online. Exiting."
            Exit
        }
    }
}

function Validate-UPNFormat {
    param (
        [string]$UPN
    )

    if ($UPN -match '^[^@]+@[^@]+\.[^@]+$') {
        return $true
    } else {
        return $false
    }
}

function Get-ValidUPN {
    do {
        $UPN = Read-Host "Please enter the Username/Email of the user to check (or type 'exit' to quit):"

        if ($UPN -eq "exit") {
            Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
            Write-Host "Disconnected active ExchangeOnline session"
            Exit
        }

        $isValidFormat = Validate-UPNFormat -UPN $UPN
        if (-not $isValidFormat) {
            Write-Host "Invalid UPN format. Please enter a valid email address (e.g., user@domain.com)."
            continue
        }

        $UserAccount = Get-Mailbox -Identity $UPN -ErrorAction SilentlyContinue
        if (-not $UserAccount) {
            Write-Host "Account $UPN does not exist. Please enter a valid account."
            continue
        }

        Write-Host "Account $UPN verified." -ForegroundColor Green
        return $UPN

    } while ($true)
}

function FullAccess {
    Write-Verbose "Retrieving FullAccess permissions for $UPN..."
    $MB_FullAccess = Get-Mailbox | Get-MailboxPermission -User $UPN -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains "FullAccess" } | Select-Object Identity
    if ($MB_FullAccess.count -ne 0) {
        Write-Verbose "Found FullAccess permissions for $UPN"
        return $MB_FullAccess.Identity | ForEach-Object {
            @{'Display Name' = $_; 'Access' = "FullAccess"}
        }
    } else {
        Write-Verbose "No FullAccess permissions found for $UPN"
        return @{'Display Name' = "-"; 'Access' = "FullAccess"}
    }
}

function SendAs {
    Write-Verbose "Retrieving SendAs permissions for $UPN..."
    $MB_SendAs = Get-RecipientPermission -Trustee $UPN -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains "SendAs" } | Select-Object Identity
    if ($MB_SendAs.count -ne 0) {
        Write-Verbose "Found SendAs permissions for $UPN"
        return $MB_SendAs.Identity | ForEach-Object {
            @{'Display Name' = $_; 'Access' = "SendAs"}
        }
    } else {
        Write-Verbose "No SendAs permissions found for $UPN"
        return @{'Display Name' = "-"; 'Access' = "SendAs"}
    }
}

function SendOnBehalfTo {
    Write-Verbose "Retrieving SendOnBehalfTo permissions for $UPN..."
    $MB_SendOnBehalfTo = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.GrantSendOnBehalfTo -ne $null } | ForEach-Object {
        $mailbox = $_
        $delegates = $mailbox.GrantSendOnBehalfTo | ForEach-Object {
            $delegate = (Get-Mailbox -Identity $_ -ErrorAction SilentlyContinue).UserPrincipalName
            if ($delegate -eq $UPN) {
                return $mailbox.Name
            }
        }
        $delegates | Where-Object { $_ -ne $null }
    }

    if (-not $MB_SendOnBehalfTo) {
        Write-Verbose "No SendOnBehalfTo permissions found for $UPN"
        return @{'Display Name' = "-"; 'Access' = "SendOnBehalf"}
    } else {
        Write-Verbose "Found SendOnBehalfTo permissions for $UPN"
        return $MB_SendOnBehalfTo | ForEach-Object {
            @{'Display Name' = $_; 'Access' = "SendOnBehalf"}
        }
    }
}

function Get-DelegatedMailboxInfo($DelegatedMailbox) {
    $MailboxInfo = Get-Recipient -Identity $DelegatedMailbox -ErrorAction SilentlyContinue
    if ($MailboxInfo -ne $null) {
        return @{'Display Name' = $MailboxInfo.DisplayName; 'Username' = $MailboxInfo.PrimarySmtpAddress}
    } else {
        return @{'Display Name' = $DelegatedMailbox; 'Username' = "N/A"}
    }
}

function Remove-Delegation {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$DelegationResult
    )

    $targetMailbox = $DelegationResult.Username
    $userToRemove = $UPN

    if ($DelegationResult.Access -like "*FullAccess*") {
        if ($global:RemoveAll -eq $true) {
            Remove-MailboxPermission -Identity "$targetMailbox" -User "$userToRemove" -AccessRights FullAccess -InheritanceType All -Confirm:$false
            Write-Host "FullAccess delegation has been removed for $userToRemove from $targetMailbox."
        } else {
            $removeFullAccess = Read-Host -Prompt 'Do you want to remove FullAccess? (yes/no)'
            if ($removeFullAccess -eq 'yes') {
                Remove-MailboxPermission -Identity "$targetMailbox" -User "$userToRemove" -AccessRights FullAccess -InheritanceType All -Confirm:$false
                Write-Host "FullAccess delegation has been removed for $userToRemove from $targetMailbox."
            } else {
                Write-Host "Skipped removing FullAccess for $userToRemove from $targetMailbox."
            }
        }
    }

    if ($DelegationResult.Access -like "*SendAs*") {
        if ($global:RemoveAll -eq $true) {
            Remove-RecipientPermission -Identity "$targetMailbox" -Trustee "$userToRemove" -AccessRights SendAs -Confirm:$false
            Write-Host "SendAs delegation has been removed for $userToRemove from $targetMailbox."
        } else {
            $removeSendAs = Read-Host -Prompt 'Do you want to remove SendAs? (yes/no)'
            if ($removeSendAs -eq 'yes') {
                Remove-RecipientPermission -Identity "$targetMailbox" -Trustee "$userToRemove" -AccessRights SendAs -Confirm:$false
                Write-Host "SendAs delegation has been removed for $userToRemove from $targetMailbox."
            } else {
                Write-Host "Skipped removing SendAs for $userToRemove from $targetMailbox."
            }
        }
    }

    if ($DelegationResult.Access -like "*SendOnBehalf*") {
        if ($global:RemoveAll -eq $true) {
            Set-Mailbox -Identity "$targetMailbox" -GrantSendOnBehalfTo @{Remove="$userToRemove"} -Confirm:$false
            Write-Host "SendOnBehalf delegation has been removed for $userToRemove from $targetMailbox."
        } else {
            $removeSendOnBehalf = Read-Host -Prompt 'Do you want to remove SendOnBehalf? (yes/no)'
            if ($removeSendOnBehalf -eq 'yes') {
                Set-Mailbox -Identity "$targetMailbox" -GrantSendOnBehalfTo @{Remove="$userToRemove"} -Confirm:$false
                Write-Host "SendOnBehalf delegation has been removed for $userToRemove from $targetMailbox."
            } else {
                Write-Host "Skipped removing SendOnBehalf for $userToRemove from $targetMailbox."
            }
        }
    }
}

function Start-Search {
    $Results = @()
    $PermissionsFound = @()

    if (($UPN -ne "")) {
        Write-Verbose "Searching for permissions across all mailboxes..."

        $UserInfo = Get-Mailbox | Where-Object { $_.UserPrincipalName -eq "$UPN" } | Select-Object Identity
        $Identity = $UserInfo.Identity

        $fullAccessResult = FullAccess
        if ($fullAccessResult['Display Name'] -ne "-") {
            $Results += $fullAccessResult
            $PermissionsFound += "FullAccess"
        }

        $sendAsResult = SendAs
        if ($sendAsResult['Display Name'] -ne "-") {
            $Results += $sendAsResult
            $PermissionsFound += "SendAs"
        }

        $sendOnBehalfResult = SendOnBehalfTo
        if ($sendOnBehalfResult['Display Name'] -ne "-") {
            $Results += $sendOnBehalfResult
            $PermissionsFound += "SendOnBehalf"
        }
    }

    $CombinedResults = @{}

    foreach ($Result in $Results) {
        $DelegatedMailbox = $Result['Display Name']
        $AccessType = $Result['Access']

        if ($CombinedResults.ContainsKey($DelegatedMailbox)) {
            $CombinedResults[$DelegatedMailbox] += "{$AccessType}"
        } else {
            $CombinedResults[$DelegatedMailbox] = "{$AccessType}"
        }
    }

    $global:FormattedResults = $CombinedResults.GetEnumerator() | ForEach-Object {
        $MailboxInfo = Get-DelegatedMailboxInfo($_.Key)
        [PSCustomObject]@{
            'Display Name' = $MailboxInfo['Display Name']
            'Username'     = $MailboxInfo['Username']
            'Access'       = $_.Value
        }
    } | Sort-Object Username

    if ($global:FormattedResults.Count -eq 0) {
    } else {
        $global:FormattedResults | Format-Table -AutoSize
    }
}

Connect_Exo

do {
    $UPN = Get-ValidUPN

    do {
        $importCSV = Read-Host "Do you want to import permissions from CSV? (yes/no)"
        if ($importCSV -ne "yes" -and $importCSV -ne "no") {
            Write-Host "Invalid input. Please enter 'yes' or 'no'."
        }
    } while ($importCSV -ne "yes" -and $importCSV -ne "no")

    if ($importCSV -eq "yes") {
        Import-CSVResults
    } else {
        Write-Host "Searching tenant for permissions. This may take a while..."
        Start-Search
    }

    if ($global:FormattedResults.Count -eq 0) {
        Write-Host "No permissions found for $UPN."
    } else {
        do {
            $response = Read-Host "Select an option: 1 for Remove All, 2 for Partial, 0 for reprompt"
            
            switch ($response) {
                "1" {
                    $global:RemoveAll = $true
                    Write-Host "You have selected to remove all permissions automatically."
                }
                "2" {
                    $global:RemoveAll = $false
                    Write-Host "You have selected to remove permissions partially."
                }
                "0" {
                    break
                }
                default {
                    Write-Host "Invalid selection. Please enter 1, 2, or 0."
                    continue
                }
            }

            if ($response -eq "1" -or $response -eq "2") {
                $targetMailbox = Read-Host "Enter the Username of the mailbox to remove access from:"

                $delegationResult = $global:FormattedResults | Where-Object { $_.Username -eq $targetMailbox }
                if ($delegationResult) {
                    Remove-Delegation -DelegationResult $delegationResult
                } else {
                    Write-Host "No delegation found for $targetMailbox."
                }
            }

        } while ($response -ne "0")
    }

    do {
        $checkAnother = Read-Host "Do you want to check another user? (yes/no)"
        if ($checkAnother -ne "yes" -and $checkAnother -ne "no") {
            Write-Host "Invalid input. Please enter 'yes' or 'no'."
        }
    } while ($checkAnother -ne "yes" -and $checkAnother -ne "no")

} while ($checkAnother -eq "yes")

Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
Write-Host "Disconnected active ExchangeOnline session"
