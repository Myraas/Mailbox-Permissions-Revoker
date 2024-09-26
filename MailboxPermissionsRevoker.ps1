Param
(
    [Parameter(Mandatory = $false)]
    [string]$UPN = $NULL
)

$global:RemoveAll = $false

function Connect_Exo {
    $Module = Get-Module ExchangeOnlineManagement -ListAvailable
    if ($Module.count -ne 0) {
        Write-Host "Connecting to Exchange Online..."
        Import-Module ExchangeOnline -ErrorAction SilentlyContinue -Force
        Connect-ExchangeOnline
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
            Write-Host "Exiting the script."
            Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
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

        Write-Host "Account $UPN verified. Searching the tenant for permissions. This may take a while..."
        return $UPN

    } while ($true)
}

function FullAccess {
    $MB_FullAccess = $global:Mailbox | Get-MailboxPermission -User $UPN -ErrorAction SilentlyContinue | Select-Object Identity
    if ($MB_FullAccess.count -ne 0) {
        return $MB_FullAccess.Identity | ForEach-Object {
            @{'Display Name' = $_; 'Access' = "FullAccess"}
        }
    } else {
        return @{'Display Name' = "-"; 'Access' = "FullAccess"}
    }
}

function SendAs {
    $MB_SendAs = Get-RecipientPermission -Trustee $UPN -ErrorAction SilentlyContinue | Select-Object Identity
    if ($MB_SendAs.count -ne 0) {
        return $MB_SendAs.Identity | ForEach-Object {
            @{'Display Name' = $_; 'Access' = "SendAs"}
        }
    } else {
        return @{'Display Name' = "-"; 'Access' = "SendAs"}
    }
}

function SendOnBehalfTo {
    $MB_SendOnBehalfTo = $global:Mailbox | Where-Object { $_.GrantSendOnBehalfTo -match $Identity } -ErrorAction SilentlyContinue | Select-Object Name
    if ($MB_SendOnBehalfTo.count -ne 0) {
        return $MB_SendOnBehalfTo.Name | ForEach-Object {
            @{'Display Name' = $_; 'Access' = "SendOnBehalf"}
        }
    } else {
        return @{'Display Name' = "-"; 'Access' = "SendOnBehalf"}
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

    # Check for FullAccess
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

    # Check for SendAs
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

    # Check for SendOnBehalf
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

Connect_Exo

$UPN = Get-ValidUPN

$Results = @()
$PermissionsFound = @()

if (($UPN -ne "")) {
    $UserInfo = $global:Mailbox | Where-Object { $_.UserPrincipalName -eq "$UPN" } | Select-Object Identity
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

$FormattedResults = $CombinedResults.GetEnumerator() | ForEach-Object {
    $MailboxInfo = Get-DelegatedMailboxInfo($_.Key)
    [PSCustomObject]@{
        'Display Name' = $MailboxInfo['Display Name']
        'Username'     = $MailboxInfo['Username']
        'Access'       = $_.Value
    }
}

$FormattedResults | Format-Table -AutoSize

if ($FormattedResults.Count -gt 0) {
    do {
        $response = Read-Host "Select an option: 1 for Remove All, 2 for Partial, 0 for Exit"
        
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
                Write-Host "Exiting without removing any permissions."
                break
            }
            default {
                Write-Host "Invalid selection. Please enter 1, 2, or 0."
                continue
            }
        }

        if ($response -eq "1" -or $response -eq "2") {
            $targetMailbox = Read-Host "Enter the Username of the mailbox to remove access from:"

            $delegationResult = $FormattedResults | Where-Object { $_.Username -eq $targetMailbox }
            if ($delegationResult) {
                Remove-Delegation -DelegationResult $delegationResult
            } else {
                Write-Host "No delegation found for $targetMailbox."
            }
        }

    } while ($response -ne "0")
} else {
    Write-Host "No permissions found to remove for $UPN."
}

Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
Write-Host "Disconnected active ExchangeOnline session"
