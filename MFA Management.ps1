<#
  .Synopsis
    MFA Management, reports
    Get the MFA status for a single user or report on all Licensed users.
  .DESCRIPTION
    This script will Disable or Enable MFA and get the Azure MFA Status for your users. You can query all Licensed users, or a single user.
   
  .NOTES
    Name: MFA Management
    Author: Madina Gotova
    Version: 1.1
    DateCreated: May 2021
    Purpose/Change: Automate user Management
#>

#Check if required module installed
if (-not (Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue)) {
    Write-Host
    $msg = "INFO: MSOnline PowerShell module not installed."
    Write-Host $msg
    $msg = "INFO: Installing MSOnline PowerShell module."
    Write-Host $msg
    
    try {
        Install-Module MSOnline -AllowClobber -Force
    }
    catch {
        $msg = "ERROR: Failed to install MSOnline module. Script will abort."
        Write-Host -ForegroundColor Red $msg
        Write-Host
        $msg = "ACTION: Run this script 'As administrator' to intall the MSOnline module."
        Write-Host -ForegroundColor Yellow $msg
        Exit
    }
}

#-----------------------------------------------
function Show-Menu {
    param (
        [string]$Title = 'Options'
    )
    Write-Host "================ $Title ================"
    
    Write-Host "1: Press '1' to Enable MFA for user."
    Write-Host "2: Press '2' to Disable MFA for user."
    Write-Host "3: Press '3' to check a single user MFA status."
    Write-Host "4: Press '4' to check all Licensed users MFA Status."
    Write-Host "5: Press '5' to Exit."
}


#-----------------------------------------------
function Enable-MFA {
    $UPN = Read-host -Prompt "Enter email address of the user to enable MFA"
    $st = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
    $st.RelyingParty = "*"
    $st.State = "Enabled"
    $sta = @($st)
    try {
        # Enabling MFA for user
        Set-MsolUser -UserPrincipalName $UPN -StrongAuthenticationRequirements $sta  -ErrorAction Stop
    }
    catch {
        write-warning " Entered email address is not valid`r`n"
    }
}

#-----------------------------------------------
function Disable-MFA {
    $UPN = Read-host -Prompt "Enter email address of the user to disable MFA"
    
    try {
        # Disabling MFA for user
        Set-MsolUser -UserPrincipalName $UPN -StrongAuthenticationRequirements @() -ErrorAction Stop
    }
    catch {
        write-warning " Entered email address is not valid`r`n"
    }
}

#-----------------------------------------------
function Get-MFAStatus {
    $UPN = Read-host -Prompt "Enter email address of the user to check MFA status"
    
    try {
        $MsolUser = Get-MsolUser -UserPrincipalName $UPN -ErrorAction Stop
        if ($MSOLUser.StrongAuthenticationRequirements.State -ne $NULL) {
            $MFAEnabled = "Enabled"
        }
        else {
            $MFAEnabled = "Disabled"
        }
        write-host "`r`n MFA status of user $UPN is $MFAEnabled `r`n" -ForegroundColor green
        #    [void](Read-Host 'Press Enter to continue…')
    }
    catch {
        write-warning " Entered email address is not valid`r`n"
    }
}

#-----------------------------------------------

function Get-MFAStatusReport {
    
    try {
        Get-MsolUser -EnabledFilter EnabledOnly -MaxResults 3000 | `
            Where-Object { ($_.Licenses).AccountSkuID -match "htseng:ENTERPRISEPACK" -and $_.StrongAuthenticationRequirements.State -eq $NULL } |` 
        Select-Object userprincipalname | Export-csv c:\tools\MFAReport.csv -NoTypeInformation -ErrorAction Stop

        write-host "`nThe MFA status report has been generated and saved to the following location: c:\tools\MFAReport.csv`r`n" -foregroundcolor green
        invoke-item "c:\tools\MFAReport.csv"
    }
    catch {
        write-warning "Please close MFAReport.csv file and run report again`r`n" 
    }
}

Connect-MsolService

while ($true) {
    Show-Menu
    $inputaction = Read-Host "Please make your selection"
    clear-host
    switch ($inputaction) {
        '1' { Enable-MFA }
        '2' { Disable-MFA }
        '3' { Get-MFAStatus }
        '4' { Get-MFAStatusReport }
        '5' { exit }
    }
}