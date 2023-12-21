<# 
	.SYNOPSIS
    This script changes the main SMTP in proxy addresses from one domain name to a new one, keeping the old one as an alias.
    If the mail with the new domain name is already present in the proxy addresses, the script takes no action and puts it on warning.
    Only proxy addresses with the old domain name are changed, all other addresses are kept

    Active directory module is required

	.PARAMETER Message 
	-OUfilter : Filter by organization unit
    -OldMailDomain : The domain name to be changed (Keep in alias)
    -NewMailDomain : The new main SMTP domain
    -Whatif : Demonstrates the script without actions

    .EXAMPLE
    Import the function :
    .\Update-SMTPProxyAddresses.ps1

    Execute the script in demonstation mode :
    Update-SMTPProxyAddresses -OUfilter 'OU=Utilisateurs,OU=Lab-01,DC=lab-01,DC=fr' -OldMailDomain lab-01.fr -NewMailDomain lab-02.fr -Whatif $true

    Execute the script :
    Update-SMTPProxyAddresses -OUfilter 'OU=Utilisateurs,OU=Lab-01,DC=lab-01,DC=fr' -OldMailDomain lab-01.fr -NewMailDomain lab-02.fr
	
	.NOTES
	Version				1.0
	Auteur      		Maxime LAMBEL
	Date de création	27/12/2023

	Changelog:
		- 1.0: Version initiale
#>
function Update-SMTPProxyAddresses {
    param (
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [String] 
        $OUfilter,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [String]
        $OldMailDomain,

        [Parameter(Mandatory = $true, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [String]
        $NewMailDomain,

        [Parameter(Mandatory = $false, Position = 4)]
        [ValidateNotNullOrEmpty()]
        [boolean]
        $Whatif = $false
    )

    try {
        $ADUsers = Get-ADUser -Filter * -SearchBase $OUfilter -Properties Proxyaddresses -ErrorAction stop
    }
    catch {
        Write-host "ERROR : Import Active Direcory users failed : $($_.Exception.message)" -ForegroundColor red
    }

    $UsersToProcess = @()

    foreach ($ADUser in $ADUsers) {
        try {    
            foreach ($ProxyAddrresse in $ADUser.Proxyaddresses) {
                $PAsplit = $ProxyAddrresse.split('@')
                if ($PAsplit[1] -eq $OldMailDomain -and $PAsplit[0] -cmatch '^SMTP*') {
                    $UsersToProcess += $ADUser
                }
            }
        }
        catch {
        }
    }

    $UsersWithDuplicate = @()
    $UsersToModify = @()
    foreach ($UserToProcess in $UsersToProcess) {
        $ProxyAddrresseToSet = @()
        $AddedMail = @()
        $NoDuplicateAddresse = $true
        try {
            foreach ($ProxyAddrresse in $UserToProcess.Proxyaddresses) {
                if ($ProxyAddrresse -cmatch '^SMTP*') {
                    $SMTPsplit = $ProxyAddrresse.split(':')
                    $MAILsplit = $SMTPsplit[1].split('@')
                    $NewSMTPAddresse = "SMTP:$($MAILsplit[0])@$($NewMailDomain)"
                    $NewAliasAddresse = "smtp:$($SMTPsplit[1])"
                    
                    if (!$addedMail.Contains($NewSMTPAddresse.Split(':')[1])) {
                        $AddedMail += $NewSMTPAddresse.Split(':')[1]
                        $ProxyAddrresseToSet += $NewSMTPAddresse, $NewAliasAddresse
                    }
                    else {
                        $UsersWithDuplicate += $UserToProcess
                        Write-host "Warning : User : $($UserToProcess.SamAccountName). This new entry ($($NewSMTPAddresse)) conflicts another already existing one ($($ProxyAddrresse)). No action was taken" -ForegroundColor Yellow
                        $NoDuplicateAddresse = $false
                        break
                    }
                }
                Else {
                    if (!$addedMail.Contains($ProxyAddrresse.Split(':')[1])) {
                        $AddedMail += $ProxyAddrresse.Split(':')[1]
                        $ProxyAddrresseToSet += $ProxyAddrresse
                    }
                    else {
                        $UsersWithDuplicate += $UserToProcess
                        Write-host "Warning : User : $($UserToProcess.SamAccountName). This new entry ($($NewSMTPAddresse)) conflicts another already existing one ($($ProxyAddrresse)). No action was taken" -ForegroundColor Yellow
                        $NoDuplicateAddresse = $false
                        break
                    }
                }
            }
            if ($NoDuplicateAddresse){
                $UsersToModify += $UserToProcess
            }
            $UserToProcess.Proxyaddresses = $ProxyAddrresseToSet -join ','
        }
        catch {
            Write-host "Error : Proxy addresses Generation failed : $($_.Exception.message)" -ForegroundColor Yellow
            continue
        }
    }
    foreach ($UserToModify in $UsersToModify){
        try {
            if ($Whatif) {
                Set-ADUser -Identity $UserToModify.SamAccountName -replace @{ProxyAddresses=$UserToModify.Proxyaddresses -split ","} -WhatIf
                Write-host "Sucess : User $($UserToModify.SamAccountName) was modify : New ProxyAddresses : $($UserToModify.Proxyaddresses)" -ForegroundColor Green
            }
            else{
                Set-ADUser -Identity $UserToModify.SamAccountName -replace @{ProxyAddresses=$UserToModify.Proxyaddresses -split ","}
                Write-host "Sucess : User $($UserToModify.SamAccountName) was modify : New ProxyAddresses : $($UserToModify.Proxyaddresses)" -ForegroundColor Green
            }
        }
        catch {
            Write-host "Error : Failed while replace ProxyAddresses for $($UserToModify.SamAccountName): $($_.Exception.message)" -ForegroundColor Yellow
            continue
        }
    }
}