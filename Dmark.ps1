cls
$TenantCredentials = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/?targerServer=grxpr80mb030.lamprd80.prod.outlook.com -Credential $TenantCredentials -Authentication Basic -AllowRedirection

Import-PSSession $Session -AllowClobber

Connect-MsolService -Credential $TenantCredentials

 $domain = Get-MsolDomain
 $defaultDomain = $domain.name[0]
 $UniqValue ="dmarc=fail action=none header.from=$defaultDomain" 


$answer = ''
while ($answer -notin 'y', 'n') {
    $answer = Read-Host 'Do you want to enter header message matching to' $UniqValue '[Y/N]'
}

if ($answer -eq 'y') {
    Write-Host "You have answered 'Yes'."
    New-TransportRule -name "Blocking Spoof with DMARC"  -FromScope NotInOrganization -SetSCL 9 -HeaderMatchesMessageHeader 'Authentication-Results' -HeaderMatchesPatterns $UniqValue
} elseif ($answer -eq 'n') {
    Write-Host "You have answered 'No'."
    $InputValue = read-host "Enter specific DMARC value E.G dmarc=fail action=none header.from=controlledlogistics.co.uk"
    New-TransportRule -name "Blocking Spoof with DMARC"  -FromScope NotInOrganization -SetSCL 9 -HeaderMatchesMessageHeader 'Authentication-Results' -HeaderMatchesPatterns $InputValue

}

Get-PSSession | Remove-PSSession
Write-Host "Complete