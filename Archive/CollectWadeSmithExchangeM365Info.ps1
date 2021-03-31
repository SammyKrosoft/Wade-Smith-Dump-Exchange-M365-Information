[CmdletBinding(DefaultParameterSetName="NormalRun")]
Param(
    [Parameter(Mandatory = $false, ParameterSetName="NormalRun")][switch]$IncludeUserSpecificInfo,
    [Parameter(Mandatory = $false,ParameterSetName="Check")][switch]$CheckVersion
    
)

#Collect PowerShell command result in txt files: 

# Variables declaration
$OutputFilesCollection = @'
d:\OnPrem_OrgRel.txt
d:\OnPrem_Test-OrganizationRelationship.txt
d:\OnPrem_IntraOrgCon.txt
d:\OnPrem_AvaiAddSpa.txt
d:\OnPrem_SharingPolicy.txt
d:\OnPrem_WebSerVirDir.txt
d:\OnPrem_AutoDVirDir.txt
d:\OnPrem_FedTrust.txt
d:\OnPrem_FedOrgIden.txt
d:\OnPrem_FedInfo.txt
d:\OnPrem_TestFedTrust.txt
d:\OnPrem_TestFedCert.txt
d:\OnPrem_RemoteMailbox.txt
d:\OnPrem_Mailbox.txt
d:\OnPrem_Server.txt
d:\OnPrem_HybridConfig.txt
d:\O365_OrgRel.txt
d:\O365_Test-OrganizationRelationship.txt
d:\O365_IntraOrgCon.txt
d:\O365_AvaiAddSpa.txt
d:\O365_SharingPolicy.txt
d:\O365_FedTrust.txt
d:\O365_FedInfo.txt
d:\O365_FedOrgIden.txt
d:\O365_MailUser.txt
d:\O365_Mailbox.txt
d:\OnPrem_IntraOrgCon.txt
d:\OnPrem_IntraOrgConfig.txt
d:\OnPrem_AuthServer.txt
d:\OnPrem_ParApp.txt
d:\OnPrem_PartnerAppAcct.txt
d:\OnPrem_AuthConfig.txt
d:\OnPrem_AuthConfigCert.txt
d:\OnPrem_WebSerVirDir.txt
d:\OnPrem_AutoDVirDir.tx
d:\OnPrem_OrgRel.txt
d:\OnPrem_AvaiAddSpa.txt
d:\OnPrem_TestOAuthConnectivityEWS.txt
d:\OnPrem_TestOAuthConnectivityAutoD.txt
d:\OnPrem_RemoteMailbox.txt
d:\OnPrem_Mailbox.txt
d:\OnPrem_Server.txt
d:\OnPrem_ExchangeCertificates.txt
d:\OnPrem_HybridConfig.txt
d:\O365_IntraOrgCon.txt
d:\O365_IntraOrgConfig.txt
d:\O365_AuthServer.txt
d:\O365_PartnerApp.txt
d:\O365_TestOAuthConnectivityEWS.txt
d:\O365_TestOAuthConnectivityAutoD.txt
d:\O365_OrgRel.txt
d:\O365_MailUser.txt
d:\O365_Mailbox.txt
d:\Msol_ServicePrincipal.txt
d:\Msol_ServicePrincipalNames.txt
d:\Msol_ServicePrincipalCredential.txt
'@ -split "`n" | ForEach-Object { $_.trim() }

$OnPremisesMailbox = "User1@Contoso.ca"
$CloudMailbox = "UserCloud1@Contoso.ca"
$CustomerOnMicrosoftDomain = "Contoso.mail.onmicrosoft.com"
$CustomerDomain = "Contoso.ca"


# In Exchange On-premises<Connect to Exchange management Shell>

Get-OrganizationRelationship | FL > $OutputFilesCollection[0]
Test-OrganizationRelationship -Identity "On-premises to O365 Organization Relationship" -UserIdentity $OnPremisesMailbox -Verbose > $OutputFilesCollection[1]
Get-IntraOrganizationConnector | FL > $OutputFilesCollection[2]
Get-AvailabilityAddressSpace | FL > $OutputFilesCollection[3]
Get-SharingPolicy | FL > $OutputFilesCollection[4]
Get-WebServicesVirtualDirectory | FL > $OutputFilesCollection[5]
Get-AutoDiscoverVirtualDirectory | FL > $OutputFilesCollection[6]
Get-FederationTrust | Set-FederationTrust -RefreshMetadata 
Get-FederationTrust | FL > $OutputFilesCollection[7]
Get-FederatedOrganizationIdentifier | FL > $OutputFilesCollection[8]
Get-FederationInformation -Domainname $CustomerOnMicrosoftDomain | FL > $OutputFilesCollection[9]
Test-FederationTrust -USerIdentity $OnPremisesMailbox > $OutputFilesCollection[10]
Test-FederationTrustCertificate | FL > $OutputFilesCollection[11]
Get-RemoteMailbox $CloudMailbox | FL > $OutputFilesCollection[12]
Get-Mailbox $OnPremisesMailbox | FL > $OutputFilesCollection[13]
Get-ExchangeServer | FT name, serverrole, AdminDisplayVersion > $OutputFilesCollection[14]
Get-HybridConfiguration | FL > $OutputFilesCollection[15]
 
# In Exchange Online<Connect to Exchange Online service>：  

Get-OrganizationRelationship | FL > $OutputFilesCollection[16]
Test-OrganizationRelationship -UserIdentity $CloudMailbox  -Identity "Exchange Online to On Premises Organization Relationship" -Verbose > $OutputFilesCollection[17]
Get-IntraOrganizationConnector | FL > $OutputFilesCollection[18]
Get-AvailabilityAddressSpace |  FL > $OutputFilesCollection[19]
Get-SharingPolicy | FL > $OutputFilesCollection[20]
Get-FederationTrust | FL > $OutputFilesCollection[21]
Get-FederationInformation -DomainName $CustomerDomain | FL > $OutputFilesCollection[22]
Get-FederatedOrganizationIdentifier | FL > $OutputFilesCollection[23]
Get-MailUser $OnPremisesMailbox | FL  > $OutputFilesCollection[24]
Get-Mailbox $CloudMailbox | FL > $OutputFilesCollection[25]
 
#We also need to check the Oauth settings: 
 
# In Exchange On-premises<Connect to Exchange management Shell>： 

New cmdlet line
Get-IntraOrganizationConnector | FL > $OutputFilesCollection[26]
Get-IntraOrganizationConfiguration | FL > $OutputFilesCollection[27]
Get-AuthServer | FL > $OutputFilesCollection[28]
Get-PartnerApplication | FL > $OutputFilesCollection[29]
Get-PartnerApplication 00000002-0000-0ff1-ce00-000000000000 | Select-Object -ExpandProperty LinkedAccount | Get-User | FL > $OutputFilesCollection[30]
Get-AuthConfig | FL > $OutputFilesCollection[31]
Get-ExchangeCertificate -Thumbprint (Get-AuthConfig).CurrentCertificateThumbprint | FL > $OutputFilesCollection[32]
Get-WebServicesVirtualDirectory | FL > $OutputFilesCollection[33]
Get-AutoDiscoverVirtualDirectory | FL > $OutputFilesCollection[34]
Get-OrganizationRelationship | FL > $OutputFilesCollection[35]
Get-AvailabilityAddressSpace | FL > $OutputFilesCollection[36]
Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/ews/exchange.asmx -Mailbox $OnPremisesMailbox -Verbose | FL > $OutputFilesCollection[37]
Test-OAuthConnectivity -Service AutoD  -TargetUri https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc -Mailbox $OnPremisesMailbox -Verbose | FL > $OutputFilesCollection[38]
Get-RemoteMailbox $CloudMailbox | FL > $OutputFilesCollection[39]
Get-Mailbox $OnPremisesMailbox | FL > $OutputFilesCollection[40]
Get-ExchangeServer | FT name, serverrole, AdminDisplayVersion > $OutputFilesCollection[41]
Foreach ($i in (Get-ExchangeServer)) {Write-Host $i.FQDN; Get-ExchangeCertificate -Server $i.Identity} > $OutputFilesCollection[42]
Get-HybridConfiguration | FL > $OutputFilesCollection[43]

# In Exchange Online<Connect to Exchange Online service>： 

Get-IntraOrganizationConnector | FL > $OutputFilesCollection[44]
Get-IntraOrganizationConfiguration | FL > $OutputFilesCollection[45]
Get-AuthServer -Identity 00000001-0000-0000-c000-000000000000 | FL > $OutputFilesCollection[46]
Get-PartnerApplication | FL > $OutputFilesCollection[47]
Test-OAuthConnectivity -Service EWS -TargetUri "<OnPremises External EWS url, for example, https://mail.domain.com/ews/exchange.asmx>" -Mailbox $CloudMailbox -Verbose | FL > $OutputFilesCollection[48]
Test-OAuthConnectivity -Service AutoD -TargetUri "<OnPremises Autodiscover.svc endpoint, for example, https://mail.domain.com/autodiscover/autodiscover.svc>" -Mailbox $CloudMailbox -Verbose | FL > $OutputFilesCollection[49]
Get-OrganizationRelationship | FL > $OutputFilesCollection[50]
Get-MailUser $OnPremisesMailbox | FL  > $OutputFilesCollection[51]
Get-Mailbox $CloudMailbox | FL > $OutputFilesCollection[52]

 
# Azure/MSOLPowershell: 

Get-MsolServicePrincipal -ServicePrincipalName "00000002-0000-0ff1-ce00-000000000000" | FL  > $OutputFilesCollection[53]
(Get-MsolServicePrincipal -ServicePrincipalName "00000002-0000-0ff1-ce00-000000000000").ServicePrincipalNames > $OutputFilesCollection[54]
Get-MsolServicePrincipalCredential -ServicePrincipalName "00000002-0000-0ff1-ce00-000000000000" -ReturnKeyValues $true > $OutputFilesCollection[55]
