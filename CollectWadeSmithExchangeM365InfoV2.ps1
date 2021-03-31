<#
.SYNOPSIS
    Quick description of this script

.DESCRIPTION
    Longer description of what this script does

.PARAMETER FirstNumber
    This parameter does blablabla

.PARAMETER CheckVersion
    This parameter will just dump the script current version.

.INPUTS
    None. You cannot pipe objects to that script.

.OUTPUTS
    None for now

.EXAMPLE
.\Do-Something.ps1
This will launch the script and do someting

.EXAMPLE
.\Do-Something.ps1 -CheckVersion
This will dump the script name and current version like :
SCRIPT NAME : Do-Something.ps1
VERSION : v1.0

.NOTES
None

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
    https://github.com/SammyKrosoft
#>
[CmdletBinding(DefaultParameterSetName="NormalRun")]
Param(
    [Parameter(Mandatory = $false, ParameterSetName="NormalRun")][switch]$IncludeUserSpecificInfo,
    [Parameter(Mandatory = $false, ParameterSetName="NormalRun")][switch]$OnPremExchangeManagementShellCommands,
    [Parameter(Mandatory = $false, ParameterSetName="NormalRun")][switch]$OnLineExchangeManagementShellCommands,
    [Parameter(Mandatory = $false, ParameterSetName="NormalRun")][switch]$MSOLCommands,
    [Parameter(Mandatory = $false,ParameterSetName="Check")][switch]$CheckVersion
    
)

<# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
#Initializing a $Stopwatch variable to use to measure script execution
$stopwatch = [system.diagnostics.stopwatch]::StartNew()
#Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
# and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
$DebugPreference = "Continue"
# Set Error Action to your needs
$ErrorActionPreference = "SilentlyContinue"
#Script Version
$ScriptVersion = "1"
<# Version changes
v1 : added Write-Log and SammyKrosoft Scripting headers
v0.1 : first script version
#>
$ScriptName = $MyInvocation.MyCommand.Name
If ($CheckVersion) {Write-Host "SCRIPT NAME     : $ScriptName `nSCRIPT VERSION  : $ScriptVersion";exit}
# Log or report file definition
$UserDocumentsFolder = "$($env:Userprofile)\Documents"
$OutputReport = "$UserDocumentsFolder\$($ScriptName)_Output_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
$ScriptLog = "$UserDocumentsFolder\$($ScriptName)_Logging_$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>

<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
function Write-Log
{
	<#
	.SYNOPSIS
		This function creates or appends a line to a log file.
	.PARAMETER  Message
		The message parameter is the log message you'd like to record to the log file.
	.EXAMPLE
		PS C:\> Write-Log -Message 'Value1'
		This example shows how to call the Write-Log function with named parameters.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory=$true,position = 0)]
		[string]$Message,
		[Parameter(Mandatory=$false,position = 1)]
        [string]$LogFileName=$ScriptLog,
        [Parameter(Mandatory=$false, position = 2)][switch]$Silent
	)
	
	try
	{
		$DateTime = Get-Date -Format 'MM-dd-yy HH:mm:ss'
		$Invocation = "$($MyInvocation.MyCommand.Source | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)"
		Add-Content -Value "$DateTime - $Invocation - $Message" -Path $LogFileName
		if (!($Silent)){Write-Host $Message -ForegroundColor Green}
	}
	catch
	{
		Write-Error $_.Exception.Message
	}
}
<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
Write-Log "************************** Script Start **************************"

#Collect PowerShell command result in txt files: 

# Variables declaration
$OutputFilesCollection = @'
$($env:Userprofile)\Documents\OnPrem_OrgRel.txt
$($env:Userprofile)\Documents\OnPrem_Test-OrganizationRelationship.txt
$($env:Userprofile)\Documents\OnPrem_IntraOrgCon.txt
$($env:Userprofile)\Documents\OnPrem_AvaiAddSpa.txt
$($env:Userprofile)\Documents\OnPrem_SharingPolicy.txt
$($env:Userprofile)\Documents\OnPrem_WebSerVirDir.txt
$($env:Userprofile)\Documents\OnPrem_AutoDVirDir.txt
$($env:Userprofile)\Documents\OnPrem_FedTrust.txt
$($env:Userprofile)\Documents\OnPrem_FedOrgIden.txt
$($env:Userprofile)\Documents\OnPrem_FedInfo.txt
$($env:Userprofile)\Documents\OnPrem_TestFedTrust.txt
$($env:Userprofile)\Documents\OnPrem_TestFedCert.txt
$($env:Userprofile)\Documents\OnPrem_RemoteMailbox.txt
$($env:Userprofile)\Documents\OnPrem_Mailbox.txt
$($env:Userprofile)\Documents\OnPrem_Server.txt
$($env:Userprofile)\Documents\OnPrem_HybridConfig.txt
$($env:Userprofile)\Documents\O365_OrgRel.txt
$($env:Userprofile)\Documents\O365_Test-OrganizationRelationship.txt
$($env:Userprofile)\Documents\O365_IntraOrgCon.txt
$($env:Userprofile)\Documents\O365_AvaiAddSpa.txt
$($env:Userprofile)\Documents\O365_SharingPolicy.txt
$($env:Userprofile)\Documents\O365_FedTrust.txt
$($env:Userprofile)\Documents\O365_FedInfo.txt
$($env:Userprofile)\Documents\O365_FedOrgIden.txt
$($env:Userprofile)\Documents\O365_MailUser.txt
$($env:Userprofile)\Documents\O365_Mailbox.txt
$($env:Userprofile)\Documents\OnPrem_IntraOrgCon.txt
$($env:Userprofile)\Documents\OnPrem_IntraOrgConfig.txt
$($env:Userprofile)\Documents\OnPrem_AuthServer.txt
$($env:Userprofile)\Documents\OnPrem_ParApp.txt
$($env:Userprofile)\Documents\OnPrem_PartnerAppAcct.txt
$($env:Userprofile)\Documents\OnPrem_AuthConfig.txt
$($env:Userprofile)\Documents\OnPrem_AuthConfigCert.txt
$($env:Userprofile)\Documents\OnPrem_WebSerVirDir.txt
$($env:Userprofile)\Documents\OnPrem_AutoDVirDir.tx
$($env:Userprofile)\Documents\OnPrem_OrgRel.txt
$($env:Userprofile)\Documents\OnPrem_AvaiAddSpa.txt
$($env:Userprofile)\Documents\OnPrem_TestOAuthConnectivityEWS.txt
$($env:Userprofile)\Documents\OnPrem_TestOAuthConnectivityAutoD.txt
$($env:Userprofile)\Documents\OnPrem_RemoteMailbox.txt
$($env:Userprofile)\Documents\OnPrem_Mailbox.txt
$($env:Userprofile)\Documents\OnPrem_Server.txt
$($env:Userprofile)\Documents\OnPrem_ExchangeCertificates.txt
$($env:Userprofile)\Documents\OnPrem_HybridConfig.txt
$($env:Userprofile)\Documents\O365_IntraOrgCon.txt
$($env:Userprofile)\Documents\O365_IntraOrgConfig.txt
$($env:Userprofile)\Documents\O365_AuthServer.txt
$($env:Userprofile)\Documents\O365_PartnerApp.txt
$($env:Userprofile)\Documents\O365_TestOAuthConnectivityEWS.txt
$($env:Userprofile)\Documents\O365_TestOAuthConnectivityAutoD.txt
$($env:Userprofile)\Documents\O365_OrgRel.txt
$($env:Userprofile)\Documents\O365_MailUser.txt
$($env:Userprofile)\Documents\O365_Mailbox.txt
$($env:Userprofile)\Documents\Msol_ServicePrincipal.txt
$($env:Userprofile)\Documents\Msol_ServicePrincipalNames.txt
$($env:Userprofile)\Documents\Msol_ServicePrincipalCredential.txt
'@ -split "`n" | ForEach-Object { $_.trim() }

If ($IncludeUserSpecificInfo){
    Write-Log "Including user specific information..."
    
    $OnPremisesMailbox = "User1@Contoso.ca"
    $CloudMailbox = "UserCloud1@Contoso.ca"
    $CustomerOnMicrosoftDomain = "Contoso.mail.onmicrosoft.com"
    $CustomerDomain = "Contoso.ca"
    $OnPremisesExternalEWSURL = "https://mail.domain.com/ews/exchange.asmx"
    $OnPremisesAutodiscoverURL = "https://mail.domain.com/autodiscover/autodiscover.xml"
    
    Write-Log "OnPrem Mailbox: $OnPremisesMailbox"
    Write-Log "Cloud Mailbox: $CloudMailbox"
    Write-Log "Customer OnMicrosoft Domain : $CustomerOnMicrosoftDomain"
    Write-Log "Curstomer Domain: $CustomerDomain"
    Write-Log "On-Premises External EWS URL: $OnPremisesExternalEWSURL"
    Write-Log "On-Premises Autodiscover URL: $OnPremisesAutodiscoverURL"
}

# -------------------------------------------------------------------------------------------------
# In Exchange On-premises<Connect to Exchange management Shell>
# -------------------------------------------------------------------------------------------------
If ($OnPremExchangeManagementShellCommands){
    Write-Log "Used -OnPremExchangeManagementShellCommands switch ... dumping Exchange OnPrem info"
    Get-FederationTrust | Set-FederationTrust -RefreshMetadata 
    Get-AutoDiscoverVirtualDirectory | FL > $OutputFilesCollection[6]
    Get-AvailabilityAddressSpace | FL > $OutputFilesCollection[3]
    Get-FederatedOrganizationIdentifier | FL > $OutputFilesCollection[8]
    Get-FederationTrust | FL > $OutputFilesCollection[7]
    Get-HybridConfiguration | FL > $OutputFilesCollection[15]
    Get-IntraOrganizationConnector | FL > $OutputFilesCollection[2]
    Get-OrganizationRelationship | FL > $OutputFilesCollection[0]
    Get-ExchangeServer | FT name, serverrole, AdminDisplayVersion > $OutputFilesCollection[14]
    Get-SharingPolicy | FL > $OutputFilesCollection[4]
    Test-FederationTrustCertificate | FL > $OutputFilesCollection[11]
    Get-WebServicesVirtualDirectory | FL > $OutputFilesCollection[5]

    If ($IncludeUserSpecificInfo){
        Write-Log "Used -IncludeUserSpecificInfo switch ... dumping User specific info for Exchange OnPrem"
        # User specific informtion
        Get-FederationInformation -Domainname $CustomerOnMicrosoftDomain | FL > $OutputFilesCollection[9]
        # User specific information
        Get-Mailbox $OnPremisesMailbox | FL > $OutputFilesCollection[13]
        # User specific information
        Get-RemoteMailbox $CloudMailbox | FL > $OutputFilesCollection[12]
        # User specific information
        Test-FederationTrust -USerIdentity $OnPremisesMailbox > $OutputFilesCollection[10]
        Test-OrganizationRelationship -Identity "On-premises to O365 Organization Relationship" -UserIdentity $OnPremisesMailbox -Verbose > $OutputFilesCollection[1]
    }
}

# -------------------------------------------------------------------------------------------------
# In Exchange Online<Connect to Exchange Online service>：  
# -------------------------------------------------------------------------------------------------
If ($OnLineExchangeManagementShellCommands){
    Write-Log "Used -OnLineExchangeManagementShellCommands switch ... dumping Exchange OnLine info"
    Get-AvailabilityAddressSpace |  FL > $OutputFilesCollection[19]
    Get-FederatedOrganizationIdentifier | FL > $OutputFilesCollection[23]
    Get-FederationTrust | FL > $OutputFilesCollection[21]
    Get-IntraOrganizationConnector | FL > $OutputFilesCollection[18]
    Get-OrganizationRelationship | FL > $OutputFilesCollection[16]
    Get-SharingPolicy | FL > $OutputFilesCollection[20]
    If ($IncludeUserSpecificInfo){
        Write-Log "Used -IncludeUserSpecificInfo switch ... dumping User specific info for Exchange Online"
        # User specific information
        Get-FederationInformation -DomainName $CustomerDomain | FL > $OutputFilesCollection[22]
        # User specific information
        Get-Mailbox $CloudMailbox | FL > $OutputFilesCollection[25]
        Get-MailUser $OnPremisesMailbox | FL  > $OutputFilesCollection[24]
        # User specific information
        Test-OrganizationRelationship -UserIdentity $CloudMailbox  -Identity "Exchange Online to On Premises Organization Relationship" -Verbose > $OutputFilesCollection[17]
    }
}

# -------------------------------------------------------------------------------------------------
#We also need to check the Oauth settings: 
# In Exchange On-premises<Connect to Exchange management Shell>： 
# -------------------------------------------------------------------------------------------------
If ($OnPremExchangeManagementShellCommands){
    Write-Log "Now dumping Oauth related information."
    Write-Log "Used -OnPremExchangeManagementShellCommands switch ... dumping Exchange OnPrem info for Oauth settings"
    Get-AuthConfig | FL > $OutputFilesCollection[31]
    Get-ExchangeCertificate -Thumbprint (Get-AuthConfig).CurrentCertificateThumbprint | FL > $OutputFilesCollection[32]
    Get-AuthServer | FL > $OutputFilesCollection[28]
    Foreach ($i in (Get-ExchangeServer)) {Write-Host $i.FQDN; Get-ExchangeCertificate -Server $i.Identity} > $OutputFilesCollection[42]
    Get-IntraOrganizationConfiguration | FL > $OutputFilesCollection[27]
    Get-PartnerApplication | FL > $OutputFilesCollection[29]
    Get-PartnerApplication 00000002-0000-0ff1-ce00-000000000000 | Select-Object -ExpandProperty LinkedAccount | Get-User | FL > $OutputFilesCollection[30]
    If ($IncludeUserSpecificInfo){
        Write-Log "Used -IncludeUserSpecificInfo switch ... dumping User specific info for Exchange OnPrem for Oauth settings"
        # User specific information
        Test-OAuthConnectivity -Service AutoD  -TargetUri https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc -Mailbox $OnPremisesMailbox -Verbose | FL > $OutputFilesCollection[38]
        Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/ews/exchange.asmx -Mailbox $OnPremisesMailbox -Verbose | FL > $OutputFilesCollection[37]
    }
}

# -------------------------------------------------------------------------------------------------
# In Exchange Online<Connect to Exchange Online service>： 
# -------------------------------------------------------------------------------------------------
If ($OnLineExchangeManagementShellCommands){
    Write-Log "Now dumping Oauth related information."
    Write-Log "Used -OnLineExchangeManagementShellCommands switch ... dumping Exchange OnLine info for Oauth settings"
    if ($IncludeUserSpecificInfo){
        Write-Log "Used -IncludeUserSpecificInfo switch ... dumping User specific info for Exchange OnLine for Oauth settings"
        # User specific information
        Test-OAuthConnectivity -Service AutoD -TargetUri $OnPremisesAutodiscoverURL -Mailbox $CloudMailbox -Verbose | FL > $OutputFilesCollection[49]
        Test-OAuthConnectivity -Service EWS -TargetUri $OnPremisesExternalEWSURL -Mailbox $CloudMailbox -Verbose | FL > $OutputFilesCollection[48]
    }
    Get-AuthServer -Identity 00000001-0000-0000-c000-000000000000 | FL > $OutputFilesCollection[46]
    Get-IntraOrganizationConfiguration | FL > $OutputFilesCollection[45]
    Get-PartnerApplication | FL > $OutputFilesCollection[47]
}
# -------------------------------------------------------------------------------------------------
# Azure/MSOLPowershell: 
# -------------------------------------------------------------------------------------------------
If ($MSOLCommands){
    Write-Log "Used -MSOLCommands switch ... dumping MS OnLine Azure info for Oauth settings"
    Get-MsolServicePrincipal -ServicePrincipalName "00000002-0000-0ff1-ce00-000000000000" | FL  > $OutputFilesCollection[53]
    (Get-MsolServicePrincipal -ServicePrincipalName "00000002-0000-0ff1-ce00-000000000000").ServicePrincipalNames > $OutputFilesCollection[54]
    Get-MsolServicePrincipalCredential -ServicePrincipalName "00000002-0000-0ff1-ce00-000000000000" -ReturnKeyValues $true > $OutputFilesCollection[55]
}


<# /EXECUTIONS #>
<# -------------------------- CLEANUP VARIABLES -------------------------- #>

<# /CLEANUP VARIABLES#>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
Write-Log "************************** Script End **************************"
$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
Write-Host $msg
$msg = $null
$StopWatch = $null
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>