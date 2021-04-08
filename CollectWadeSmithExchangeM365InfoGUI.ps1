#region Functions
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

#endregion Functions
#region setting default variables
$OnPremisesMailbox = "User1@Contoso.ca"
$CloudMailbox = "UserCloud1@Contoso.ca"
$CustomerOnMicrosoftDomain = "Contoso.mail.onmicrosoft.com"
$CustomerDomain = "Contoso.ca"
$OnPremisesExternalEWSURL = "https://mail.domain.com/ews/exchange.asmx"
$OnPremisesAutodiscoverURL = "https://mail.domain.com/autodiscover/autodiscover.xml"
#endregion setting default variables

#Region WPF Form
# Load a WPF GUI from a XAML file build with Visual Studio
Add-Type -AssemblyName presentationframework, presentationcore
$wpf = @{}
# NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
#$inputXML = Get-Content -Path ".\WPFGUIinTenLines\MainWindow.xaml"
$inputXML = @"
<Window x:Name="WECH" x:Class="WpfApp3.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp3"
        mc:Ignorable="d"
        Title="Wade Exchange Hybrid Collection" Height="450" Width="800">
    <Grid>
        <Label x:Name="lblOnPremExchMailbox" Content="On-Premises Mailbox:" HorizontalAlignment="Left" Margin="14,172,0,0" VerticalAlignment="Top" Width="138" IsEnabled="False"/>
        <Label x:Name="lblCloudMailbox" Content="Cloud Mailbox:" HorizontalAlignment="Left" Margin="14,203,0,0" VerticalAlignment="Top" Width="138" IsEnabled="False"/>
        <Label x:Name="lblOnMicrosoftDomain" Content="OnMicrosoft domain:" HorizontalAlignment="Left" Margin="14,234,0,0" VerticalAlignment="Top" Width="138" IsEnabled="False"/>
        <Label x:Name="lblCustomDomain" Content="Custom Domain Name:" HorizontalAlignment="Left" Margin="14,265,0,0" VerticalAlignment="Top" Width="138" IsEnabled="False"/>
        <Label x:Name="lblEWSExternalURL" Content="EWS External URL:" HorizontalAlignment="Left" Margin="14,296,0,0" VerticalAlignment="Top" Width="138" IsEnabled="False"/>
        <Label x:Name="lblAutodiscoverInternalURI" Content="Autodiscover Internal URL:" HorizontalAlignment="Left" Margin="14,327,0,0" VerticalAlignment="Top" Width="162" IsEnabled="False"/>
        <TextBox x:Name="txtOnPremisesMailbox" HorizontalAlignment="Left" Height="23" Margin="213,176,0,0" TextWrapping="Wrap" Text="User1@Contoso.ca" VerticalAlignment="Top" Width="438" IsEnabled="False"/>
        <TextBox x:Name="txtCloudMailbox" HorizontalAlignment="Left" Height="23" Margin="213,207,0,0" TextWrapping="Wrap" Text="CloudUser1@Contoso.ca" VerticalAlignment="Top" Width="438" IsEnabled="False"/>
        <TextBox x:Name="txtOnMicrosoftDomain" HorizontalAlignment="Left" Height="23" Margin="213,238,0,0" TextWrapping="Wrap" Text="Contoso.mail.OnMicrosoft.com" VerticalAlignment="Top" Width="438" IsEnabled="False"/>
        <TextBox x:Name="txtCustomDomain" HorizontalAlignment="Left" Height="23" Margin="213,269,0,0" TextWrapping="Wrap" Text="Contoso.ca" VerticalAlignment="Top" Width="438" IsEnabled="False"/>
        <TextBox x:Name="txtEWSExternalURL" HorizontalAlignment="Left" Height="23" Margin="213,299,0,0" TextWrapping="Wrap" Text="https://mail.domain.com/ews/exchange.asmx" VerticalAlignment="Top" Width="438" IsEnabled="False"/>
        <TextBox x:Name="txtAutodiscoverInternalURI" HorizontalAlignment="Left" Height="23" Margin="213,330,0,0" TextWrapping="Wrap" Text="https://mail.domain.com/autodiscover/autodiscover.xml" VerticalAlignment="Top" Width="438" IsEnabled="False"/>
        <CheckBox x:Name="chkIncludeUserSpecificInfo" Content="IncludeUserSpecificInfo&#xD;&#xA;" HorizontalAlignment="Left" Margin="14,140,0,0" VerticalAlignment="Top" Width="191" Height="16"/>
        <CheckBox x:Name="chkOnPremExchangeManagementShellCommands" Content="On-Premises Exchange Management Shell commands&#xA;" HorizontalAlignment="Left" Margin="14,10,0,0" VerticalAlignment="Top" Width="325" Height="16"/>
        <CheckBox x:Name="chkOnlineExchangeManagementShellCommands" Content="Online Exchange Management commands" HorizontalAlignment="Left" Margin="14,31,0,0" VerticalAlignment="Top" Width="325" Height="16"/>
        <CheckBox x:Name="chkMSOLcommands" Content="MS Online commands" HorizontalAlignment="Left" Margin="14,52,0,0" VerticalAlignment="Top" Width="325" Height="16"/>
        <Button x:Name="btnRun" Content="Collect" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="213,375,0,0" IsEnabled="False" Height="34" />
        <Button x:Name="btnCancel" Content="Cancel" HorizontalAlignment="Left" VerticalAlignment="Top" Width="82" Margin="569,375,0,0" Height="34" />

    </Grid>
</Window>
"@

$inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
[xml]$xaml = $inputXMLClean
$reader = New-Object System.Xml.XmlNodeReader $xaml
$tempform = [Windows.Markup.XamlReader]::Load($reader)
$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

#Get the form name to be used as parameter in functions external to form...
$FormName = $NamedNodes[0].Name

write-host "$($wpf.$txtOnPremisesMailbox.text)"
#Define events functions
#region Load, Draw (render) and closing form events
#Things to load when the WPF form is loaded aka in memory
# $wpf.$FormName.Add_Loaded({
#     $wpf.$txtOnPremisesMailbox.Text = $OnPremisesMailbox
#     $wpf.$txtCloudMailbox.Text = $CloudMailbox
#     $wpf.$txtOnMicrosoftDomain.Text = $CustomerOnMicrosoftDomain
#     $wpf.$txtCustomDomain.Text = $CustomerDomain
#     $wpf.$txtEWSExternalURL.Text = $OnPremisesExternalEWSURL
#     $wpf.$txtAutodiscoverInternalURI.Text = $OnPremisesAutodiscoverURL
#     #Update-Cmd
# })

#Things to load when the WPF form is rendered aka drawn on screen
$wpf.$FormName.Add_ContentRendered({
    #Update-Cmd
})
$wpf.$FormName.add_Closing({
    $msg = "bye bye !"
    write-host $msg
})

#endregion Load, Draw and closing form events
#End of load, draw and closing form events

#HINT: to update progress bar and/or label during WPF Form treatment, add the following:
# ... to re-draw the form and then show updated controls in realtime ...
$wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})


# Load the form:
# Older way >>>>> $wpf.MyFormName.ShowDialog() | Out-Null >>>>> generates crash if run multiple times
# Newer way >>>>> avoiding crashes after a couple of launches in PowerShell...
# USing method from https://gist.github.com/altrive/6227237 to avoid crashing Powershell after we re-run the script after some inactivity time or if we run it several times consecutively...
$async = $wpf.$FormName.Dispatcher.InvokeAsync({
    $wpf.$FormName.ShowDialog() | Out-Null
})
$async.Wait() | Out-Null
#endregion Form