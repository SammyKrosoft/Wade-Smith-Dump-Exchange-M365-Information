# Wade-Smith-Dump-Exchange-M365-Information

Series of command lines put together in a script to collect Exchange and M365 information

# Usage

> NOTE: for User specific info, make sure you populate the following variables in the script accordingly to your environment information:

```powershell
    $OnPremisesMailbox = "User1@Contoso.ca"
    $CloudMailbox = "UserCloud1@Contoso.ca"
    $CustomerOnMicrosoftDomain = "Contoso.mail.onmicrosoft.com"
    $CustomerDomain = "Contoso.ca"
    $OnPremisesExternalEWSURL = "https://mail.domain.com/ews/exchange.asmx"
    $OnPremisesAutodiscoverURL = "https://mail.domain.com/autodiscover/autodiscover.xml"
```

## To collect Exchange OnPrem only info

### Not user specific

```powershell
CollectWadeSmithExchangeM365InfoV2.ps1 -OnPremExchangeManagementShellCommands
```

### User specific
```powershell
CollectWadeSmithExchangeM365InfoV2.ps1 -OnPremExchangeManagementShellCommands -IncludeUserSpecificInfo
```

## To collect Exchange Online only info (not user specific)

### Not user specific

```powershell
CollectWadeSmithExchangeM365InfoV2.ps1 -OnLineExchangeManagementShellCommands
```

### User specific

```powershell
CollectWadeSmithExchangeM365InfoV2.ps1 -OnLineExchangeManagementShellCommands -IncludeUserSpecificInfo
```

## To collect MSOL only info (not user specific)

### Not user specific

```powershell
CollectWadeSmithExchangeM365InfoV2.ps1 -MSOLCommands
```

### User specific

```powershell
CollectWadeSmithExchangeM365InfoV2.ps1 -MSOLCommands -IncludeUserSpecificInfo
```

## To collect Exchange Online, Exchange OnPrem and MSOL info (only not user specific, global info required)

> NOTE: you need to have MSOL loaded, as well as Exchange Management Shell and Exchange Online module

```powershell
CollectWadeSmithExchangeM365InfoV2.ps1 -OnPremExchangeManagementShellCommands -OnLineExchangeManagementShellCommands -MSOLCommands
```
