# Microsoft 365 Tools

A collection of PowerShell scripts for Microsoft 365 administration.

## Requirements

- PowerShell 7+
- Modules: `Microsoft.Graph`, `ExchangeOnlineManagement`, `PnP.PowerShell`, `ImportExcel`

```powershell
Install-Module Microsoft.Graph
Install-Module ExchangeOnlineManagement
Install-Module PnP.PowerShell
Install-Module ImportExcel
```

## Usage

**Clone the repo**

```powershell
git clone https://github.com/emildosen/m365-toolkit.git
cd m365-toolkit
```

**Run scripts**

```powershell
.\entra\Export-ConditionalAccess.ps1
```
