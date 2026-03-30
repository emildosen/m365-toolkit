## PowerShell Script Naming Convention

To maintain consistency and readability across all scripts in this repository, follow the **Verb-Noun.ps1** naming pattern.  
Use approved PowerShell verbs whenever possible. Each script name should clearly describe its purpose and action.  
Names must use **PascalCase**, singular nouns, and no underscores.


| Prefix       | Example Script        | Purpose / Usage                               |
| ------------ | --------------------- | --------------------------------------------- |
| Get-         | Get-Users.ps1         | Retrieve or list existing data or state       |
| Set-         | Set-Config.ps1        | Modify existing data or configuration         |
| New-         | New-User.ps1          | Create a new object or resource               |
| Remove-      | Remove-TempFiles.ps1  | Delete or remove a resource                   |
| Add-         | Add-UserToGroup.ps1   | Add to an existing collection or container    |
| Clear-       | Clear-Logs.ps1        | Empty or reset contents without deleting      |
| Copy-        | Copy-Files.ps1        | Duplicate resources from one place to another |
| Move-        | Move-Data.ps1         | Relocate resources to a new location          |
| Rename-      | Rename-Computer.ps1   | Change the name or identifier of something    |
| Test-        | Test-Connection.ps1   | Verify a condition, connection, or state      |
| Measure-     | Measure-DiskUsage.ps1 | Quantify or calculate something               |
| Compare-     | Compare-Configs.ps1   | Evaluate differences between two sets of data |
| Invoke-      | Invoke-Backup.ps1     | Run a specific action or process              |
| Start-       | Start-Service.ps1     | Initiate or begin a process or service        |
| Stop-        | Stop-Service.ps1      | Halt or terminate a running process           |
| Restart-     | Restart-VM.ps1        | Stop and start a resource again               |
| Enable-      | Enable-Feature.ps1    | Turn on or activate functionality             |
| Disable-     | Disable-Account.ps1   | Turn off or deactivate functionality          |
| Export-      | Export-Report.ps1     | Save data to an external format or file       |
| Import-      | Import-Data.ps1       | Bring external data into the environment      |
| Out-         | Out-File.ps1          | Send output to a file or external destination |
| Write-       | Write-Log.ps1         | Output information (logs, messages, etc.)     |
| Read-        | Read-Config.ps1       | Read input or configuration data              |
| Update-      | Update-Modules.ps1    | Refresh or bring something up to date         |
| ConvertTo-   | ConvertTo-Json.ps1    | Transform data into a different format        |
| ConvertFrom- | ConvertFrom-Xml.ps1   | Parse or extract data from a format           |
| Find-        | Find-Files.ps1        | Search for a specific resource or match       |
| Search-      | Search-Logs.ps1       | Perform a broad search operation              |
| Show-        | Show-Status.ps1       | Display information interactively or visually |
| Initialize-  | Initialize-System.ps1 | Prepare or set up environment prerequisites   |
| Deploy-      | Deploy-App.ps1        | Orchestrate a deployment process              |
