# Win-Session-Manager
 A GUI tool for managing Windows remote sessions.


## Instructions
 Download the [latest portable release](https://github.com/DiadNetworks/Win-Session-Manager/releases/latest/Win-Session-Manager.zip) and extract it. Run `Session-Manager.bat`.  
 **Note:** Some functions of the tool require `Session-Manager.bat` be run as admin.  
   
 You'll also need to either [allow scripts](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.4) to be run from PowerShell **or** [unblock](https://github.com/DiadNetworks/Win-Session-Manager?tab=readme-ov-file#unblock-only-the-necessary-script-files) `Session-Manager.ps1`.  

### Unblock only the necessary script files:
 If you prefer not to change your script execution policy, you can unblock just the files you need to by opening a PowerShell terminal in the extracted folder and typing this command: `Unblock-File *`.
