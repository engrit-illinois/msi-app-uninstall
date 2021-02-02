# Summary
This script uninstalls an MSI-based installed app on Windows systems (tested on Win7 and Win10).  

# Behavior
- It requires a search query string.
- It searches the registry for installed apps that contain the string in the app's display name, grabs the app's MSI product code and uses that to uninstall it.  
- Make sure your query string matches only one app. Currently, behavior is undefined if the search query matches more than one app.  

# Usage

1. Download this script or run from remote location
2. Run the script from an administrative commandline or powershell prompt, including the necessary and option parameters desired.

You may need to set the Execution Policy for the script if it is not digitally signed. You can do this temporarily by setting the execution policy to bypass for only the current PowerShell session:  
`Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`  

# Parameters
Parameters are documented at the top of the script.  

# Example
Uninstall the BigFix Enterprise Client and removes some relevant directories and registry entries:
`.\msi-app-uninstall.ps1 -AppTitleQuery "*bigfix" -StopServices "bes client" -RemoveDirs "c:\program files\bigfix,c:\program files (x86)\ibm\tivoli" -RemoveRegKeysValues "hklm:\software\bigfix\enterpriseclient,hklm:\software\wow6432node\bigfix\enterpriseclient;version"`

# Output
The only output the script sends to the console is the exit code. All other output is sent to the specified log file(s), which can be specified by optional parameters. Exit codes are also documented near the top of the script.  

# Notes
- By mseng3. See my other projects here: https://github.com/mmseng/code-compendium.
