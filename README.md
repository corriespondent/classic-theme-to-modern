# Description

This is a demo for applying a SharePoint classic theme and background image to a Modern SharePoint site.

# Usage

Apply theme

PS C:\> $Credentials = Get-Credential
PS C:\> .\applyTheme.ps1  -url "https://yourtenant.sharepoint.com/sites/yoursite" -Credentials $Credentials 
