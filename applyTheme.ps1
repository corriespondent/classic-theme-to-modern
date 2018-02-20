<#
.SYNOPSIS
Applies a Classic Theme for a modern site

.EXAMPLE
Apply theme
PS C:\> $Credentials = Get-Credential
PS C:\> .\applyTheme.ps1  -url "https://yourtenant.sharepoint.com/sites/yoursite" -Credentials $Credentials 

#>


[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Enter the Modern site URL, e.g. 'https://yourtenant.sharepoint.com/sites/yoursite'")]
    [String]
    $url,

    [Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    [PSCredential]
    $Credentials
)


#===================================================================================
# Get credentials if not passed as parameter
#===================================================================================
if($Credentials -eq $null)
{
	$Credentials = Get-Credential -Message "Enter Admin Credentials"
}

# first upload files with PnP
# connect to SP Online
Write-Host -ForegroundColor Green "Connecting to SP Online"
Connect-PnPOnline $url -Credentials $Credentials

Write-Host

# Debug
Write-Host -ForegroundColor Green "Turning on debug mode"
Set-PnPTraceLog -On -Level Debug
Write-Host


Write-Host -ForegroundColor Green "Deploy custom Classic theme to Modern site"

# Upload the theme assets
Write-Host -ForegroundColor Green " -- Upload theme assets"
Apply-PnPProvisioningTemplate -Path .\files-theme.xml

Write-Host -ForegroundColor Cyan " ---- Theme assets upload complete"

# Now use CSOM to apply the theme
# connect
Write-Host -ForegroundColor Green "Connecting to client context"
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($url)
$Context.Credentials = $Credentials
Write-Host

$web = Get-PnPWeb
Write-Host -ForegroundColor Green "Get-PnPWeb"

if($web)
{
    Write-Host -ForegroundColor Green "Web exists"
    try
    {

        $colorPaletteUrl = $web.ServerRelativeUrl + "/SiteAssets/test.spcolor"
        $fontSchemeUrl = Out-Null
        $backgroundImageUrl = Out-Null
        $shareGenerated = $true;
        Write-Host -ForegroundColor Yellow " -- Applying theme"
        $web.ApplyTheme($colorPaletteUrl, $fontSchemeUrl, $backgroundImageUrl, $shareGenerated)
        Write-Host -ForegroundColor Yellow " -- Load web"
        $web.Context.Load($web)
        Write-Host -ForegroundColor Yellow " -- Set high timeout"
        $web.Context.RequestTimeout = [System.Threading.Timeout]::Infinite
        Write-Host -ForegroundColor Yellow " -- Execute"
        $web.Context.ExecuteQuery()
        Write-Host "Theme Applied!"

    }
    catch 
    {
        Write-Host -ForegroundColor Red "Exception occurred!" 
        Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
        Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"           
    }
}


Write-Host
Write-Host -ForegroundColor Cyan "Site deployment complete!"