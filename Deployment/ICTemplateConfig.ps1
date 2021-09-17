

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true)]
    [SecureString]$AdminPassword
)

BEGIN {
   
    # Init XML
    try {
        $ConfigString = Get-Content "./config.xml" -EA Stop
        [xml]$Config = $ConfigString.Replace('&', '&amp;')
    }
    catch {
        Write-Host "Error reading from XML Config"  -ForegroundColor Red
        Write-Host   $_.Exception.Message -ForegroundColor Red
    }


    try {

        $creds = New-Object System.Management.Automation.PSCredential $Config.Parameters.AdminAccount, $AdminPassword
        Connect-PnPOnline -Url $Config.Parameters.TenantAdmin  -Credentials $creds -EA Stop
        $realm = Get-PnPAuthenticationRealm -EA Stop
    }
    catch {
        Write-Host "Error connecting to SharePoint Online"  -ForegroundColor Red
        Write-Host   $_.Exception.Message -ForegroundColor Red
    }
}

PROCESS {
    Connect-PnPOnline -Url $Config.Parameters.Templates -Credentials $creds -EA Stop

    # Web Part App ID
    $app = Get-PnPApp -Identity 55dc217b-8879-4ed3-acda-e765e94f7e20
    if (!$app.InstalledVersion) { Install-PnPApp -Identity 55dc217b-8879-4ed3-acda-e765e94f7e20 } 
    





    $app = Get-PnPApp -Identity 69d93904-1b0f-4f8a-9619-446fce2702ef
    if (!$app.InstalledVersion) { Install-PnPApp -Identity 69d93904-1b0f-4f8a-9619-446fce2702ef } 

    Invoke-PnPSiteTemplate -Path ./ICTemplate.xml -Parameters @{
        Realm        = $Realm;
    }
}

END {
}



<#
Owners  Normal
Members  =>>>>>>>>>>>>>>> Readonly access to the pages (except for homepage)
Visitors   Normal

#>


