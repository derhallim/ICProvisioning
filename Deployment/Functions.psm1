$bdoDepartmentPageId = "0x0101009D1CB255DA76424F860D91F20E6C411800CABC49B7300718469EAE347A8217914C"
$bdoAlbumPageId = "0x0101009D1CB255DA76424F860D91F20E6C411800CABC49B7300718469EAE347A8217925D"

function Get-BdoDepartmentPageId{
    return "0x0101009D1CB255DA76424F860D91F20E6C411800CABC49B7300718469EAE347A8217914C"
}

function Ensure-Page{
    param
    (
        [Parameter(Mandatory = $true)] [string] $pageName,
        [Parameter(Mandatory = $true)] [string] $pageTitle
    )
    try {
        get-PnPPage $pageName  -ErrorAction Stop
    }
    catch {
        $page = Add-PnPPage -Name $pageName -ContentType "BDO Site Page"
        Set-PnPListItem -Identity $page.PageId -List "Site Pages" -Values @{"Title" = $pageTitle }
    }
}




function Set-Department {
    param
    (
        [Parameter(Mandatory = $true)] [string] $SiteUrl,
        [Parameter(Mandatory = $true)] [string] $DepTaxValue,
        [Parameter(Mandatory = $true)] [string] $PostTaxValue,
        [Parameter(Mandatory = $true)] [string] $Realm, 
        [Parameter(Mandatory = $true)] [boolean] $isHub, 
        [Parameter(Mandatory = $true)] [System.Management.Automation.PSCredential] $creds ,
        [Parameter(Mandatory = $false)] [boolean] $isInD 
    )

    Write-Host "Set department settings for $SiteUrl"
    #region ###################### Configure Department ######################
    try {
        Connect-PnPOnline -Url $SiteUrl -Credentials $creds -EA Stop
        Set-PnPSite -NoScriptSite:$false
       
        Set-PnPPropertyBagValue -Key "Intranet" -Value "TRUE" -Indexed 

        $app = Get-PnPApp -Identity 55dc217b-8879-4ed3-acda-e765e94f7e20
        if (!$app.InstalledVersion) { Install-PnPApp -Identity 55dc217b-8879-4ed3-acda-e765e94f7e20 } 
        $app = Get-PnPApp -Identity 69d93904-1b0f-4f8a-9619-446fce2702ef
        if (!$app.InstalledVersion) { Install-PnPApp -Identity 69d93904-1b0f-4f8a-9619-446fce2702ef } 
    
        Set-PnPSite -NoScriptSite:$true
        if ($isInD) {
            Invoke-PnPSiteTemplate -ErrorAction Stop -Path ./InDTemplate.xml -Parameters @{
                Realm        = $Realm;
                DepTaxValue  = $DepTaxValue; 
                PostTaxValue = $PostTaxValue;
                     
            }
        }
        else {
            Invoke-PnPSiteTemplate  -ErrorAction Stop  -Path ./DepartmentTemplate.xml -Parameters @{
                Realm        = $Realm;
                DepTaxValue  = $DepTaxValue; 
                PostTaxValue = $PostTaxValue;
                 
            }
        }

        # Set-PnPSite -NoScriptSite:$true
        
        if ($isHub) {
            Set-HubSite #-SiteUrl $SiteUrl
        }
        else {
            Set-AssociatedSite
        }


    }  

    catch {
        Start-Sleep -s 60
        Write-Host $_.Exception.Message -ForegroundColor yellow
        Write-Host "Retry to configure department" -ForegroundColor yellow
        Set-Department $SiteUrl $DepTaxValue $PostTaxValue $Realm $isHub $creds $isInD
     
    }
}

function Set-Header {
    set-PnPWeb -HeaderLayout Extended -MegaMenuEnabled -HeaderEmphasis "Strong"

    Set-PnPSite -NoScriptSite:$false
    Set-PnPPropertyBagValue -Key "Intranet" -Value "TRUE" -Indexed 
    $W = GET-pnpweb
    Set-PnPPropertyBagValue -Key "BackgroundImageUrl" -Value $($w.Url + "/siteassets/intranet/BannerMar9.gif")
    Set-PnPPropertyBagValue -Key "RectSiteLogoUrl" -Value $($w.Url + "/siteassets/intranet/MyBDO_ColourWhite.png")

    # Set-PnPPropertyBagValue -Key "ThemePrimary" -Value "#E81A3B" 
    Set-PnPSite -NoScriptSite:$true
}


function Set-HubSite {
    # param
    # (
    #     [Parameter(Mandatory = $true)] [string] $SiteUrl
    # )
    # Register-PnPHubSite -Site $SiteUrl
    Set-Categories
    Set-Header
    Set-Nav
}

function Set-AssociatedSite {
    Set-Categories
    Set-Header
}

function Set-Categories {

    $ctx = Get-PnPContext
    $field = Get-PnPField -Identity "Category" -List "Events"
    $choiceField = New-Object Microsoft.SharePoint.Client.FieldChoice($ctx, $field.Path)
    $Ctx.Load($ChoiceField)
    Invoke-PnPQuery

    $ChoiceField.Choices = "Firm-wide"
    $ChoiceField.Choices += "Community"
    $ChoiceField.Choices += "Industry"
    $ChoiceField.Choices += "Meetings & Conferences"
    $ChoiceField.Choices += "Celebration"

    $ChoiceField.UpdateAndPushChanges($True)
    Invoke-PnPQuery
}


function Set-PartnerCategories {
  
    $ctx = Get-PnPContext
    $field = Get-PnPField -Identity "Category" -List "Events"
    $choiceField = New-Object Microsoft.SharePoint.Client.FieldChoice($ctx, $field.Path)
    $Ctx.Load($ChoiceField)
    Invoke-PnPQuery

    $ChoiceField.Choices = "Partner Event"

    $ChoiceField.UpdateAndPushChanges($True)
    Invoke-PnPQuery
}




function Set-DocCentreWeb {
    param
    (
        [Parameter(Mandatory = $true)] [string] $DocCentreUrl,
        [Parameter(Mandatory = $true)] [string] $WebUrl,
        [Parameter(Mandatory = $true)] [string] $WebTitle, 
        [Parameter(Mandatory = $true)] [string] $DepTaxValue, 
        [Parameter(Mandatory = $true)] [System.Management.Automation.PSCredential] $creds,
        [Parameter(Mandatory = $true)] [string] $DepartmentUrl

    )

    write-host "Connecting to "  $WebTitle
    Connect-PnPOnline -Url $WebUrl  -Credentials $creds -EA Stop

    Invoke-PnPSiteTemplate -ErrorAction Stop -Path ./DocCentreWeb.xml -Parameters @{
        DepTaxValue  = $DepTaxValue; 
        DepartmentUrl = $DepartmentUrl;
        DocCentreUrl = $DocCentreUrl
    }

    Set-Header
}

