[CmdletBinding()]
Param(
    [Parameter(Mandatory = $false)]
    [string]$IncludeContent,

    [Parameter(Mandatory = $true)]
    [SecureString]$AdminPassword
)

BEGIN {
    try {
        $ConfigString = Get-Content .\Config.xml -EA Stop
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

    ## For extracting the template
    # try{
    #     Get-PnPSiteTemplate -Out .\Template\PnPTemplate.xml -IncludeAllPages -PersistBrandingFiles -PersistPublishingFiles
    # }
    # catch {
    #     Write-Host   "Failed to extract the template" -ForegroundColor Red
    #     Write-Host   $_.Exception.Message -ForegroundColor Red
    # }

    ##For applying the template

   try {
       Write-Host "Creating " $Config.Parameters.SiteTitle " site collection" -ForegroundColor Yellow
       New-PnPSite -Type CommunicationSite -Title $Config.Parameters.SiteTitle -Url $Config.Parameters.SiteUrl
       Start-Sleep -s 60
       Write-Host $Config.Parameters.SiteTitle " site collection created" -ForegroundColor Green

       Write-Host "Connecting to the new site collection" -ForegroundColor White
       $creds = New-Object System.Management.Automation.PSCredential $Config.Parameters.AdminAccount, $AdminPassword
       Connect-PnPOnline -Url $Config.Parameters.SiteUrl  -Credentials $creds -EA Stop

       Write-Host "Applying the PnP Template to the site collection" -ForegroundColor Yellow
       Invoke-PnPSiteTemplate -Path .\Template\PnPTemplate.xml -EA Stop
       Write-Host "PnP Template applied to the site collection" -ForegroundColor Green
   }
   
   catch {
       Write-Host   "Failed to Create Site and applying template" -ForegroundColor Red
       Write-Host   $_.Exception.Message -ForegroundColor Red
   }

#    $IncludeListContent = Write-Host  "Do you want to include the content? If yes, then please type 'Yes' or else type 'No'"
   if($IncludeContent.ToLower() -eq 'includecontent'){

   ##Add Ittems to the It Operations list

   $ITOperationsItems = @(
       [PSCustomObject]@{
           Title = "User Account and Email Management";
           Content = '<div class="ExternalClass28F60FBF2C964074A709F290B5FCD056"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><div style="color&#58;rgb(50, 49, 48);line-height&#58;1.4em;margin&#58;0px 0px 1.4em;font-size&#58;18px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><h3 style="margin&#58;0px 0px 12px;font-weight&#58;600 !important;font-size&#58;24px;">User Account and Email Management</h3></div><ul style="color&#58;rgb(50, 49, 48);counter-reset&#58;item 0;line-height&#58;1.4em;margin-top&#58;0px;margin-bottom&#58;0px;margin-left&#58;0px !important;font-size&#58;18px;padding&#58;0px 0px 0px 50px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">One account per person, shared accounts are not used</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Email aliases are discouraged - Shared Inbox preferred for alternate email addresses</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Use Shared Inboxes when multiple people require access</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Daily-use accounts should not have Global Admin. ADM-Username accounts are created.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Role-Based Access Controls (RBAC) are in place&#160;</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">If archival of email is required, convert to shared inbox, export to PST and save in OneDrive or SharePoint</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">When deleting accounts, it is preferred to transfer OneDrive content to another user for review/transfer/archive</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">All permissions changes or Shared Inbox access require approval by Client Lead</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Standard security groups are defined to manage such things as license assignments and communications. See the current client manual for the details on those groups.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">User accounts should contain at least full name, title and phone number. Fully populate where possible for a rich internal address book, especially with multi-site larger companies.</span></li></ul></span><br></div></div>';
       }
       [PSCustomObject]@{
        Title = "Security Management";
        Content = '<div class="ExternalClassBA8AAFBDE71E4E48914FE32DB387F6F4"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><div style="color&#58;rgb(50, 49, 48);line-height&#58;1.4em;margin&#58;0px 0px 1.4em;font-size&#58;18px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><h3 style="margin&#58;0px 0px 12px;font-weight&#58;600 !important;font-size&#58;24px;">Security Management</h3></div><ul style="color&#58;rgb(50, 49, 48);counter-reset&#58;item 0;line-height&#58;1.4em;margin-top&#58;0px;margin-bottom&#58;0px;margin-left&#58;0px !important;font-size&#58;18px;padding&#58;0px 0px 0px 50px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Multifactor Authentication applied to all accounts</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">BitLocker on all hard drives managed by Intune</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Random attack simulations are to be conducted to train users on how to identify attacks.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Alerts configured to go to IC 360 for medium and high-risk events that make require intervention</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Standardized attack surface reduction policies</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Baseline Windows 10 security policies applied via Intune</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Tailored Defender and Office Advanced Threat Protection (ATP) Policies defined and deployed via Intune</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Leverage Azure Sentinel for monitoring of M365 activities&#160;</span></li></ul></span><br></div></div>'
       }
       [PSCustomObject]@{
           Title = "Information Management";
           Content = '<div class="ExternalClassA75C66A8CA0C4268AD2BC82C6A5C8D5F"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><div style="color&#58;rgb(50, 49, 48);line-height&#58;1.4em;margin&#58;0px 0px 1.4em;font-size&#58;18px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><h3 style="margin&#58;0px 0px 12px;font-weight&#58;600 !important;font-size&#58;24px;">Information Management</h3></div><ul style="color&#58;rgb(50, 49, 48);counter-reset&#58;item 0;line-height&#58;1.4em;margin-top&#58;0px;margin-bottom&#58;0px;margin-left&#58;0px !important;font-size&#58;18px;padding&#58;0px 0px 0px 50px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><li style="margin-bottom&#58;0px;margin-left&#58;8px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Data is stored in Teams/SharePoint or Azure Files</span></li><li style="margin-bottom&#58;0px;margin-left&#58;8px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Client is to identify an &quot;Information Management&quot; lead to assist with IM decisions</span></li><li style="margin-bottom&#58;0px;margin-left&#58;8px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Teams is designed to&#160;align with your organization</span></li><li style="margin-bottom&#58;0px;margin-left&#58;8px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Create Teams templates and policies to control sprawl and improve user-experience</span></li><li style="margin-bottom&#58;0px;margin-left&#58;8px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Fine-tuned&#160;settings for sharing of data internally and externally, depending on Team-level requirements</span></li><li style="margin-bottom&#58;0px;margin-left&#58;8px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Backups are discontinued when a user is deleted. They are not for long-term archive of terminated users. Use export and&#160; SharePoint/OneDrive if files must be archived.</span></li></ul></span><br></div></div>'
       }
       [PSCustomObject]@{
           Title = "Disaster Recovery";
           Content = '<div class="ExternalClassFACFB70C128943F899A63E74FE8556CE"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><div style="color&#58;rgb(50, 49, 48);line-height&#58;1.4em;margin&#58;0px 0px 1.4em;font-size&#58;18px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><h3 style="margin&#58;0px 0px 12px;font-weight&#58;600 !important;font-size&#58;24px;">Disaster Recovery</h3></div><ul style="color&#58;rgb(50, 49, 48);counter-reset&#58;item 0;line-height&#58;1.4em;margin-top&#58;0px;margin-bottom&#58;0px;margin-left&#58;0px !important;font-size&#58;18px;padding&#58;0px 0px 0px 50px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Systems are designed with the assumption that a ransomware attack will happen eventually. This is mitigated with ATP, however, if it occurs, all data must be protected.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Datto SaaS backup is set to automatically backup all M365 content. Monthly reviews are conducted to confirm this and clean-up old accounts.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Any traditional servers or workstations with production data require either Datto backup or Azure backup that syncs to an offsite location.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">If data backups are not being kept in Canada, the client must sign off on this. Generally, they are kept in Canada.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">All backups must be air-gapped from Microsoft systems/credentials.&#160;</span></li></ul></span><br></div></div>'
       }
       [PSCustomObject]@{
           Title = "Network Management";
           Content = '<div class="ExternalClass0E9965AE2EF1459893C6798803844521"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><div style="color&#58;rgb(50, 49, 48);line-height&#58;1.4em;margin&#58;0px 0px 1.4em;font-size&#58;18px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><h3 style="margin&#58;0px 0px 12px;font-weight&#58;600 !important;font-size&#58;24px;">Network Management</h3></div><ul style="color&#58;rgb(50, 49, 48);counter-reset&#58;item 0;line-height&#58;1.4em;margin-top&#58;0px;margin-bottom&#58;0px;margin-left&#58;0px !important;font-size&#58;18px;padding&#58;0px 0px 0px 50px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Managed Devices are using Datto Networking</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Switches must be managed and gigabit &amp; PoE</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Where possible, internal&#160;static IP addresses should be avoided. If a DHCP reservation can be used, that is preferred, to avoid having to ever &quot;touch&quot; the end device.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">For static IPs, leverage the bottom portion of the subnet (e.g. 192.168.1.1-30. Document this in the IP scheme notes.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Clients dependent on cloud should have backup internet, ideally over LTE if available</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Alerts of issues that affect functionality sent to IC 360s alerts inbox</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Datto and Meraki are to be set to auto-update at a time the client agrees to, generally overnight on Friday night is preferred</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Content filtering (e.g. blocking types of sites) is optional and possible with Datto and Meraki networking. The client can decide on this.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Guests should use a different network ID, and we encourage scheduling access to business hours, and changing the password/access methods periodically. Guests should only have internet access; no local network access.</span></li></ul></span><br></div></div>'
       }
       [PSCustomObject]@{
            Title = "Device Management";
            Content = '<div class="ExternalClass3D182EB9B151407CA4038C2D881A27D3"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><div style="color&#58;rgb(50, 49, 48);line-height&#58;1.4em;margin&#58;0px 0px 1.4em;font-size&#58;18px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><h3 style="margin&#58;0px 0px 12px;font-weight&#58;600 !important;font-size&#58;24px;">Device Management</h3></div><ul style="color&#58;rgb(50, 49, 48);counter-reset&#58;item 0;line-height&#58;1.4em;margin-top&#58;0px;margin-bottom&#58;0px;margin-left&#58;0px !important;font-size&#58;18px;padding&#58;0px 0px 0px 50px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Autopilot configured with main user having local admin and computername includes serial number in naming convention (e.g. client-serial)</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Tailored compliance policies defined and monitored</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Transferred devices to be reset with &quot;Fresh Start&quot; before being redeployed.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Corporate Windows and Mac computers are to run Datto RMM, deployed via Intune. Personal devices should not run the RMM software.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Devices are monitored for unwanted software using RMM and EndPoint Manager, and applications may be removed manually or automatically if detected</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Auto-purge Intune devices after a between 60 - 180 days, depending on client</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Patching of systems is through Intune. A &quot;Pilot&quot; group is used for Feature Updates 2 weeks before all devices are allowed to update.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">iOS and Android Smartphones are supported when using Intune policies. For tight&#160;corporate control including app deployment capacity, we defined Mobile Device Management policies using Company Portal registration. For a &quot;lighter&quot; management approach, we define Mobile Application Management (MAM) policies. All phones (corporate or personal) require one of these 2 policies.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Location of devices is never to be tracked by IC 360.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Corporate Mac computer management is limited, when compared to Windows management.</span></li></ul></span><br></div></div>'
        }
        [PSCustomObject]@{
            Title = "Phones";
            Content = '<div class="ExternalClass137E37ED16A044EE81F9305FA3DF5F81"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><div style="color&#58;rgb(50, 49, 48);line-height&#58;1.4em;margin&#58;0px 0px 1.4em;font-size&#58;18px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><h3 style="margin&#58;0px 0px 12px;font-weight&#58;600 !important;font-size&#58;24px;">Phones</h3></div><ul style="color&#58;rgb(50, 49, 48);counter-reset&#58;item 0;line-height&#58;1.4em;margin-top&#58;0px;margin-bottom&#58;0px;margin-left&#58;0px !important;font-size&#58;18px;padding&#58;0px 0px 0px 50px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">The client should identify a phones lead</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Teams Voice is our preferred phone solution.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Document auto-attendant config should be documented.&#160;&#160;</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Simple auto-attendant with a minimum number of options.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Review hold music with the client. No annoying or repetitive hold music.</span></li></ul></span><br></div></div>'
        }
        [PSCustomObject]@{
            Title = "Printer Management";
            Content = '<div class="ExternalClass6CC8C0A7AFBF40B2A0C96048C0878BAB"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><div style="color&#58;rgb(50, 49, 48);line-height&#58;1.4em;margin&#58;0px 0px 1.4em;font-size&#58;18px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><h3 style="margin&#58;0px 0px 12px;font-weight&#58;600 !important;font-size&#58;24px;">Printer Management</h3></div><ul style="color&#58;rgb(50, 49, 48);counter-reset&#58;item 0;line-height&#58;1.4em;margin-top&#58;0px;margin-bottom&#58;0px;margin-left&#58;0px !important;font-size&#58;18px;padding&#58;0px 0px 0px 50px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Networked printers are managed through Printer Logic. This deploys drivers, settings and standardized naming.&#160;</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Printer Logic app is deployed via Intune to fully automate the setup from the start.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Naming convention for Managed Printers is &quot;Printer Name - Managed&quot;.&#160;</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Printer names are to be intuitive and avoid any meaningless codes. A user should see it and know exactly where it is. A label with the same name should be printed and fixed to the device.</span></li></ul></span><br></div></div>'
        }
        [PSCustomObject]@{
            Title = "3rd Party Paid Applications";
            Content = '<div class="ExternalClassC26605E02E6647A4992EE209FA796BF9"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><div style="color&#58;rgb(50, 49, 48);line-height&#58;1.4em;margin&#58;0px 0px 1.4em;font-size&#58;18px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><h3 style="margin&#58;0px 0px 12px;font-weight&#58;600 !important;font-size&#58;24px;">3rd Party Paid Applications</h3></div><ul style="color&#58;rgb(50, 49, 48);counter-reset&#58;item 0;line-height&#58;1.4em;margin-top&#58;0px;margin-bottom&#58;0px;margin-left&#58;0px !important;font-size&#58;18px;padding&#58;0px 0px 0px 50px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">We provide the operating system configuration needed to support the application</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Clients are to contact the vendor support for app-specific issues</span></li></ul></span><br></div></div>'
        }
        [PSCustomObject]@{
            Title = "User Experience Management";
            Content = '<div class="ExternalClassF90FE34844B942D3965A78F1469B8A95"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><div style="color&#58;rgb(50, 49, 48);line-height&#58;1.4em;margin&#58;0px 0px 1.4em;font-size&#58;18px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><h3 style="margin&#58;0px 0px 12px;font-weight&#58;600 !important;font-size&#58;24px;">User Experience Management</h3></div><ul style="color&#58;rgb(50, 49, 48);counter-reset&#58;item 0;line-height&#58;1.4em;margin-top&#58;0px;margin-bottom&#58;0px;margin-left&#58;0px !important;font-size&#58;18px;padding&#58;0px 0px 0px 50px;text-align&#58;left;background-color&#58;rgb(255, 255, 255);"><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Tailored desktop shortcuts</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">New users should be introduced to IC 360 directly and provided our manual to explain systems and services</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Line of business applications to be deployed by Intune</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Edge browser configured to auto-sync</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">OneDrive configured to auto sign in and sync&#160;desktop/docs/pictures folders</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Ensure all users are informed with a manual and training and orientation sessions unique to your environment for onboarding future staff</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Any company-wide browser settings are to be defined in Edge polices in EndPoint manager. If a team requires settings, they should have a Microsoft Team and that team can be used to target user-based policies.</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">Team owners should be trained and encouraged to manage Team memberships</span></li><li style="margin-bottom&#58;0px;"><span style="font-size&#58;16px;line-height&#58;1.4;">We never connect to a device without permission. All connections are logged in an audit trail.</span></li></ul></span><br></div></div>'
        }
   )

   try{
    
    Write-Host "Adding items to the IT Operations list" -ForegroundColor Yellow
    $batch = New-PnPBatch
    for($i=0; $i -lt $ITOperationsItems.Length ;$i++)
    {
        Add-PnPListItem -List "IT Operations" -Values @{"Title"= $ITOperationsItems.Title[$i] ; "Content" = $ITOperationsItems.Content[$i]} -Batch $batch
    }
    Invoke-PnPBatch -Batch $batch

    Write-Host "Items added to the IT Operations list" -ForegroundColor Green

   }
   catch {
    Write-Host   "Failed to add items to It Operations list" -ForegroundColor Red
    Write-Host   $_.Exception.Message -ForegroundColor Red
   }

   ##Adding items to the Procedures list
   $ProceduresItem = @(
       [PSCustomObject]@{
           Title = 'Subtitles';
           Content = '<div class="ExternalClass113F3495A13F41CDB624BF3883046554"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><div><span style="color&#58;black;font-size&#58;16pt;font-family&#58;verdana;"></span></div><ul><span style="font-size&#58;14pt;">1 Change Request</span><br><span style="font-size&#58;14pt;">2 Onboarding</span><br><span style="font-size&#58;14pt;">3 Offboarding</span><br><span style="font-size&#58;14pt;">4 Authorizations (purchases, licenses, ACts)</span><br><span style="font-size&#58;14pt;">5 Health Checks (backups, security, hardware)</span><br><span style="font-size&#58;14pt;">6 Workplan Review</span><br><span style="font-size&#58;14pt;">7 Incident Response</span><br><span style="font-size&#58;14pt;">8 Support Procedures - Support Tools</span><p><br></p><span style="background-color&#58;rgb(0, 255, 255);font-family&#58;verdana;font-size&#58;16pt;"></span></ul></div></div>'
       }
       [PSCustomObject]@{
           Title = 'Change Request Procedure';
           Content = '<div class="ExternalClass6878B6CD0497487882D52A5A0FF3F581"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><h2>Change Request (CR) Procedure</h2><ol><li><span>Complete CR Form</span></li><li><span>Review of CR by Technical Lead</span></li><li><span>Submit for Approval</span></li><li><span>Communications sent to staff if necessary&#160;</span></li><li><span>Testing Completed</span></li><li><span>Notification of Completion Sent</span></li><li><span>Monitor for Issues</span></li><li><span>Close Ticket</span></li></ol></span><br></div></div>'
       }
       [PSCustomObject]@{
           Title = 'Authorizations';
           Content = '<div class="ExternalClass649CF70A9BF346DE9DB8FB3F1E92A8F2"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"></span><span style="color&#58;black;font-size&#58;16pt;"><b>Authorizations<div style="margin-top&#58;21.3333px;margin-bottom&#58;21.3333px;"><table style="border-collapse&#58;collapse;border&#58;none;"> <tbody><tr style="height&#58;25.95pt;">  <td width="184" style="width&#58;138.25pt;border&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;25.95pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>&#160;</b></p>  </td>  <td width="208" style="width&#58;156.1pt;border&#58;solid windowtext 1.0pt;border-left&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;25.95pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>Name</b></p>  </td> </tr> <tr style="height&#58;40.9pt;">  <td width="184" style="width&#58;138.25pt;border&#58;solid windowtext 1.0pt;border-top&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;40.9pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>Purchases  (Licenses and Hardware)</b></p>  </td>  <td width="208" style="width&#58;156.1pt;border-top&#58;none;border-left&#58;none;border-bottom&#58;solid windowtext 1.0pt;border-right&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;40.9pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>&#160;</b></p>  </td> </tr> <tr style="height&#58;25.95pt;">  <td width="184" style="width&#58;138.25pt;border&#58;solid windowtext 1.0pt;border-top&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;25.95pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>User  Creation/Deletion</b></p>  </td>  <td width="208" style="width&#58;156.1pt;border-top&#58;none;border-left&#58;none;border-bottom&#58;solid windowtext 1.0pt;border-right&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;25.95pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>&#160;</b></p>  </td> </tr> <tr style="height&#58;24.5pt;">  <td width="184" style="width&#58;138.25pt;border&#58;solid windowtext 1.0pt;border-top&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;24.5pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>Password  Reset</b></p>  </td>  <td width="208" style="width&#58;156.1pt;border-top&#58;none;border-left&#58;none;border-bottom&#58;solid windowtext 1.0pt;border-right&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;24.5pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>&#160;</b></p>  </td> </tr></tbody></table><br></div></b></span><br></div></div>'
       }
       [PSCustomObject]@{
           Title = 'User Onboarding';
           Content = '<div class="ExternalClassF7261A13658444ADAE38F107CF9BD79F"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><h2>User Onboarding</h2><p><span>Generic Onboarding for MSP Clients</span></p><ol><li><span>Use details in the client ticket to creating account</span></li><li><span>Order Hardware&#160;</span></li><li><span>Assign Licenses</span></li><li><span>Add to Azure AD Groups</span></li><li><span>Send account login details to personal email</span></li><li><span>Prep of Hardware - Autopilot</span></li><li><span>Install LOB Apps</span></li></ol><br></div></div>'
       }
       [PSCustomObject]@{
           Title = 'User Offboarding';
           Content = '<div class="ExternalClass2931AC8270064F428CD841F5FD6F006A"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;font-size&#58;18pt;"><b></b></span><span style="color&#58;black;"><span style="font-size&#58;18pt;"><b>User Offboarding<br></b></span><span style="font-size&#58;14pt;">Generic Onboarding for MSP Clients</span><span><br></span></span></div><blockquote style="margin&#58;0 0 0 40px;border&#58;none;padding&#58;0px;"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span style="font-size&#58;14pt;">1. Use details in the client ticket to offboard account </span></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span style="font-size&#58;14pt;">2. Refresh Hardware - Wipe and Prep for next</span></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span style="font-size&#58;14pt;">3. Remove User from Groups</span></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span style="font-size&#58;14pt;">4. Remove from Licenses</span></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span style="font-size&#58;14pt;">5. LOB Apps- Remove Access</span></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span style="font-size&#58;14pt;">6. Convert Mailbox to Shared Mailbox - provide manager access</span></span></div></blockquote></div>'
       }
   )

   try{
    
    Write-Host "Adding items to the Procedures list" -ForegroundColor Yellow
    $batch = New-PnPBatch
    for($i=0; $i -lt $ProceduresItem.Length ;$i++)
    {
        Add-PnPListItem -List "Procedures" -Values @{"Title"= $ProceduresItem.Title[$i] ; "Content" = $ProceduresItem.Content[$i]} -Batch $batch
    }
    Invoke-PnPBatch -Batch $batch

    Write-Host "Items added to the Procedures list" -ForegroundColor Green

   }
   catch {
    Write-Host   "Failed to add items to Procedures list" -ForegroundColor Red
    Write-Host   $_.Exception.Message -ForegroundColor Red
   }

    ##Adding items to the Networking list
    $NetworkingItem = @(
        [PSCustomObject]@{
            Title = 'zone 3';
            Content = '<div class="ExternalClass03B7B1721D884E3B9C295242E8AA725D"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><table style="border-collapse&#58;collapse;border&#58;none;"> <tbody><tr style="height&#58;20.45pt;">  <td width="127" style="width&#58;95.45pt;border&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;20.45pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;">VPN Type</p>  </td>  <td width="255" style="width&#58;191.45pt;border&#58;solid windowtext 1.0pt;border-left&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;20.45pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;">&#160;</p>  </td> </tr> <tr style="height&#58;19.3pt;">  <td width="127" style="width&#58;95.45pt;border&#58;solid windowtext 1.0pt;border-top&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;19.3pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;">URL</p>  </td>  <td width="255" style="width&#58;191.45pt;border-top&#58;none;border-left&#58;none;border-bottom&#58;solid windowtext 1.0pt;border-right&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;19.3pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;">&#160;</p>  </td> </tr> <tr style="height&#58;20.45pt;">  <td width="127" style="width&#58;95.45pt;border&#58;solid windowtext 1.0pt;border-top&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;20.45pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;">Username</p>  </td>  <td width="255" style="width&#58;191.45pt;border-top&#58;none;border-left&#58;none;border-bottom&#58;solid windowtext 1.0pt;border-right&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;20.45pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;">&#160;</p>  </td> </tr> <tr style="height&#58;19.3pt;">  <td width="127" style="width&#58;95.45pt;border&#58;solid windowtext 1.0pt;border-top&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;19.3pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;">Password</p>  </td>  <td width="255" style="width&#58;191.45pt;border-top&#58;none;border-left&#58;none;border-bottom&#58;solid windowtext 1.0pt;border-right&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;19.3pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;">&#160;</p>  </td> </tr> <tr style="height&#58;20.45pt;">  <td width="127" style="width&#58;95.45pt;border&#58;solid windowtext 1.0pt;border-top&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;20.45pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;">Software</p>  </td>  <td width="255" style="width&#58;191.45pt;border-top&#58;none;border-left&#58;none;border-bottom&#58;solid windowtext 1.0pt;border-right&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;20.45pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;">&#160;</p>  </td> </tr> <tr style="height&#58;19.3pt;">  <td width="127" style="width&#58;95.45pt;border&#58;solid windowtext 1.0pt;border-top&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;19.3pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;">Token  Location</p>  </td>  <td width="255" style="width&#58;191.45pt;border-top&#58;none;border-left&#58;none;border-bottom&#58;solid windowtext 1.0pt;border-right&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;19.3pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;">&#160;</p>  </td> </tr> <tr style="height&#58;20.45pt;">  <td width="127" style="width&#58;95.45pt;border&#58;solid windowtext 1.0pt;border-top&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;20.45pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;">Notes</p>  </td>  <td width="255" style="width&#58;191.45pt;border-top&#58;none;border-left&#58;none;border-bottom&#58;solid windowtext 1.0pt;border-right&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;20.45pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;">&#160;</p>  </td> </tr></tbody></table></span><br></div></div>'
        }
        [PSCustomObject]@{
            Title = 'zone 1';
            Content = ''
        }
    )
 
    try{
     
     Write-Host "Adding items to the Networking list" -ForegroundColor Yellow
     $batch = New-PnPBatch
     for($i=0; $i -lt $NetworkingItem.Length ;$i++)
     {
         Add-PnPListItem -List "Networking" -Values @{"Title"= $NetworkingItem.Title[$i] ; "Content" = $NetworkingItem.Content[$i]} -Batch $batch
     }
     Invoke-PnPBatch -Batch $batch
 
     Write-Host "Items added to the Networking list" -ForegroundColor Green
 
    }
    catch {
     Write-Host   "Failed to add items to Networking list" -ForegroundColor Red
     Write-Host   $_.Exception.Message -ForegroundColor Red
    }
}

   ##Uploading file to document libraries
   try{
    Write-Host "Uploading visio file to document libraries"

    Add-PnPFile -Path '.\Template\Shared Documents\MP-PerimeterNetworkDiagram2.vsdx' -Folder "Network Library"  
    Add-PnPFile -Path '.\Template\Shared Documents\MP-Azure-O365 EndState.vsdx' -Folder "Home Library" 
    Add-PnPFile -Path .\Template\SiteAssets\__footerlogo__IC-360-Logo-V-1clr-Inverse.png -Folder "SiteAssets" 

    Write-Host "Uploaded visio file to document libraries"
   }
   catch {
    Write-Host   "Failed to add items to Procedures list" -ForegroundColor Red
    Write-Host   $_.Exception.Message -ForegroundColor Red
   }

   try{
        ##Get Owners group Title
        $OwnersGroupTitle = ""
        $MembersGroupTitle = ""
        $VisitorsGroupTitle = ""
        $AllGroups = Get-PnPGroup
        for($i =0; $i -lt $AllGroups.Length; $i++) 
        { 
            if($AllGroups[$i].Title -match "Owners")
            { 
                $OwnersGroupTitle = $AllGroups[$i].Title 
                } 
                if($AllGroups[$i].Title -match "Members")
            { 
                $MembersGroupTitle = $AllGroups[$i].Title 
                }
                if($AllGroups[$i].Title -match "Visitors")
            { 
                $VisitorsGroupTitle = $AllGroups[$i].Title 
                }
            }
    
        Write-Host "Permissions provisioning to Site Pages" -ForegroundColor Yellow

        #Provision permission to site pages and publish them
        $SitePagesItems = (Get-PnPListItem -List "Site Pages" -Fields "Name","Title", "ID").FieldValues
        for($j = 0; $j -lt $SitePagesItems.Length; $j++){
            if($SitePagesItems[$j].FileLeafRef -eq "Networking.aspx" -Or $SitePagesItems[$j].FileLeafRef -eq "Systems--Apps.aspx" -Or $SitePagesItems[$j].FileLeafRef -eq "IT-Operations.aspx" -Or $SitePagesItems[$j].FileLeafRef -eq "Procedures.aspx "){
                Set-PnPListItemPermission -List 'Site Pages' -Identity $SitePagesItems[$j].ID -ClearExisting -Group $OwnersGroupTitle
                Set-PnPListItemPermission -List 'Site Pages' -Identity $SitePagesItems[$j].ID -Group $OwnersGroupTitle -AddRole $Config.Parameters.FullControlPermission
                Set-PnPListItemPermission -List 'Site Pages' -Identity $SitePagesItems[$j].ID -Group $MembersGroupTitle -AddRole $Config.Parameters.ReadPermission
                Set-PnPListItemPermission -List 'Site Pages' -Identity $SitePagesItems[$j].ID -Group $VisitorsGroupTitle -AddRole $Config.Parameters.ReadPermission

                Set-PnPPage -Identity $SitePagesItems[$j].FileLeafRef -Publish
            }
        }

        Set-PnPPage -Identity "Home" -Publish

        Write-Host "Permission provisioned to all the Site Pages and Published them" -ForegroundColor Green
    }
    catch {
        Write-Host   "Failed to provision permissions to Site pages and publishing them" -ForegroundColor Red
        Write-Host   $_.Exception.Message -ForegroundColor Red
    }

    #endregion
}

END {
    # Set-PnPTraceLog -Off
}



