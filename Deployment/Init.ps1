[CmdletBinding()]
Param(
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
    #     Get-PnPSiteTemplate -Out $Config.Parameters.ExtractedTemplatePath -IncludeAllPages -PersistBrandingFiles -PersistPublishingFiles
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
       Invoke-PnPSiteTemplate -Path .\PnPTemplate.xml -EA Stop
       Write-Host "PnP Template applied to the site collection" -ForegroundColor Green
   }
   
   catch {
       Write-Host   "Failed to Create Site and applying template" -ForegroundColor Red
       Write-Host   $_.Exception.Message -ForegroundColor Red
   }

   
   ##Add Ittems to the It Operations list

   $ITOperationsItems = @(
       [PSCustomObject]@{
           Title = "PC Summary";
           Content = '<div class="ExternalClassBD9C13EBB1244577B01EBE6E7B82E9A1"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="font-size&#58;18pt;color&#58;black;"><b></b></span><span style="font-size&#58;18pt;color&#58;black;"><b></b></span><span style="font-size&#58;18pt;"><b>PC Summary</b></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="font-size&#58;18pt;"><div style="margin-top&#58;24px;margin-bottom&#58;24px;"><ul><li><span style="font-size&#58;12pt;">All PCs are managed with Intune this includes Compliance Policies, Application Push and Company</span></li><li><span style="font-size&#58;12pt;">Portal, AutoPilot for onboarding of new PCs&#160;</span></li><li><span style="font-size&#58;12pt;">All PCs have Defender ATP deployed and managed, integrated with Intune.</span></li></ul></div></span></div></div>';
       }
       [PSCustomObject]@{
        Title = "IT Operating Systems";
        Content = '<div class="ExternalClass059753BB40F84A27B3BEB130D81D9F39"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;font-size&#58;16pt;"><b>IT Operating Systems</b></span><span style="color&#58;black;"><div style="margin-top&#58;14.6667px;margin-bottom&#58;14.6667px;"><ul><li>Endpoints</li><ul><li><span style="font-size&#58;12pt;">MM N-central agent is deployed to all PCs for monitoring © Compliance Plies are Applied and mentored an the</span></li><li><span style="font-size&#58;12pt;">Daas Enrolled PCs use the HP DasS Insight portal for monitoring and reporting © lune baseline securty polices ae applied fr hardening af the OS</span></li><li><span style="font-size&#58;12pt;">Encryption - via BitLocker and Intune</span></li></ul></ul><br></div></span><br></div></div>'
       }
       [PSCustomObject]@{
           Title = "Mobile Devices";
           Content = '<div class="ExternalClassC3E41CC8AB23457C955C445F360F9483"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;font-size&#58;16pt;"><b></b></span><span style="color&#58;black;font-size&#58;16pt;"><b>Mobile Devices</b></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><div style="margin-top&#58;14.6667px;margin-bottom&#58;14.6667px;"><ul><li><span style="font-size&#58;12pt;">Intune is used to manage these devices, they are classified into two types</span></li><ul><li><span style="font-size&#58;12pt;">Corp Owned - These are owned by the firm and fully managed Printers;</span></li><li><span style="font-size&#58;12pt;">&#160;Personal BYOD - These are owned by the end-user and only the applications are managed</span></li><li><span style="font-size&#58;12pt;">and secured (0365 Apps)</span></li></ul></ul></div></span><br></div></div>'
       }
       [PSCustomObject]@{
           Title = 'Printers';
           Content = '<div class="ExternalClass649CF70A9BF346DE9DB8FB3F1E92A8F2"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"></span><span style="color&#58;black;font-size&#58;16pt;"><b>Printers</b></span><br></div></div>'
       }
       [PSCustomObject]@{
           Title = 'Network';
           Content = '<div class="ExternalClass2794730D58734EADBC4564CD397EAB16"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"></span><span style="color&#58;black;font-size&#58;18pt;"><b>Network</b></span><br></div></div>'
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
           Content = '<div class="ExternalClassDBCD8477DBE440ED93A0E102475D715C"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="font-size&#58;18pt;"><b>Change Request Procedure<br></b></span><span style="font-size&#58;14pt;">1. Fill in form</span><br><span style="font-size&#58;14pt;">2. Review of CR by Technical Lead</span><br><span style="font-size&#58;14pt;">3. Submit for Approval</span><br><span style="font-size&#58;14pt;">4. Communications sent to staff if nesseary</span><br><span style="font-size&#58;14pt;">5. Complete CF and Review CR form</span><br><span style="font-size&#58;14pt;">6. Testing Completed</span><br><span style="font-size&#58;14pt;">7. Notification of Completion sent</span><br><span style="font-size&#58;14pt;">8. Monitor for Issues</span><br><span style="font-size&#58;14pt;">9. Close Out ticket</span> <br></div></div>'
       }
       [PSCustomObject]@{
           Title = 'Authorizations';
           Content = '<div class="ExternalClass649CF70A9BF346DE9DB8FB3F1E92A8F2"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"></span><span style="color&#58;black;font-size&#58;16pt;"><b>Authorizations<div style="margin-top&#58;21.3333px;margin-bottom&#58;21.3333px;"><table style="border-collapse&#58;collapse;border&#58;none;"> <tbody><tr style="height&#58;25.95pt;">  <td width="184" style="width&#58;138.25pt;border&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;25.95pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>&#160;</b></p>  </td>  <td width="208" style="width&#58;156.1pt;border&#58;solid windowtext 1.0pt;border-left&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;25.95pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>Name</b></p>  </td> </tr> <tr style="height&#58;40.9pt;">  <td width="184" style="width&#58;138.25pt;border&#58;solid windowtext 1.0pt;border-top&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;40.9pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>Purchases  (Licenses and Hardware)</b></p>  </td>  <td width="208" style="width&#58;156.1pt;border-top&#58;none;border-left&#58;none;border-bottom&#58;solid windowtext 1.0pt;border-right&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;40.9pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>&#160;</b></p>  </td> </tr> <tr style="height&#58;25.95pt;">  <td width="184" style="width&#58;138.25pt;border&#58;solid windowtext 1.0pt;border-top&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;25.95pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>User  Creation/Deletion</b></p>  </td>  <td width="208" style="width&#58;156.1pt;border-top&#58;none;border-left&#58;none;border-bottom&#58;solid windowtext 1.0pt;border-right&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;25.95pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>&#160;</b></p>  </td> </tr> <tr style="height&#58;24.5pt;">  <td width="184" style="width&#58;138.25pt;border&#58;solid windowtext 1.0pt;border-top&#58;none;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;24.5pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>Password  Reset</b></p>  </td>  <td width="208" style="width&#58;156.1pt;border-top&#58;none;border-left&#58;none;border-bottom&#58;solid windowtext 1.0pt;border-right&#58;solid windowtext 1.0pt;padding&#58;0in 5.4pt 0in 5.4pt;height&#58;24.5pt;">  <p style="margin&#58;0in 0in 8pt;line-height&#58;107%;font-size&#58;11pt;font-family&#58;Calibri, sans-serif;margin-bottom&#58;0in;line-height&#58;normal;"><b>&#160;</b></p>  </td> </tr></tbody></table><br></div></b></span><br></div></div>'
       }
       [PSCustomObject]@{
           Title = 'User Onboarding';
           Content = '<div class="ExternalClass1EC12EEEABFF48BDBA80F1BC47BC4FA1"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;font-size&#58;18pt;"><b></b></span><span style="color&#58;black;font-size&#58;18pt;"><b></b></span><span style="color&#58;black;"><span style="font-size&#58;18pt;"><b>User Onboarding</b></span><span> </span></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span><p><span style="font-size&#58;14pt;">Generic Onboarding for MSP Clients&#160;</span></p></span></span></div></div><blockquote style="margin&#58;0 0 0 40px;border&#58;none;padding&#58;0px;"><div class="ExternalClass1EC12EEEABFF48BDBA80F1BC47BC4FA1"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span><p><span style="font-size&#58;14pt;">1. Use details in the client ticket to creating account&#160;</span></p></span></span></div></div></blockquote><div class="ExternalClass1EC12EEEABFF48BDBA80F1BC47BC4FA1"><blockquote style="margin&#58;0 0 0 40px;border&#58;none;padding&#58;0px;"><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span><p><span style="font-size&#58;14pt;">2. Order Hardware&#160;</span></p></span></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span><p><span style="font-size&#58;14pt;">3. Assign Licenses&#160;</span></p></span></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span><p><span style="font-size&#58;14pt;">4. Add to Azure AD Groups&#160;</span></p></span></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span><p><span style="font-size&#58;14pt;">5. Send account login details to personal email</span></p></span></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span><p><span style="font-size&#58;14pt;">&#160;6. Prep of Hardware - Autopilot</span></p></span></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span><p><span style="font-size&#58;14pt;">&#160;7.</span></p></span></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span><p><span style="font-size&#58;14pt;">&#160;8. LOB - ?</span></p></span></span></div><div style="font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(0, 0, 0);"><span style="color&#58;black;"><span><p><span style="font-size&#58;14pt;">&#160;9.</span></p></span></span></div></blockquote></div>'
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

   ##Uploading file to document libraries
   try{
    Write-Host "Uploading visio file to document libraries"

    Add-PnPFile -Path .\Match-Invoice-Process-Flow-Chart.vsd -Folder "Network Library" 
    Add-PnPFile -Path .\Match-Invoice-Process-Flow-Chart.vsd -Folder "Home Library" 

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

        #Provision permission to site pages
        for($j = 2; $j -lt 7; $j++){
                Set-PnPListItemPermission -List 'Site Pages' -Identity $j -ClearExisting -Group $OwnersGroupTitle
                Set-PnPListItemPermission -List 'Site Pages' -Identity $j -Group $OwnersGroupTitle -AddRole $Config.Parameters.FullControlPermission
                Set-PnPListItemPermission -List 'Site Pages' -Identity $j -Group $MembersGroupTitle -AddRole $Config.Parameters.ReadPermission
                Set-PnPListItemPermission -List 'Site Pages' -Identity $j -Group $VisitorsGroupTitle -AddRole $Config.Parameters.ReadPermission
        }

        Write-Host "Permission provisioned to all the Site Pages" -ForegroundColor Green
    }
    catch {
        Write-Host   "Failed to provision permissions to Site pages" -ForegroundColor Red
        Write-Host   $_.Exception.Message -ForegroundColor Red
    }

    #endregion
}

END {
    # Set-PnPTraceLog -Off
}



