# Teams self-service request script
# Author: Kdships (https://insterswap.wordpress.com/)
# Note: The reason for the multiple start-sleep entries is due to a known feedback failures from Microsoft. 
# Adding some delays helped fix the issue. Microsoft have been informed about the PS bug.
# Remember to assign your credentials to $UserCredential
$SiteURL = "https://XXXXXXX.sharepoint.com/sites/XXXXXXXXXX"
$TeamListName = "XXXXXXXXXX"

#Set more variables
$FailedtoProvision = @()
$Today3 = Get-date
$Today4 = $Today3.ToShortDateString()
$Today5 = $Today4 -replace '/','_'
Write-host "Logging started: $Today3"

####Connect to site via PnP Online
Connect-PnPOnline -Url $SiteURL -Credentials $UserCredential -Verbose
Start-Sleep -Seconds 1

####Connect to AzureAD
Connect-AzureAD -Credential $UserCredential -Verbose
Start-Sleep -Seconds 1

####Connect to MS Teams
Import-module -Name MicrosoftTeams #-Verbose
#Start-Sleep -Seconds 1
Connect-MicrosoftTeams -Credential $UserCredential -Verbose
Start-Sleep -Seconds 1

#Capture list items and store in ListEntries
$ListEntries = @()
$Entries = $Null
$AllGroups = $Null
$Entries = Get-PnpListItem -List $TeamListName
if($Entries.count -gt 0)
{
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber -Verbose
$Blockedwords = $Null #New
$Blockedwords = import-csv "<Insert FilePath to blockedwords here_Header is BlockedWords>"
foreach ($Entry in $Entries)
{
#Set variable
$VisibilityType = $Null
$Description = $Null
$RequestorsEmail = $Null
$MergewithO365Group = $Null
$O365GroupNameorEmail = $Null
$EntryID = $Null
$TeamsName = $Null
$GivenName = $Null
$Array = $Null #New
$WordCheck = 0 #New
$word = $Null #New
$UPN = $Null
$Name = $Null
$Name = $entry.Fieldvalues.Title
$Name5 = $Name -replace '[\W]', ''
$VisibilityType = $entry.Fieldvalues.VisibilityType
$Description = $entry.Fieldvalues.Description
$RequestorsEmail = $entry.Fieldvalues.Author.Email
$MergewithO365Group = $entry.Fieldvalues.MergewithO365Group
$O365GroupNameorEmail = $entry.Fieldvalues.O365GroupNameorEmail
$EntryID = $entry.Fieldvalues.ID
$UPN = (Get-mailbox -identity $RequestorsEmail).userprincipalname
$GivenName = (Get-AzureADUser -ObjectId $UPN).GivenName
$TeamsName = $Name5
$Array = $Name.Split(" ") #New
Foreach($word in $Array)
{
  $Wordvalue = $Null
  $Wordvalue = $word -replace ' ', ''
  if ($Blockedwords.blockedwords -contains $Wordvalue)
  {
  $WordCheck++
  }
}
    If($WordCheck -eq 0)
    {
    #Email status of the request to the requestor
    $UriT = $Null
    $StatusValue = $Null
    $EntryIDNull = $EntryID
    $EntryIDNull = $Null
    $UriTNote = "<Insert Invoke PowerAutomate URL Path HERE>"
    [int]$StatusValueNote = "1"
            $BodyTNote = [ordered] @{
            EntryID = $EntryIDNull
            email   = $RequestorsEmail
            checkvalue = $StatusValue
            subject = "Your Microsoft Team request is being processed"
            bodyvalue    = "Your Microsoft Team request is being processed.
            
            Team Name: $Name. 
            
            Privacy Preference: $VisibilityType."
            firstname = "Hi $GivenName"
          } | ConvertTo-Json -Depth 10
    Invoke-RestMethod -Method POST -Uri $UriTNote -Body $BodyTNote -ContentType application/json
    if($MergewithO365Group -eq "Yes")
    {
    Start-Sleep -Seconds 3
    #Email status of the request to the requestor, email request details to IT team and delete request from SharePoint
    $UriT = $Null
    $StatusValue = $Null
    $UriT = "Insert Invoke PowerAutomate URL Path HERE"
    [int]$StatusValue = "1"
            $BodyT = [ordered] @{
            EntryID = $EntryID
            email   = $RequestorsEmail
            checkvalue = $StatusValue
            subject = "Your Team provisioning request has been received"
            bodyvalue    = "Your request to merge '$O365GroupNameorEmail' Office 365 Group with your Microsoft Team, $Name, has been assigned to IT team. A member of the team will be in touch as soon as the merge is complete."
            firstname = "Hi $GivenName"
          } | ConvertTo-Json -Depth 10
    Invoke-RestMethod -Method POST -Uri $UriT -Body $BodyT -ContentType application/json
    $Subject = "Action Required: A request to create Teams with an O365 group merge"
    $Body = $Null
    $Body = "Hi Team,
    
    Please create MS Team on my behalf and merge it with my existing O365 Group. Please find details below.
    Team Name: $Name
    Team NickName/Alias: $TeamsName
    Visibility: $VisibilityType
    Description: $Description
    Existing O365 Group: [$O365GroupNameorEmail]
    
    If you have any questions, please let me know.
    
    Thank you
    
    $GivenName
    
    $RequestorsEmail
    $UPN
    
    Sent by a trigger for 'TeamsProvisioningSelfServiceApp' on behalf of $UPN"
    Send-MailMessage -From $RequestorsEmail -to "XXXXXXXX" -Subject $Subject -Body $Body -priority High -SmtpServer XXXXXX
    Write-host "Team and Group merge request for $TeamsName by $UPN has been sent to the IT team" -ForegroundColor Green
    }
    elseif ($MergewithO365Group -eq "No")
    {
    #Check if this Team already exist
    $Check1 = $Null
    $Check2 = $Null
    $Check3 = $Null
    $Check4 = $Null
    $Check1 = Get-Recipient -Identity $TeamsName
    Start-Sleep -Seconds 3
    $Check2 = Get-Recipient -Identity ($TeamsName + '@Company.com')
    Start-Sleep -Seconds 3
    $Check3 = Get-Recipient -Identity ($TeamsName + '@Company.onmicrosoft.com')
    Start-Sleep -Seconds 3
    $Check4 = Get-Recipient -Identity "$Name"
    if (($Check1 -eq $Null) -and ($Check2 -eq $Null) -and ($Check3 -eq $Null) -and ($Check4 -eq $Null))
    {
    Start-Sleep -Seconds 10
                    If($VisibilityType -eq "Private - Only Team owners can add members")
                    {
                $VisibilityType = "Private"
                }
                    else
                    {
                $VisibilityType = "Public"
                }
        Try {
            #Create Team
            $ReadyTeam = $Null
            $RunCheck = 0
            $ReadyTeam = New-Team -MailNickname $TeamsName -displayname $Name -Visibility $VisibilityType -Description $Description
            $RunCheck = 1
            Start-Sleep -Seconds 10
            }
        Catch {
                #Email status of the request to the requestor, email request details to IT team and delete request from SharePoint
                $UriT = $Null
                $StatusValue = $Null
                $UriT = "https://prod-121.REGION.logic.azure.com:443/workflows/FLOWPATHID"
                [int]$StatusValue = "3"
                        $BodyT = [ordered] @{
                        EntryID = $EntryID
                        email   = $RequestorsEmail
                        checkvalue = $StatusValue
                        subject = "Your Team request is pending - $Name"
                        bodyvalue    = "Your Team provisioning request is pending. A member of the IT team will be in touch as soon as it is rectified."
                        firstname = "Hi $GivenName"
                      } | ConvertTo-Json -Depth 10
                Invoke-RestMethod -Method POST -Uri $UriT -Body $BodyT -ContentType application/json
                $Subject = "Action Required: Failed to create Team - $Name"
                $Body = $Null
                $Body = "Hi Team,
    
                MS Teams provisioning self-service app failed to create this Team. Using the details below, please kindly retry creatng the Team through the self-service app and update the ownership. Remember to remove yourself as the owner.
    
                Self-Service App Link: https://apps.powerapps.com/play/APPID
                Team Display Name: $Name
                Team NickName/Alias: $TeamsName
                Visibility preference: $VisibilityType
                Description: $Description
                Requestor's Email Address: $RequestorsEmail
                Requestor's UPN: $UPN
                Reason for the failure: $_
    
                Sent by a trigger for 'TeamsProvisioningSelfServiceApp' task"
                Send-MailMessage -From "TeamsSelfServiceApp@Company.com" -to "Company@ticketing.com" -Subject $Subject -Body $Body -priority normal -SmtpServer email.Company.com
                }
        If($RunCheck -eq 1)
        {       #Add the requestor as owner of the Team and remove the service account
                Start-Sleep -Seconds 13
                Try{
                $RunCheckB = 0
                Add-TeamUser -GroupId $ReadyTeam.GroupId -User $UPN -Role Owner
                $RunCheckB = 1
                Start-Sleep -Seconds 3
                #Email status of the request to the requestor, email request details to IT team and delete request from SharePoint
                    $UriT = $Null
                    $StatusValue = $Null
                    $UriT = "https://prod-121.REGION.logic.azure.com:443/workflows/FLOWPATHID"
                    [int]$StatusValue = "2"
                            $BodyT = [ordered] @{
                            EntryID = $EntryID
                            email   = $RequestorsEmail
                            checkvalue = $StatusValue
                            subject = "[$Name] Team has been successfully created"
                            bodyvalue    = "[$Name] Team has been created. Shortly, it will be visible within your Teams client. If you do not find it, please scroll to the bottom of your Teams list.
                    
                            Team Name: $Name. 
            
                            Privacy Preference: $VisibilityType.
                    
                            If you have any questions, please reach out to IT team."
                            firstname = "Hi $GivenName"
                          } | ConvertTo-Json -Depth 10
                    Invoke-RestMethod -Method POST -Uri $UriT -Body $BodyT -ContentType application/json
                    Write-host "$TeamsName was successfully created for $UPN" -ForegroundColor Green
                    Send-MailMessage -From "TeamsSelfServiceApp@Company.com" -to "kdships@Company.com" -Subject "New Teams Request by $UPN" -Body "$TeamsName - New Teams Request by $RequestorsEmail | $UPN" -priority Low -SmtpServer email.Company.com
                }
                Catch{
                #Email status of the request to the requestor, email request details to IT team and delete request from SharePoint
                    $UriT = $Null
                    $StatusValue = $Null
                    $UriT = "https://prod-121.REGION.logic.azure.com:443/workflows/FLOWPATHID"
                    [int]$StatusValue = "2"
                            $BodyT = [ordered] @{
                            EntryID = $EntryID
                            email   = $RequestorsEmail
                            checkvalue = $StatusValue
                            subject = "[$Name] Team has been successfully created but pending one final step"
                            bodyvalue    = "[$Name] Team has been created. We are applying finishing touches and would let you know as soon as it is ready.
                    
                            Team Name: $Name. 
            
                            Privacy Preference: $VisibilityType.
                    
                            If you have any questions, please reach out to IT team."
                            firstname = "Hi $GivenName"
                          } | ConvertTo-Json -Depth 10
                    Invoke-RestMethod -Method POST -Uri $UriT -Body $BodyT -ContentType application/json
                $Subject = "Action Required: Failed to create Team owner"
                $Body = $Null
                $Body = "Hi Team,
    
                MS Teams provisioning self-service app failed to add the owner to the team. Using the details below, kindly add the owner through Teams admin portal or through the Client. Please remember to remove yourself and Collab service account from the Team.
    
                Team Display Name: $Name
                Requestor's Email Address: $RequestorsEmail
                Requestor's UPN: $UPN
                Reason for the failure: $_
    
                Sent by a trigger for 'TeamsProvisioningSelfServiceApp' task"
                Send-MailMessage -From "TeamsSelfServiceApp@Company.com" -to "Company@ticketing.com" -Subject $Subject -Body $Body -priority normal -SmtpServer email.Company.com
                Send-MailMessage -From "TeamsSelfServiceApp@Company.com" -to "kdships@Company.com" -Subject $Subject -Body $Body -priority High -SmtpServer email.Company.com
                Write-host "Adding owner to $Name failed - Owner: $UPN - Reason: $_" -ForegroundColor Red
                }
                Finally{
                If($RunCheckB -eq 1){
                Start-Sleep -Seconds 3
                Remove-TeamUser -GroupId $ReadyTeam.GroupId -User serviceaccount@Company.com -Role Owner
                Start-Sleep -Seconds 3
                Remove-TeamUser -GroupId $ReadyTeam.GroupId -User serviceaccount@Company.com}
                }
        }
        elseif($RunCheck -eq 0)
        {
        Send-MailMessage -From "TeamsSelfServiceApp@Company.com" -to "kdships@Company.com" -Subject $Subject -Body $Body -priority High -SmtpServer email.Company.com
        Write-host "Failed to complete processing Team for $UPN. Reason: $_" -ForegroundColor Red
        }             
    }
    elseif (($Check1 -ne $Null) -or ($Check2 -ne $Null) -or ($Check3 -ne $Null) -or ($Check4 -ne $Null))
    {
    #Check who owns the team
    $Owners = 0
    $Owners = (Get-UnifiedGroup -Identity "$Name").ManagedBy
        If($Owners.count -gt 0)
                {
                #Email status of the request to the requestor and delete request from SharePoint
                $UriT = $Null
                $StatusValue = $Null
                $UriT = "https://prod-121.REGION.logic.azure.com:443/workflows/FLOWPATHID"
                [int]$StatusValue = "4"
                        $BodyT = [ordered] @{
                        EntryID = $EntryID
                        email   = $RequestorsEmail
                        checkvalue = $StatusValue
                        subject = "Failed to create $Name Team - Reason: Name Conflict"
                        bodyvalue    = "
                        Your Team request failed due to a conflict with an existing Team or Group. The name you provided is already in use by: $Owners. 
            
                        Please change the name and try again."
                        firstname = "Hi $GivenName"
                      } | ConvertTo-Json -Depth 10
                Invoke-RestMethod -Method POST -Uri $UriT -Body $BodyT -ContentType application/json
                }
                else
                {
                #Email status of the request to the requestor and delete request from SharePoint
                $UriT = $Null
                $StatusValue = $Null
                $UriT = "https://prod-121.REGION.logic.azure.com:443/workflows/FLOWPATHID"
                [int]$StatusValue = "4"
                        $BodyT = [ordered] @{
                        EntryID = $EntryID
                        email   = $RequestorsEmail
                        checkvalue = $StatusValue
                        subject = "Failed to create $Name Team - Reason: Name Conflict"
                        bodyvalue    = "
                        Your Team request failed due to a conflict with an existing object. Please change the name and try again."
                        firstname = "Hi $GivenName"
                      } | ConvertTo-Json -Depth 10
                Invoke-RestMethod -Method POST -Uri $UriT -Body $BodyT -ContentType application/json
                }
    Write-host "Failed to create Team due to a name conflict - $Name" -ForegroundColor Red
    }
}
}
    else
    {
    #Email status of the request to the requestor and delete request from SharePoint
    $UriT = $Null
    $StatusValue = $Null
    $UriT = "<Insert Invoke PowerAutomate URL Path HERE>"
    [int]$StatusValue = "4"
            $BodyT = [ordered] @{
            EntryID = $EntryID
            email   = $RequestorsEmail
            checkvalue = $StatusValue
            subject = "Failed to create $Name Team - Reason: Restricted keyword"
            bodyvalue    = "
            Your Team request ($Name) failed due to a restricted keyword. Please change the name and try again." 
            firstname = "Hi $GivenName"
          } | ConvertTo-Json -Depth 10
    Invoke-RestMethod -Method POST -Uri $UriT -Body $BodyT -ContentType application/json
    Write-host "Failed to create Team due to a restricted keyword -$Name" -ForegroundColor Red
    }
}
}
elseif ($Entries.count -eq 0)
{
Write-host "No pending request"
}
$Today3 = Get-date
Write-host "Logging stopped: $Today3" 
#Stop-Transcript
# Delete all Files older than 7 day(s) in XXXX
$Path = "PATH TO LOG FILE"
$Dayscount = "-7"
$Deletedate = $Today3.AddDays($Dayscount)
Get-ChildItem $Path | Where-Object { $_.LastWriteTime -lt $Deletedate } | Remove-Item
