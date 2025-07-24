function Reset-UserForwardOOO { 
<#
.SYNOPSIS
This command will be run when a user is terminated.
.DESCRIPTION
The Reset-UserForward commandlet will add the person to the O365-Exchange-Online group, enable the mailbox, hide the email address from the address list,
connect to exchange online, set retention policy to 90 days, set the forwarding address, match the exchange GUID, show the forwarding address, schedule a Job
to remove the mailbox 30 days henceforth.
.PARAMETER TUDName
Get the Display Name for the Termed User and paste/type here.
.PARAMETER TUUPN
Get the User logon name (UPN) from Active Directory and paste/type here.
.PARAMETER RUDName
Get the Display Name for the User receiving forwarded email from Active Directory and paste/type here.
.PARAMETER RUUPN
Get the User Logon name (UPN) for the user recieving forwarded email from Active Directory and paste/type here.
.PARAMETER TaskNumber
Get the TaskNumber from the SOM ticket and paste/type here.
.PARAMETER Requester
Enter the first name of the person makeing the request. Use the first name only
.PARAMETER RequesterUPN
Enter the User logon name (UPN) of the person making the request. 
.EXAMPLE
Reset-UserForward
Will prompt for 
.EXAMPLE
Reset-UserForward -TUDName Mary Jones -TUUPN mjones -RUDName Billy Smith -RUUPN bsmith -TaskNumber TASK1234567 -Requester Jane -RequesterUPN jdoe
Enter the TUDName TUUPN RUDName RUUPN TaskNumber Requester RequesterUPN parameters
.INPUTS
Types of objects input
.OUTPUTS
Types of objects returned
.NOTES
My notes.
.LINK
http://
.COMPONENT
.ROLE
.FUNCTIONALITY
#>
    [alias('ruf')]
    [cmdletbinding()]
    Param(
        [parameter(Mandatory=$True,HelpMessage="Enter termed user's Display Name without quotes, like: Mary Jones")]
        [string]$TUDName = (Read-Host "Enter Display Name without quotes, like: Mary Jones"),

        [parameter(Mandatory=$True,HelpMessage="Enter the termed user's UPN without quotes, like: mjones")]
        [string]$TUUPN = (Read-Host "Enter User Logon Name without quotes, like: mjones"),

        [parameter(Mandatory=$True,HelpMessage="Enter the receiving user's Display Name without quotes, like: Mary Jones")]
        [string]$RUDName = (Read-Host "Enter Recieving users Display Name without quotes, like: Mary Jones"),

        [parameter(Mandatory=$True,HelpMessage="Enter the receiving user's UPN without quotes, like mjones")]
        [string]$RUUPN = (Read-Host "Enter the receiving User Logon Name without quotes, like mjones"),
        
        [parameter(Mandatory=$True,HelpMessage="Enter the Task Number without quotes, like TASK1234567")]
        [string]$TaskNumber = (Read-Host "Enter the Task Number without quotes, like TASK1234567"),

        [parameter(Mandatory=$True,HelpMessage="Enter the name of the user requesting the action without quotes, like Mary Jones")]
        [string]$Requester = (Read-Host "Enter the Task Number without quotes, like Mary Jones"),

        [parameter(Mandatory=$True,HelpMessage="Enter the UPN of the user requesting the action, like MJones")]
        [string]$RequesterUPN = (Read-Host "Enter the Task Number without quotes, like MJones"),

        [parameter(Mandatory=$True,HelpMessage="Enter the Text of the OOO reply, like I am no longer with Your Corp, please contact Joe Cool at joecool@yourcorp.com")]
        [string]$OOO = (Read-Host "Enter the Text of the OOO reply, like I am no longer with Your Corp, please contact Joe Cool at joecool@yourcorp.com")
        
        )       
        
    #Connect to ExchangeOnPrem
    Powershell.exe -executionpolicy remotesigned -File "C:\Users\yourname\OneDrive - Your Corp\Documents\Joe\ConnectToExchangeOnPrem.ps1"  

    #Connect to Exchange online
    $CertThumbPrnt = import-clixml "c:\KeyPath\CertificateThumbPrint.xml"
    $EXOAppID = import-clixml "c:\KeyPath\EXOAppID.xml"
    Connect-ExchangeOnline -CertificateThumbPrint "$CertThumbPrnt" -AppID "$EXOAppID" -Organization "kbonline.onmicrosoft.com"       
     
    #These commands will ensure the termed user's account is added to the KBOrgs-O365-Exchange-Online group.
    Add-ADGroupMember -Identity KBOrgs-O365-Exchange-Online -Members $TUUPN
    Enable-RemoteMailbox -Identity "$TUUPN@yourcorp.com" -RemoteRoutingAddress "$TUUPN@kbonline.mail.onmicrosoft.com"

    #Enable Archive
    Enable-RemoteMailbox -Identity "$TUUPN@yourcorp.com" -Archive
    
    #Hide name from the address list.
    Set-ADUser -Identity "$TUUPN" -Add @{msExchHideFromAddressLists="TRUE"}
        
    #set the retention policy
    Set-Mailbox -Identity "$TUUPN@yourcorp.com" -RetentionPolicy "90 Day Delete"

    #set the mailbox GUID
    $ExGuid = (Get-Mailbox -Identity "$TUUPN@yourcorp.com").ExchangeGUID
    $ExGuid
    do {
        Start-Sleep -Seconds 5
        Set-RemoteMailbox -Identity "$TUUPN@yourcorp.com" -ExchangeGuid $ExGuid -DomainController lmnopdc01.wxyz.corp.yourcorp.com
        $RExGuid = (Get-RemoteMailbox -Identity "$TUUPN@yourcorp.com").ExchangeGUID
        } until ($RExGuid -eq $ExGuid)

    Write-Host "ExchangeGUID has been updated."

    #set the archivce GUID
    $ArchGuid = (Get-Mailbox -Identity "$TUUPN@yourcorp.com").ArchiveGuid
    $ArchGuid
    do {
        Start-Sleep -Seconds 5
        Set-RemoteMailbox -Identity "$TUUPN@yourcorp.com" -ArchiveGuid $ArchGuid -DomainController lmnopdc01.wxyz.corp.yourcorp.com
        $RArchGuid = (Get-RemoteMailbox -Identity "$TUUPN@yourcorp.com").ArchiveGuid
        } until ($RArchGuid -eq $ArchGuid)

    Write-Host "ArchiveGUID has been updated."
    
    #set the forwarding address.
    Set-Mailbox -Identity "$TUUPN@yourcorp.com" -ForwardingAddress "$RUUPN@yourcorp.com" -DeliverToMailboxAndForward $true

    #display the results, show the SMTP address and forwarding address.
    Get-ADUser $TUUPN -Properties proxyAddresses
    Get-Mailbox -Identity "$TUUPN@yourcorp.com" | Select-Object -Property ForwardingAddress,ForwardingSmtpAddress | Out-String -Stream | sort
    Get-EXOMailbox -Identity "$UserLN@yourcorp.com" -Properties Name,ExchangeGuid
    Get-RemoteMailbox -Identity "$UserLN@yourcorp.com"  | select Name,ExchangeGuid
    
    #create PS1 with the Value only of the user. used ::Create to insert the variable 
 
    #Shows the value only of the users.This appears to display the variable correctly in the script block.  
    $VUserLN = get-variable TUUPN -ValueOnly 

    #create the script block to run for the scheduled task.  Add the snappin exchange to enalble the exhchange cmdlets.

    $ScriptBlock = [ScriptBlock]::Create("Remove-ADGroupMember -Identity KBOrgs-O365-Exchange-Online -Members $VUserLN -confirm:`$false`
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
Disable-RemoteMailbox -Identity $VUserLN@yourcorp.com -confirm:`$false")

    $ScriptBlock  | Out-File -FilePath C:\Temp\$VUserLN.ps1

    #Scheduled powershell job rather than task

    $cred = Import-CliXml -Path 'C:\keypath\cred.xml'
    $CurrentTime = Get-Date
    $ThirtyDaysHence = $CurrentTime.date.AddDays(30)
    $TriggerTime = $ThirtyDaysHence.AddHours(9)
    $VTriggerTime = Get-Variable TriggerTime -ValueOnly
    $jo = New-ScheduledJobOption -RunElevated
    $jt = New-JobTrigger -Once -At $TriggerTime
    $jd = Register-ScheduledJob -Name "$TUUPN - Disable Mailbox" -ScheduledJobOption $jo -Trigger $jt -ScriptBlock $scriptblock -Credential $cred

    
    #Send Mail Message to the mail team that the job has been scheduled.
    $SMTPServer = "smtp.office365.com"
    $SMTPPort = "587"
    $credential = Import-CliXml -Path 'C:\keypath\yournamecialcred.xml'
    $From = "Automated <yourname@yourcorp.com>"
    $Subject = "Scheduled job registered for $TUUPN - $TaskNumber"
    $Body = 
"A scheduled registered job was created on: $VTriggerTime
Job Purpose: Remove terminated user from KBOrgs-O365-Exchange-Online and disable mailbox.
Requested by: $Requester
For Termed User: $TUUPN
Tasknumber: $TaskNumber"
    

    Send-MailMessage -From $From -To "distributionlist@yourcorp.com" -Subject $Subject -Body $Body -SmtpServer $SMTPServer -Port $SMTPPort -UseSsl -Credential $Credential

    #Send Mail Message to receiving user and requester.
    $RFSubject = "$TaskNumber - Term ($TUDName) Email Forward"
    $PRUDName = $RUDName.IndexOf(" ")
    $RPRUDFirstName = $RUDName.Substring(0, $PRUDName)
    $BodyReceiver =

"Hi $Requester/$RPRUDFirstName,

Email for ""$TUDName"" is being forwarded to $RUDName.

The requested 'Out of office' reply was added to the terminated user's mailbox: $OOO
    
your name
Messaging Specialist
Your Corp IT
(123) 456-7890
yourname@yourcorp.com"

    
    Send-MailMessage -From $From -To "$RequesterUPN@yourcorp.com", "$RUUPN@yourcorp.com" -Cc "distributionlist@yourcorp.com" -Subject $RFSubject -Body $BodyReceiver -SmtpServer $SMTPServer -Port $SMTPPort -UseSsl -Credential $credential
    

    Disconnect-ExchangeOnline -Confirm:$false
    
    } # end function Reset-UserForwardOOO




