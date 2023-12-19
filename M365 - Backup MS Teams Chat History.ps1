#########
# au2mator PS Services
# New Service
# M365 - Backup MS Teams Chat History
# v 1.0 Initial Release
# Init Release: 13.08.2020
# Last Update: 13.08.2020
# Code Template V 1.1
# URL: 
# Github: 
# au2mator 4.x or above
#
# Special Thanks to Christian Schreiber for the Idea: https://techcommunity.microsoft.com/t5/windows-powershell/script-for-teams-chat-backup/m-p/1547371#M1519
#
##################


#region InputParamaters
##Question in au2mator
param (
    [parameter(Mandatory = $true)]
    [String]$c_TeamsTeam,

    [parameter(Mandatory = $false)]
    [String]$c_AddRecipients,


    ## au2mator Initialize Data
    [parameter(Mandatory = $true)]
    [String]$InitiatedBy,

    [parameter(Mandatory = $true)]
    [String]$RequestId,

    [parameter(Mandatory = $true)]
    [String]$Service,

    [parameter(Mandatory = $true)]
    [String]$TargetUserId
)
#endregion  InputParamaters



#region Variables
##Script Handling
$DoImportPSSession = $false
$ErrorCount = 0
#$ErrorActionPreference = "SilentlyContinue"

## Environment
[string]$DCServer = 'svdc01'
[string]$LogPath = "C:\_SCOworkingDir\TFS\PS-Services\M365 - Backup MS Teams Chat History"
[string]$LogfileName = "Backup MS Teams Chat History"
[string]$CredentialStorePath = "C:\_SCOworkingDir\TFS\PS-Services\CredentialStore" #see for details: https://click.au2mator.com/PSCreds/
$Modules = @("ActiveDirectory", "SharePointPnPPowerShellOnline")


## au2mator Settings
[string]$PortalURL = "http://demo01.au2mator.local"
[string]$au2matorDBServer = "demo01"
[string]$au2matorDBName = "au2matorNew"

## Control Mail
$SendMailToInitiatedByUser = $true #Send a Mail after Service is completed
$SendMailToTargetUser = $true #Send Mail to Target User after Service is completed

## SMTP Settings
$SMTPServer = "smtp.office365.com"
$SMPTAuthentication = $true #When True, User and Password needed
$EnableSSLforSMTP = $true
$SMTPSender = "michael.seidl@au2mator.com"
$SMTPPort="587"

$SMTPCredential_method = "Stored" #Stored, Manual
#Use stored Credentials
$SMTPcredential_File = "SMTPCreds.xml"
#Use Manual Credentials
$SMTPUser = ""
$SMTPPassword = ""


if ($SMTPCredential_method -eq "Stored") {
    $SMTPcredential = Import-CliXml -Path (Get-ChildItem -Path $CredentialStorePath -Filter $SMTPcredential_File).FullName
}

if ($SMTPCredential_method -eq "Manual") {
    $f_secpasswd = ConvertTo-SecureString $SMTPPassword -AsPlainText -Force
    $SMTPcredential = New-Object System.Management.Automation.PSCredential ($SMTPUser, $f_secpasswd)
}

# Additional Settings

#PowerShell Online Credentials
$PSOnlineCredential_method = "Stored" #Stored, Manual
#Use stored Credentials
$PSOnlineCredential_File = "PSOnlineCreds.xml"
#Use Manual Credentials
$PSOnlineUser = ""
$PSOnlinePassword = ""


if ($PSOnlineCredential_method -eq "Stored") {
    $PSOnlineCredential = Import-CliXml -Path (Get-ChildItem -Path $CredentialStorePath -Filter $PSOnlineCredential_file).FullName
}

if ($PSOnlineCredential_method -eq "Manual") {
    $f_secpasswd = ConvertTo-SecureString $PSOnlinePassword -AsPlainText -Force
    $PSOnlineCredential = New-Object System.Management.Automation.PSCredential ($PSOnlineUser, $f_secpasswd)
}

$TempBackupStorage = "C:\_SCOworkingDir\TFS\PS-Services\M365 - Backup MS Teams Chat History" #Path to store Export File



#endregion Variables

#region Functions

function ConnectToDB {
    # define parameters
    param(
        [string]
        $servername,
        [string]
        $database
    )
    # create connection and save it as global variable
    $global:Connection = New-Object System.Data.SQLClient.SQLConnection
    $Connection.ConnectionString = "server='$servername';database='$database';trusted_connection=false; integrated security='true'"
    $Connection.Open()
    Write-Verbose 'Connection established'
}

function ExecuteSqlQuery {
    # define parameters
    param(

        [string]
        $sqlquery

    )

    Begin {
        If (!$Connection) {
            Throw "No connection to the database detected. Run command ConnectToDB first."
        }
        elseif ($Connection.State -eq 'Closed') {
            Write-Verbose 'Connection to the database is closed. Re-opening connection...'
            try {
                # if connection was closed (by an error in the previous script) then try reopen it for this query
                $Connection.Open()
            }
            catch {
                Write-Verbose "Error re-opening connection. Removing connection variable."
                Remove-Variable -Scope Global -Name Connection
                throw "Unable to re-open connection to the database. Please reconnect using the ConnectToDB commandlet. Error is $($_.exception)."
            }
        }
    }

    Process {
        #$Command = New-Object System.Data.SQLClient.SQLCommand
        $command = $Connection.CreateCommand()
        $command.CommandText = $sqlquery

        Write-Verbose "Running SQL query '$sqlquery'"
        try {
            $result = $command.ExecuteReader()
        }
        catch {
            $Connection.Close()
        }
        $Datatable = New-Object "System.Data.Datatable"
        $Datatable.Load($result)
        return $Datatable
    }
    End {
        Write-Verbose "Finished running SQL query."
    }
}

function Write-au2matorLog {
    [CmdletBinding()]
    param
    (
        [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR')]
        [string]$Type,
        [string]$Text
    )

    # Set logging path
    if (!(Test-Path -Path $logPath)) {
        try {
            $null = New-Item -Path $logPath -ItemType Directory
            Write-Verbose ("Path: ""{0}"" was created." -f $logPath)
        }
        catch {
            Write-Verbose ("Path: ""{0}"" couldn't be created." -f $logPath)
        }
    }
    else {
        Write-Verbose ("Path: ""{0}"" already exists." -f $logPath)
    }
    [string]$logFile = '{0}\{1}_{2}.log' -f $logPath, $(Get-Date -Format 'yyyyMMdd'), $LogfileName
    $logEntry = '{0}: <{1}> <{2}> <{3}> {4}' -f $(Get-Date -Format dd.MM.yyyy-HH:mm:ss), $Type, $RequestId, $Service, $Text
    Add-Content -Path $logFile -Value $logEntry
}

function Get-UserInput ($RequestID) {
    [hashtable]$return = @{ }

    ConnectToDB -servername $au2matorDBServer -database $au2matorDBName

    $Result = ExecuteSqlQuery -sqlquery "SELECT        RPM.Text AS Question, RP.Value
    FROM            dbo.Requests AS R INNER JOIN
                             dbo.RunbookParameterMappings AS RPM ON R.ServiceId = RPM.ServiceId INNER JOIN
                             dbo.RequestParameters AS RP ON RPM.ParameterName = RP.[Key] AND R.RequestId = RP.RequestId
    where RP.RequestId = '$RequestID' order by [Order]"

    $html = "<table><tr><td><b>Question</b></td><td><b>Answer</b></td></tr>"
    $html = "<table>"
    foreach ($row in $Result) {
        $row
        $html += "<tr><td><b>" + $row.Question + "</b></td><td>" + $row.Value + "</td></tr>"
    }
    $html += "</table>"

    $f_RequestInfo = ExecuteSqlQuery -sqlquery "select InitiatedBy, TargetUserId,[ApprovedBy], [ApprovedTime], Comment from Requests where RequestId =  '$RequestID'"

    $Connection.Close()
    Remove-Variable -Scope Global -Name Connection

    $f_SamInitiatedBy = $f_RequestInfo.InitiatedBy.Split("\")[1]
    $f_UserInitiatedBy = Get-ADUser -Identity $f_SamInitiatedBy -Properties Mail


    $f_SamTarget = $f_RequestInfo.TargetUserId.Split("\")[1]
    $f_UserTarget = Get-ADUser -Identity $f_SamTarget -Properties Mail

    $return.InitiatedBy = $f_RequestInfo.InitiatedBy
    $return.MailInitiatedBy = $f_UserInitiatedBy.mail
    $return.MailTarget = $f_UserTarget.mail
    $return.TargetUserId = $f_RequestInfo.TargetUserId
    $return.ApprovedBy = $f_RequestInfo.ApprovedBy
    $return.ApprovedTime = $f_RequestInfo.ApprovedTime
    $return.Comment = $f_RequestInfo.Comment
    $return.HTML = $HTML

    return $return
}

Function Get-MailContent ($RequestID, $RequestTitle, $EndDate, $TargetUserId, $InitiatedBy, $Status, $PortalURL, $RequestedBy, $AdditionalHTML, $InputHTML) {

    $f_RequestID = $RequestID
    $f_InitiatedBy = $InitiatedBy

    $f_RequestTitle = $RequestTitle
    $f_EndDate = $EndDate
    $f_RequestStatus = $Status
    $f_RequestLink = "$PortalURL/requeststatus?id=$RequestID"
    $f_RequestedBy = $RequestedBy
    $f_HTMLINFO = $AdditionalHTML
    $f_InputHTML = $InputHTML

    $f_SamInitiatedBy = $f_InitiatedBy.Split("\")[1]
    $f_UserInitiatedBy = Get-ADUser -Identity $f_SamInitiatedBy -Properties DisplayName
    $f_DisplaynameInitiatedBy = $f_UserInitiatedBy.DisplayName


    $HTML = @'
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 1.5pt; background: #F7F8F3; mso-yfti-tbllook: 1184;" border="0" width="100%" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes;">
    <td style="padding: .75pt .75pt .75pt .75pt;" valign="top">&nbsp;</td>
    <td style="width: 450.0pt; padding: .75pt .75pt .75pt .75pt; box-sizing: border-box;" valign="top" width="600">
    <div style="box-sizing: border-box;">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: white; border: solid #E9E9E9 1.0pt; mso-border-alt: solid #E9E9E9 .75pt; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="1" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes;">
    <td style="border: none; background: #6ddc36; padding: 15.0pt 0cm 15.0pt 15.0pt;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><img src="https://au2mator.com/wp-content/uploads/2018/02/HPLogoau2mator-1.png" alt="" width="198" height="43" /></p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 1; box-sizing: border-box;">
    <td style="border: none; padding: 15.0pt 15.0pt 15.0pt 15.0pt; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="width: 55.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top" width="55%">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes;">
    <td style="width: 18.75pt; border-top: solid #E3E3E3 1.0pt; border-left: solid #E3E3E3 1.0pt; border-bottom: none; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" width="25">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: none; border-bottom: none; border-right: solid #E3E3E3 1.0pt; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm; font-color: #0000;"><strong>End Date</strong>: ##EndDate</td>
    </tr>
    <tr style="mso-yfti-irow: 1;">
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: solid #E3E3E3 1.0pt; border-bottom: none; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 0cm 0cm;">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: none; border-bottom: none; border-right: solid #E3E3E3 1.0pt; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;"><strong>Status</strong>: ##Status</td>
    </tr>
    <tr style="mso-yfti-irow: 2; mso-yfti-lastrow: yes;">
    <td style="border: solid #E3E3E3 1.0pt; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; mso-border-bottom-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border: solid #E3E3E3 1.0pt; border-left: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-bottom-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;"><strong>Requested By</strong>: ##RequestedBy</td>
    </tr>
    </tbody>
    </table>
    </td>
    <td style="width: 5.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" width="5%">
    <p class="MsoNormal" style="line-height: 19.2pt;"><span style="font-size: 9.0pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
    </td>
    <td style="width: 40.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top" width="40%">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: #FAFAFA; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes;">
    <td style="width: 100.0%; border: solid #E3E3E3 1.0pt; mso-border-alt: solid #E3E3E3 .75pt; padding: 7.5pt 0cm 1.5pt 3.75pt;" width="100%">
    <p style="text-align: center;" align="center"><span style="font-size: 10.5pt; color: #959595;">Request ID</span></p>
    <p class="MsoNormal" style="text-align: center;" align="center">&nbsp;</p>
    <p style="text-align: center;" align="center"><u><span style="font-size: 12.0pt; color: black;"><a href="##RequestLink"><span style="color: black;">##REQUESTID</span></a></span></u></p>
    <p class="MsoNormal" style="text-align: center;" align="center"><span style="mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 2; box-sizing: border-box;">
    <td style="border: none; padding: 0cm 15.0pt 15.0pt 15.0pt; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><strong><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">Dear ##UserDisplayname,</span></strong></p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 1; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">the Request <strong>"##RequestTitle"</strong> has been finished.<br /> <br /> Please see the description for detailed information.<br /><b>##HTMLINFO&nbsp;</b><br /></span></p>
    <div>&nbsp;</div>
    <div>See the Details of the Request</div>
    <div>##InputHTML</div>
    <div>&nbsp;</div>
    <div>&nbsp;</div>
    Kind regards,<br /> au2mator Self Service Team
    <p>&nbsp;</p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 2; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center"><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';"><a style="border-radius: 3px; -webkit-border-radius: 3px; -moz-border-radius: 3px; display: inline-block;" href="##RequestLink"><strong><span style="color: white; border: solid #50D691 6.0pt; padding: 0cm; background: #50D691; text-decoration: none; text-underline: none;">View your Request</span></strong></a></span></p>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 3; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="border: none; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: #333333; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="width: 50.0%; border: none; border-right: solid lightgrey 1.0pt; mso-border-right-alt: solid lightgrey .75pt; padding: 22.5pt 15.0pt 22.5pt 15.0pt; box-sizing: border-box;" valign="top" width="50%">&nbsp;</td>
    <td style="width: 50.0%; padding: 22.5pt 15.0pt 22.5pt 15.0pt; box-sizing: border-box;" valign="top" width="50%">&nbsp;</td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </div>
    </td>
    <td style="padding: .75pt .75pt .75pt .75pt; box-sizing: border-box;" valign="top">&nbsp;</td>
    </tr>
    </tbody>
    </table>
    <p class="MsoNormal"><span style="mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
'@

    $html = $html.replace('##REQUESTID', $f_RequestID).replace('##UserDisplayname', $f_DisplaynameInitiatedBy).replace('##RequestTitle', $f_RequestTitle).replace('##EndDate', $f_EndDate).replace('##Status', $f_RequestStatus).replace('##RequestedBy', $f_RequestedBy).replace('##HTMLINFO', $f_HTMLINFO).replace('##InputHTML', $f_InputHTML).replace('##RequestLink', $f_RequestLink)

    return $html
}

function Send-ServiceMail ($HTMLBody, $ServiceName, $Recipient, $RequestID, $RequestStatus, $Attachment) {
    $f_Subject = "au2mator - $ServiceName Request [$RequestID] - $RequestStatus"

    if ($SMPTAuthentication) {
        if ($EnableSSLforSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Credential $SMTPcredential -UseSsl -Attachments $Attachment -Port $SMTPPort
        }
        else {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Credential $SMTPcredential -Attachments $Attachment -Port $SMTPPort
        }
    }
    else {
        if ($EnableSSLforSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -UseSsl -Attachments $Attachment -Port $SMTPPort
        }
        else {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Attachments $Attachment -Port $SMTPPort
        }
    }
}
#endregion Functions


#region Script
Write-au2matorLog -Type INFO -Text "Start Script"
if ($DoImportPSSession) {

    Write-au2matorLog -Type INFO -Text "Import-Pssession"
    $PSSession = New-PSSession -ComputerName $DCServer
    Import-PSSession -Session $PSSession -DisableNameChecking -AllowClobber
}
else {

}

#Check for Modules if installed
Write-au2matorLog -Type INFO -Text "Try to install all PowerShell Modules"
foreach ($Module in $Modules) {
    if (Get-Module -ListAvailable -Name $Module) {
        Write-au2matorLog -Type INFO -Text "Module is already installed:  $Module"
    }
    else {
        Write-au2matorLog -Type INFO -Text "Module is not installed, try simple method:  $Module"
        try {
            
            Install-Module $Module -Force -Confirm:$false 
            Write-au2matorLog -Type INFO -Text "Module was installed the simple way:  $Module"
        }
        catch {
            Write-au2matorLog -Type INFO -Text "Module is not installed, try the advanced way:  $Module"
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
                Install-PackageProvider -Name NuGet  -MinimumVersion 2.8.5.201 -Force
                Install-Module $Module -Force -Confirm:$false 
                Write-au2matorLog -Type INFO -Text "Module was installed the advanced way:  $Module"
            }
            catch {
                Write-au2matorLog -Type ERROR -Text "could not install module:  $Module"
                $au2matorReturn = "could not install module:  $Module, Error: $Error"
                $Status = "ERROR"
                $ErrorCount = 1
            }
        }
    }
}


Write-au2matorLog -Type INFO -Text "Import all PowerShell Modules"
foreach ($Module in $Modules) {
    Write-au2matorLog -Type INFO -Text "Import Module:  $Module"
    Import-Module -name $Module
}


try {
    $SecurityScope = @("Group.Read.All")
    Connect-PnPOnline -Scopes $SecurityScope -Credentials $PSOnlineCredential 
    $PnPGraphAccessToken = Get-PnPGraphAccessToken


    $Team_ID = $c_TeamsTeam
    $Headers = @{
        "Content-Type" = "application/json"
        Authorization  = "Bearer $PnPGraphAccessToken"    
    }        
    $Date = Get-Date -Format "dd.MM.yyyy, HH:mm"
    $DOCTYPE = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Strict//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd'><html xmlns='http://www.w3.org/1999/xhtml'>"
    $Style = "<style>table {border-collapse: collapse; width:100%;} table th {text-align:left; background-color: #004C99; color:#fff; padding: 4px 30px 4px 8px;} table td {border: 1px solid #004C99; padding: 4px 8px;} td {background-color: #DDE5FF}</style>"     
    $Head = "<head><title>Backup: Teams-Chat</title></head>"
    $Body = "<body><div style='width: 100%;'><table><tr><th style='text-align:center'><h1>Backup: Teams-Chat from $Date</h1></th></tr></table>"               
    $Table_body = "<div style='width: 100%;'><table><tr><th>TimeStamp</th><th>User Name</th><th>Message</th></tr>"
    $Content = ""
    $Footer = "</body>"
    Start-Sleep -Milliseconds 50
        
    $response_team = Invoke-RestMethod -Uri  "https://graph.microsoft.com/beta/teams/$Team_ID" -Method Get -Headers $Headers -UseBasicParsing
    $Content += "</br></br><hr><h2>Team: " + $response_team.displayName + "</h2>"
    $response_channels = Invoke-RestMethod -Uri  "https://graph.microsoft.com/beta/teams/$Team_ID/channels" -Method Get -Headers $Headers -UseBasicParsing
    $response_channels.value | Select-Object -Property ID, displayName | ForEach-Object {
        $Channel_ID = $_.ID
        $Channel_displayName = $_.displayName
            
        Start-Sleep -Milliseconds 50    
        $Content += "<h3>Channel: " + $Channel_displayName + "</h3>"
        $response_messages = Invoke-RestMethod -Uri  "https://graph.microsoft.com/beta/teams/$Team_ID/channels/$Channel_ID/messages" -Method Get -Headers $Headers -UseBasicParsing
        $response_messages.value | Select-Object -Property ID, createdDateTime, from | ForEach-Object {
            $Message_ID = $_.ID
            $Message_TimeStamp = $_.createdDateTime
            $Message_from = $_.from     
            try {
                $response_content = Invoke-RestMethod -Uri  "https://graph.microsoft.com/beta/teams/$Team_ID/channels/$Channel_ID/messages/$Message_ID" -Method Get -Headers $Headers -UseBasicParsing
                
                Start-Sleep -Milliseconds 50                                                         
                $Content += $Table_body + "<td>" + $Message_TimeStamp + "</td><td style='width: 10%;'>" + $Message_from.user.displayName + "</td><td style='width: 75%;'>" + $response_content.body.content + $response_content.attachments.id + "</td></table></div>"
                $response_Reply = Invoke-RestMethod -Uri  "https://graph.microsoft.com/beta/teams/$Team_ID/channels/$Channel_ID/messages/$Message_ID/replies" -Method Get -Headers $Headers -UseBasicParsing
                $response_Reply.value | Select-Object -Property ID, createdDateTime, from | ForEach-Object {
                    $Reply_ID = $_.ID
                    $Reply_TimeStamp = $_.createdDateTime
                    $Reply_from = $_.from                        
                    $response_Reply = Invoke-RestMethod -Uri  "https://graph.microsoft.com/beta/teams/$Team_ID/channels/$Channel_ID/messages/$Message_ID/replies/$Reply_ID" -Method Get -Headers $Headers -UseBasicParsing
                    Start-Sleep -Milliseconds 50
                    ForEach-Object {                                                                          
                        $Content += $Table_body + "<td>" + $Reply_TimeStamp + "</td><td style='width: 10%;'>" + $Reply_from.user.displayName + "</td><td style='width: 75%;'>" + $response_Reply.body.content + $response_Reply.attachments.id + $response_Reply.attachments.name + "</td></table></div>"
                    }
                } 
            }
            catch {
                
            }               
            
        }                                
    }
    
    $DOCTYPE + $Style + $Head + $Body + $Content + $Footer |  Out-File -FilePath "$TempBackupStorage\Backup.html"
  
    try {
        if ($c_AddRecipients -ne "") {
            Write-au2matorLog -Type INFO -Text "Additional Recipients selected"
            if ($c_AddRecipients -like '*;*') {$Split=$c_AddRecipients}else{$Split=$c_AddRecipients.split(";")}
            foreach ($R in $Split) {
                Write-au2matorLog -Type INFO -Text "Get Mail from $R"
                $User=Get-ADUser -Identity $R -Properties  mail
                Write-au2matorLog -Type INFO -Text "Mail is $($User.mail)"
                $UserInput = Get-UserInput -RequestID $RequestId
                $HTML = Get-MailContent -RequestID $RequestId -RequestTitle $Service -EndDate $UserInput.ApprovedTime -TargetUserId $TargetUserId -InitiatedBy $InitiatedBy -Status $Status -PortalURL $PortalURL -RequestedBy $InitiatedBy -AdditionalHTML $AdditionalHTML -InputHTML $UserInput.html
                Send-ServiceMail -HTMLBody $HTML -RequestID $RequestId -Recipient "$($User.mail)" -RequestStatus $Status -ServiceName $Service -Attachment "$TempBackupStorage\Backup.html"
            
            }
        }
        else {
            Write-au2matorLog -Type INFO -Text "No Additional Recipients selected, go ahead"
        }
    }
    catch {
        $ErrorCount = 1
        Write-au2matorLog -Type ERROR -Text "failed to send Mail with Chat Backup"
        Write-au2matorLog -Type ERROR -Text $Error
    }

}
catch {
    $ErrorCount = 1
    Write-au2matorLog -Type ERROR -Text "Failed to Backup Chat History"
    Write-au2matorLog -Type ERROR -Text $Error
}





    if ($ErrorCount -eq 0) {
        $au2matorReturn = "Teams Backup successfull"
        $AdditionalHTML="<br>
        Teams Backup successfull 
        <br>
        "
        $Status = "COMPLETED"
    }
    else {
        $au2matorReturn = "Failed to Backup Chat History, Error: $Error"
        $Status = "ERROR"
    }

#endregion Script

#region Return
## return to au2mator Services



Write-au2matorLog -Type INFO -Text "Service finished"



if ($SendMailToInitiatedByUser) {
    Write-au2matorLog -Type INFO -Text "Send Mail to Initiated By User"

    $UserInput = Get-UserInput -RequestID $RequestId
    $HTML = Get-MailContent -RequestID $RequestId -RequestTitle $Service -EndDate $UserInput.ApprovedTime -TargetUserId $TargetUserId -InitiatedBy $InitiatedBy -Status $Status -PortalURL $PortalURL -RequestedBy $InitiatedBy -AdditionalHTML $AdditionalHTML -InputHTML $UserInput.html
    Send-ServiceMail -HTMLBody $HTML -RequestID $RequestId -Recipient "$($UserInput.MailInitiatedBy)" -RequestStatus $Status -ServiceName $Service -Attachment "$TempBackupStorage\Backup.html"
}


if ($SendMailToTargetUser) {
    Write-au2matorLog -Type INFO -Text "Send Mail to Target User"

    $UserInput = Get-UserInput -RequestID $RequestId
    $HTML = Get-MailContent -RequestID $RequestId -RequestTitle $Service -EndDate $UserInput.ApprovedTime -TargetUserId $TargetUserId -InitiatedBy $InitiatedBy -Status $Status -PortalURL $PortalURL -RequestedBy $InitiatedBy -AdditionalHTML $AdditionalHTML -InputHTML $UserInput.html
    Send-ServiceMail -HTMLBody $HTML -RequestID $RequestId -Recipient "$($UserInput.MailTarget)" -RequestStatus $Status -ServiceName $Service -Attachment "$TempBackupStorage\Backup.html"
}


return $au2matorReturn
#endregion Return


