#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Sends an SMTP message (plain text or HTML) with optional attachments.

.DESCRIPTION
  Sends an email through the System.Net.Mail.SmtpClient directly, without
  relying on the deprecated Send-MailMessage cmdlet (which does not exist in
  PowerShell 7). Works identically under Windows PowerShell 5.1 and PowerShell
  7, needs no elevation, and takes no module dependencies.

  Intended use cases include wiring the JSONL result envelopes produced by the
  other winkit scripts (e.g. Test-ADController, Initialize-ADController) into
  operation reports - see the examples.

  When -Credential is supplied it is used for SMTP authentication (the
  password is never written to any output or log). Without it, the client
  connects anonymously, which is the common case for an on-premises relay
  (e.g. an Exchange server or a domain controller relaying for the local
  network).

  Use -DryRun to render and validate the message without sending anything.

.PARAMETER To
  Primary recipient(s), e.g. 'ops@company.com' or 'IT Ops <ops@company.com>'.

.PARAMETER Cc
  Carbon-copy recipient(s).

.PARAMETER Bcc
  Blind carbon-copy recipient(s).

.PARAMETER From
  Sender address. Defaults to winkit@<computername>.

.PARAMETER Subject
  Message subject.

.PARAMETER Body
  Plain-text message body. Ignored when -BodyHtml is supplied.

.PARAMETER BodyHtml
  HTML message body (sent as an HTML message).

.PARAMETER Attachments
  File path(s) to attach. Every path must exist.

.PARAMETER SmtpServer
  SMTP server to send through. Defaults to localhost, which works when the
  machine can reach the local relay directly.

.PARAMETER Port
  SMTP port. Defaults to 25.

.PARAMETER Ssl
  Use SSL/TLS (SmtpClient EnableSsl) for the connection.

.PARAMETER Credential
  Optional credentials for SMTP authentication, as a PSCredential.

.PARAMETER TimeoutSeconds
  Timeout for the send operation. Defaults to 60 seconds.

.PARAMETER DryRun
  Render and validate the message without sending it.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Send-SMTPMessage.ps1 -To 'ops@company.com' -Subject 'Backup report' -Body 'Backup finished at 02:00.' -SmtpServer exch01.company.com
  Sends a plain-text message through the on-premises Exchange relay.

.EXAMPLE
  PS> ./Send-SMTPMessage.ps1 -To 'ops@company.com' -Subject 'AD health check' -BodyHtml '<h1>AD health</h1><p>See attached report.</p>' -Attachments 'C:\winkit\health.jsonl' -SmtpServer exch01.company.com
  Sends an HTML message with the Test-ADController JSONL envelope attached.

.EXAMPLE
  PS> ./Send-SMTPMessage.ps1 -To 'ops@company.com' -Subject 'Test' -Body 'Dry run.' -DryRun

.LINK
  https://github.com/adnoctem/winkit
  https://nt4admins.de/powershell/e-mail-benachrichtigungen-mit-der-powershell-senden/
  https://learn.microsoft.com/en-us/dotnet/api/system.net.mail.smtpclient
  https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT

  Why not Send-MailMessage? Microsoft deprecated it (it is absent from
  PowerShell 7). This script uses System.Net.Mail.SmtpClient directly, which
  is available in both 5.1 and 7. MailKit is a viable modern alternative when
  advanced SMTP features are needed; it was not chosen here to keep the script
  dependency-free.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [string[]]
  $To,

  [Parameter(Mandatory = $false)]
  [string[]]
  $Cc,

  [Parameter(Mandatory = $false)]
  [string[]]
  $Bcc,

  [Parameter(Mandatory = $false)]
  [string]
  $From = "winkit@$env:COMPUTERNAME",

  [Parameter(Mandatory = $true)]
  [string]
  $Subject,

  [Parameter(Mandatory = $false)]
  [string]
  $Body,

  [Parameter(Mandatory = $false)]
  [string]
  $BodyHtml,

  [Parameter(Mandatory = $false)]
  [string[]]
  $Attachments,

  [Parameter(Mandatory = $false)]
  [string]
  $SmtpServer = 'localhost',

  [Parameter(Mandatory = $false)]
  [ValidateRange(1, 65535)]
  [int]
  $Port = 25,

  [Parameter(Mandatory = $false)]
  [switch]
  $Ssl,

  [Parameter(Mandatory = $false)]
  [pscredential]
  $Credential,

  [Parameter(Mandatory = $false)]
  [ValidateRange(1, 3600)]
  [int]
  $TimeoutSeconds = 60,

  [Parameter(Mandatory = $false)]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force
$ProgressPreference = 'SilentlyContinue'

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no message will be sent`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

# ---- Validation -------------------------------------------------------------

$recipients = @($To) + @($Cc) + @($Bcc) | Where-Object { $_ }
if ($recipients.Count -eq 0) {
  Write-Log -Message 'At least one recipient (-To, -Cc or -Bcc) is required.' -Color Red
  Add-OperationResult -Results $_results -Target 'Recipients' -Source 'SMTP' -Action 'Validate' -Status 'Failed' -Detail 'No recipients specified.'
  exit 1
}

$attachmentPaths = @()
foreach ($attachment in @($Attachments | Where-Object { $_ })) {
  $resolved = $null
  try {
    $resolved = (Get-Item -LiteralPath $attachment -ErrorAction Stop).FullName
  }
  catch {
    $resolved = $null
  }
  if (-not $resolved) {
    Write-Log -Message "Attachment not found: '$attachment'" -Color Red
    Add-OperationResult -Results $_results -Target $attachment -Source 'SMTP' -Action 'Validate' -Status 'Failed' -Detail 'Attachment file not found.'
    if ($PassThru -or $DryRun) { $_results }
    exit 1
  }
  $attachmentPaths += $resolved
}

if (-not $Body -and -not $BodyHtml) {
  Write-Log -Message 'A message body (-Body or -BodyHtml) is required.' -Color Red
  Add-OperationResult -Results $_results -Target 'Body' -Source 'SMTP' -Action 'Validate' -Status 'Failed' -Detail 'No message body specified.'
  exit 1
}

Add-OperationResult -Results $_results -Target 'Recipients' -Source 'SMTP' -Action 'Validate' -Status 'Completed' -Detail ($recipients -join ', ')
if ($attachmentPaths.Count -gt 0) {
  Add-OperationResult -Results $_results -Target 'Attachments' -Source 'SMTP' -Action 'Validate' -Status 'Completed' -Detail ($attachmentPaths -join ', ')
}

Write-Log -Message "Would send '$Subject' to $($recipients -join ', ') via $SmtpServer`:$Port" -Color Cyan

# ---- Send -------------------------------------------------------------------

if (-not $PSCmdlet.ShouldProcess("$($recipients -join ', ')", "Send message via $SmtpServer`:$Port")) {
  Add-OperationResult -Results $_results -Target $Subject -Source 'SMTP' -Action 'Send' -Status 'Skipped' -Detail 'WhatIf - message not sent.'
  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Send-SMTPMessage'
  if ($_operationLog) { Write-Log -Message "Operation log: $_operationLog" -Color Gray }
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

$message = $null
$client = $null
try {
  $message = New-Object System.Net.Mail.MailMessage
  $message.From = [System.Net.Mail.MailAddress]::new($From)
  foreach ($recipient in $To) {
    $message.To.Add($recipient)
  }
  foreach ($recipient in @($Cc)) {
    $message.CC.Add($recipient)
  }
  foreach ($recipient in @($Bcc)) {
    $message.Bcc.Add($recipient)
  }
  $message.Subject = $Subject

  if ($BodyHtml) {
    $message.IsBodyHtml = $true
    $message.Body = $BodyHtml
  }
  else {
    $message.Body = $Body
  }

  foreach ($attachmentPath in $attachmentPaths) {
    $attachment = New-Object System.Net.Mail.Attachment($attachmentPath)
    $message.Attachments.Add($attachment)
  }

  $client = New-Object System.Net.Mail.SmtpClient($SmtpServer, $Port)
  $client.EnableSsl = $Ssl
  $client.Timeout = ($TimeoutSeconds * 1000)
  if ($Credential) {
    $client.Credentials = $Credential.GetNetworkCredential()
  }

  $client.Send($message)
  Add-OperationResult -Results $_results -Target $Subject -Source 'SMTP' -Action 'Send' -Status 'Completed' -Detail "Sent to $($recipients -join ', ') via $SmtpServer`:$Port."
}
catch {
  Add-OperationResult -Results $_results -Target $Subject -Source 'SMTP' -Action 'Send' -Status 'Failed' -Detail $_.Exception.Message
}
finally {
  if ($message) {
    foreach ($attachment in @($message.Attachments)) {
      $attachment.Dispose()
    }
    $message.Dispose()
  }
  if ($client) {
    $client.Dispose()
  }
}

# ---- Summary ----------------------------------------------------------------

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Send-SMTPMessage'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

$failed = @($_results | Where-Object { $_.Status -eq 'Failed' })
if ($PassThru -or $DryRun) {
  $_results
}

if ($failed.Count -gt 0) {
  Write-Log -Message 'Message could not be sent - see the results above.' -Color Red
  exit 1
}

Write-Log -Message 'Message sent successfully.' -Color Green
exit 0
