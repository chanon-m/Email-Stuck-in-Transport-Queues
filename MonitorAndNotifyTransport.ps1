Import-module activedirectory
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

#--- Email Notification Function ---------
Function Send-Mail([string]$Body,[string]$MessageSubject) {
  $FromAddress = 'sender@email.com'
  $ToAddress = 'recipient01@email.com,recipient02@email.com'
  $SendingServer = 'smtp.email.com'
  $SMTPMessage = New-Object System.Net.Mail.MailMessage $FromAddress, $ToAddress, $MessageSubject, $Body -ErrorAction SilentlyContinue
  $SMTPMessage.IsBodyHTML = $true
  $SMTPMessage.Priority = [System.Net.Mail.MailPriority]::High
  $SMTPClient = New-Object System.Net.Mail.SMTPClient $SendingServer -ErrorAction SilentlyContinue

  If(Test-Connection -Cn $SendingServer -BufferSize 16 -Count 1 -ea 0 -quiet) {
     $SMTPClient.Send($SMTPMessage)
  } else {
     Write-Host "Cannot connect to SMTP server!" 
  }

#Check Event ID 4004 in Application in Exchange 2016 Server

$queueInfo = Get-Queue
  $Filter = @{
     LogName = 'Application'
     Id = 4004
     StartTime = (Get-Date).AddHours(-2)
  }  

# Define Transport queue
$count = 100

# Monitor Event and Transport queue
Foreach($queue in $queueInfo) {
   #check if the delivery type is undefined
  If($queue.DeliveryType -eq 'Undefined') {
     #Display the queue details with undefined DeiveryType
     Write-Host 'Submission Message Count\EventID: $($queue.MessageCount)'
     If($queue.MessageCount -gt $count) {
        $event = Get-WinEvent -FilterHashtable $Filter -MaxEvents 1 -ErrorAction SilentlyContinue
        Write-Host 'EventID: $($event.id)'
        if($event) {
           Write-Host 'Restart MSExchange Transport Service'
           Restart-Service MSExchangeTransport
           #Email Notication
           $ServerName = HostName
           $Subject = '$ServerName Restart MSExchangeTransport'
           $Service = Get-Service MSExchangeTransport
           $mailbody = 'Submission Message Count: $($queue.MessageCount)'
           $mailbody += '<p>' + $Service.Name + ':' + $Service.Status  
           Send-Mail $mailbody $Subject
        }
     }
  }
}
