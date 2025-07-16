# Connect to Microsoft Graph
Connect-MgGraph -Scopes "MailboxSettings.ReadWrite", "Mail.ReadWrite"

# Collect user input
$userPrincipalName = Read-Host "Enter the user's email address"
$user = Get-MgUser -UserId $userPrincipalName
$userID = $user.id

if ($null -eq $user) 
    {
        Write-Host "USER NOT FOUND.!" -ForegroundColor Red
        return
    }

#Enter Out of Office time Inputs
$startTime = Read-Host "Enter OOO start date and time (e.g. 2025-07-18 09:00)"
$endTime  = Read-Host "Enter OOO End date and time (e.g. 2025-07-18 20:00)"
$startTimeInIso = (Get-Date $startTime).ToString("yyyy-MM-ddTHH:mm:ss")
$endTimeInIso   = (Get-Date $endTime).ToString("yyyy-MM-ddTHH:mm:ss")

#Enter  Out of Office Message

$InternalMsg = Read-Host "Enter internal reply message"
$externalMsg  = Read-Host "Do you want to set an external reply message? (Yes/No)"

$externalMsg = If ($externalMsg -eq "Yes")

    {
      Read-Host "Enter external reply message"
    }

#Body

$body = @"
{
  `"automaticRepliesSetting`": {
    `"status`": `"scheduled`",
    `"externalAudience`": `"all`",
    `"scheduledStartDateTime`": {
      `"dateTime`": `"$startTimeInIso`",
      `"timeZone`": `"India Standard Time`"
    },
    `"scheduledEndDateTime`": {
      `"dateTime`": `"$endTimeInIso`",
      `"timeZone`": `"India Standard Time`"
    },
    `"internalReplyMessage`": `"$InternalMsg`",
    `"externalReplyMessage`": `"$externalMsg.`"
  }
}
"@

#Setup the OOO
$URI  = "https://graph.microsoft.com/v1.0/users/$userID/mailboxSettings"

Invoke-MgGraphRequest -Uri $URI -Method PATCH -Body $body

Write-Host "Automatic reply settings successfully applied for $userPrincipalName." -ForegroundColor Green

