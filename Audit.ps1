# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

# Create an empty array to store results
$emailData = @()

# Get all shared mailboxes in the tenant
$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox

# Loop through each shared mailbox
foreach ($mailbox in $sharedMailboxes) {
    $sharedMailbox = $mailbox.PrimarySmtpAddress

    # Get the last received email from the shared mailbox using Get-MessageTrace
    $lastReceivedEmail = Get-MessageTrace -RecipientAddress $sharedMailbox | 
                         Sort-Object Received -Descending | 
                         Select-Object -First 1

    # Get the last sent email from the shared mailbox using Get-MessageTrace
    $lastSentEmail = Get-MessageTrace -SenderAddress $sharedMailbox | 
                     Sort-Object Received -Descending | 
                     Select-Object -First 1

    # Add data for each mailbox
    $emailData += [pscustomobject]@{
        Mailbox          = $sharedMailbox
        LastReceivedTime = if ($lastReceivedEmail) { $lastReceivedEmail.Received } else { "No Data" }
        LastSentTime     = if ($lastSentEmail) { $lastSentEmail.Received } else { "No Data" }
    }
}

# Export the data to a CSV file
$emailData | Export-Csv -Path "C:\SharedMailboxesLastEmails.csv" -NoTypeInformation

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
