<#
    Copyright (c) 2025 Abdallah Salah
    All Rights Reserved.

    This script is part of the Microsoft365-Scripts repository on GitHub.
    Repository URL: https://github.com/abdallahsalah-dev/M365-Scripts

    Licensed under the MIT License.
    You may obtain a copy of the License at:
    https://opensource.org/licenses/MIT

    Author: Abdallah Salah
    
    Description: This PowerShell script uses Microsoft Graph API to remove simulation-related emails and reminders from users' mailboxes. It targets messages from a specific sender (notification@attacksimulationtraining.com) and 
    deletes them across the organization or specific users. The script requires the Mail.ReadWrite permission and is intended as a workaround for handling canceled simulations and their associated notifications..

    Script Hints:
    Target Specific Users: The script can be customized to target specific users by modifying the $users array. For example, instead of processing all users, you can specify particular email addresses to remove simulation emails from.
    
    Example:

 $users = @(
    "user1@example.com",
    "user2@example.com"
           )
#>

# Define sender email
$senderEmail = "MSSecurity-noreply@microsoft.com"

# Initialize an ArrayList (better than using += with arrays)
$allMessages = [System.Collections.ArrayList]::new()

# Get all users' UserPrincipalNames
$users = (Get-MgUser -All).UserPrincipalName

# Loop through each user and get their messages
foreach ($userUPN in $users) {
    # Get user messages with the specified sender email
    $userMessages = Get-MgUserMessage -All -UserId $userUPN -Filter "from/emailAddress/address eq '$senderEmail'" | 
        Select-Object subject, ReceivedDateTime, Id, @{Name='UserUPN'; Expression = { $userUPN }}

    # Only add non-null, non-empty messages to the ArrayList
    if ($userMessages -ne $null -and $userMessages.Count -gt 0) {
        $allMessages.AddRange($userMessages)
    }
    else {
        Write-Host "No messages found for user: $userUPN"
    }
}

# Check if any messages were found and output the result
if ($allMessages.Count -gt 0) {
    # Print all retrieved messages
    $allMessages | Select-Object Subject, ReceivedDateTime, UserUPN | Out-GridView
}
else {
    Write-Host -ForegroundColor Red "No messages found for the specified sender."
    break
}

# Initialize the deletion count
$deletedCount = 0

# Optional: Uncomment the following block to enable deletion of messages


# Loop through all messages to delete them
foreach ($message in $allMessages) {
    # Attempt to delete the message
    Remove-MgUserMessage -UserId $message.UserUPN -MessageId $message.Id -ErrorAction SilentlyContinue
    
    # Increment the deleted count
    $deletedCount++
}


# Output the final count of found and deleted messages
write-Host -ForegroundColor green  "Total found messages: $($allMessages.Count)"
write-Host -ForegroundColor Red "Total deleted messages: $deletedCount"
