<#
    Copyright (c) 2025 Abdallah Salah
    All Rights Reserved.

    This script is part of the Microsoft365-Scripts repository on GitHub.
    Repository URL: https://github.com/abdallahsalah-dev/M365-Scripts

    Licensed under the MIT License.
    You may obtain a copy of the License at:
    https://opensource.org/licenses/MIT

    Author: Abdallah Salah
    Description: This script searches Office 365 Audit Logs for "MailItemsAccessed" events.
    It extracts relevant details such as the timestamp, user ID, client IP, folder path,
    and Internet Message ID, and displays the results in an interactive grid.
#>

# Connect to Exchange Online (prompts for authentication if not connected)
Connect-ExchangeOnline

# Display a message to indicate the script's purpose
Write-Host "Searching Office 365 Audit Records to find MailItemsAccessed (Last 90 Days)" -ForegroundColor Green

# Suppress non-terminating errors to keep the output clean
$ErrorActionPreference = 'SilentlyContinue'

# Define the date range: from 90 days ago until today
$EndDate = Get-Date
$StartDate = $EndDate.AddDays(-90)

# Define the attacker IP address to use as filter for the audit log search
$AttackerIP = "1.2.3.4"  # <--- Replace this with the actual attacker IP address

# Search Unified Audit Log for MailItemsAccessed events filtered by the attacker IP
$Records = Search-UnifiedAuditLog `
    -StartDate $StartDate `
    -EndDate $EndDate `
    -Operations MailItemsAccessed `
    -Formatted `
    -FreeText $AttackerIP

# Check if any records were found
if (-not $Records) {
    Write-Host "No MailItemsAccessed records found for the specified IP address in the last 90 days." -ForegroundColor Yellow
    return
}

# Process each audit record and extract key properties into a custom object
$ProcessedRecords = $Records | ForEach-Object {
    # Convert the JSON string in AuditData property to a PowerShell object
    $data = $_.AuditData | ConvertFrom-Json
    
    # Create a custom object with relevant extracted information
    [PSCustomObject]@{
        Time              = $_.CreationDate                          # When the message was accessed
        UserId            = $data.UserId                             # User who accessed the message
        ClientIP          = $data.ClientIPAddress                    # IP address used to access the message
        FolderPath        = $data.Folders[0].Path                    # Folder location of the message
        InternetMessageId = $data.Folders[0].FolderItems[0].InternetMessageId  # Unique email message ID
    }
}

# Display the results in an interactive window for easy review and filtering
$ProcessedRecords | Out-GridView -Title "MailItemsAccessed Events - Last 90 Days"

# Output the list of InternetMessageId values to the console for reference or further use
$ProcessedRecords | Select-Object InternetMessageId
