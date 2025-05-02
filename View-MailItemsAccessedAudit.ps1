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

# Connect to Exchange Online
Connect-ExchangeOnline

# Display a message to indicate the script's purpose
Write-Host "Searching Office 365 Audit Records to find MailItemAccessed (Last 90 Days)" -ForegroundColor Green

# Suppress non-terminating errors to keep output clean
$ErrorActionPreference = 'SilentlyContinue'

# Define the date range: from 90 days ago until today
$EndDate = Get-Date
$StartDate = $EndDate.AddDays(-90)

# Perform the audit log search for the "MailItemsAccessed" operation
# Filtering results to only those related to "admin@hybrid-pro.net"
$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations MailItemsAccessed -Formatted -FreeText admin@hybrid-pro.net

# Check if records were returned
if (-not $Records) {
    Write-Host "No MailItemsAccessed records found for the specified period." -ForegroundColor Yellow
    return
}

# Process each record and extract relevant properties
$ProcessedRecords = $Records | ForEach-Object {
    $data = $_.AuditData | ConvertFrom-Json
    [PSCustomObject]@{
        Time              = $_.CreationDate                          # Timestamp of the access
        UserId            = $data.UserId                             # User who accessed the item
        ClientIP          = $data.ClientIPAddress                    # IP address used
        FolderPath        = $data.Folders[0].Path                    # Folder path where the message resides
        InternetMessageId = $data.Folders[0].FolderItems[0].InternetMessageId  # Unique message identifier
    }
}

# Display the results in an interactive grid window
$ProcessedRecords | Out-GridView -Title "MailItemsAccessed Events - Last 90 Days"
$ProcessedRecords | select InternetMessageId
