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

# Process each record and extract relevant properties
$Records | ForEach-Object {
    $data = $_.AuditData | ConvertFrom-Json
    [PSCustomObject]@{
        Time              = $_.CreationDate                        # Timestamp of the access
        UserId            = $data.UserId                          # User who accessed the item
        ClientIP          = $data.ClientIPAddress                 # IP address used
        FolderPath        = $data.Folders[0].Path                 # Folder path where the message resides
        InternetMessageId = $data.Folders[0].FolderItems[0].InternetMessageId  # Unique message identifier
    }
} | Out-GridView  # Display the results in an interactive grid window
