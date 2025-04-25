<#
    Copyright (c) 2025 Abdallah Salah
    All Rights Reserved.

    This script is part of the Microsoft365-Scripts repository on GitHub.
    Repository URL: https://github.com/abdallahsalah-dev/M365-Scripts

    Licensed under the MIT License.
    You may obtain a copy of the License at:
    https://opensource.org/licenses/MIT

    Author: Abdallah Salah
    Description: PowerShell script to export compliance search results by message ID.
#>

# Connect to SC Module
Connect-IPPSSession

# List of Message IDs to include in the search
$MessageIds = @(
    "<ea3cf85d-cc63-409b-bf11-58fab4e6f1ac@az.westus.microsoft.com>",
    "<535a13a8-2bd1-49a4-b1a3-f3a64701607d@az.centralus.microsoft.com>",
    "<CY8PR14MB6265BE0D33453F13F9C4DD8288CD2@CY8PR14MB6265.namprd14.prod.outlook.com>"
)

# Build a content query combining all Message IDs with OR
$Query = "MessageId:(" + ($MessageIds -join " OR ") + ")"

# Create the compliance search:
# - Name: the search identifier
# - ExchangeLocation: the mailbox to search 
# - ContentMatchQuery: the MessageId query we constructed
New-ComplianceSearch -Name $SearchName -ExchangeLocation admin@hybrid-pro.net -ContentMatchQuery $Query

# Define the name of the Compliance Search
$SearchName = "Compromised Mailbox"

# Save the Export search name into a variable by appending "_Export"
$SearchName_Export = $SearchName + "_Export"

# Start the Compliance Search (run this first if the search hasn't started yet)
Start-ComplianceSearch -Identity $SearchName

# Create an export action for the compliance search:
# - SharePointArchiveFormat: format of archive in SharePoint (SingleZip = one zip file)
# - ExchangeArchiveFormat: archive format for Exchange results (PerUserPst = separate PST per mailbox)
# - Format: FxStream is the internal format used for export
New-ComplianceSearchAction -SearchName $SearchName -Export -SharePointArchiveFormat SingleZip -ExchangeArchiveFormat PerUserPst -Format FxStream

# Check the status of the export action. This will return the current state of the export (e.g., Starting, Completed)
Get-ComplianceSearchAction -SearchName $SearchName

# After the export status is 'Completed', run the following to retrieve the download URL:
# Fetch the compliance search action details, including credentials (URL and SAS token) for downloading the results
$export = Get-ComplianceSearchAction $SearchName_Export -IncludeCredential

# Extract the "Results" property from the export, which contains the full text including the "Container URL" and "SAS token"
$results = $export.Results  

# Extract the container URL using split
$containerUrl = ($results -split "Container url:")[1].Split(';')[0].Trim()

# Extract the SAS token by splitting the string at "SAS token:" and taking the part after it
$sasToken = ($results -split "SAS token:")[1].Split(';')[0].Trim()

# Define the base URL for the export tool
$baseUrl = "https://complianceclientsdf.blob.core.windows.net/v16/Microsoft.Office.Client.Discovery.UnifiedExportTool.application"
$trace = 1
$lite = 1
$customizePst = 1

# URL Encode the container URL and the search name
$encodedContainerUrl = [uri]::EscapeDataString($containerUrl)
$encodedSearchName = [uri]::EscapeDataString($SearchName_Export)


# Construct the final download URL
$downloadUrl = "https://complianceclientsdf.blob.core.windows.net/v16/Microsoft.Office.Client.Discovery.UnifiedExportTool.application?source=$encodedContainerUrl&name=$encodedSearchName&trace=$trace&lite=$lite&customizepst=$customizePst"

# Output the final URL
Write-Host -ForegroundColor Green "To download the search results, please open the following URL in Microsoft Edge:`n`n$downloadUrl`n"

# Output the SAS token
Write-Host -ForegroundColor Green "Your passkey (SAS token) for the download is below:`n`n$sasToken`n"
