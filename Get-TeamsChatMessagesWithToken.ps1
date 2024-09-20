<#
.SYNOPSIS
    This script fetches all messages from a specific Microsoft Teams chat using the Microsoft Graph API.
    It handles pagination and saves the messages to a JSON file.

.DESCRIPTION
    The script uses a manually obtained OAuth 2.0 access token (either from Graph Explorer or an Azure AD App).
    This token is passed in the Authorization header to authenticate the API requests.
    The script will fetch messages from a specific chat, handle pagination using the @odata.nextLink if present, and save the retrieved messages to a specified JSON file.

.PARAMETER ChatId
    The ID of the chat from which to fetch messages. This is required to form the API endpoint.

.PARAMETER OutputFilePath
    The file path where the resulting messages will be saved in JSON format.

.PARAMETER AccessToken
    The OAuth 2.0 access token obtained from Graph Explorer or your Azure AD application. This token is required for API authentication.

.EXAMPLE
    .\Get-TeamsChatMessagesWithToken.ps1 -ChatId "19:meeting_NjgxMjA0MGUtZmRiYy00MjY2LTlkYWMtOWJmYmEzYTA3YWZi@thread.v2" -OutputFilePath "C:\temp\output\messages.json" -AccessToken "YOUR_ACCESS_TOKEN_HERE"

    This command will fetch all messages from the specified chat and save them to the specified JSON file using the provided access token.

.NOTES
    Author: Your Name
    Date: September 2024

#>

# Set Parameters for the Chat ID, Output File Path, and Access Token
param (
    [Parameter(Mandatory = $true)]
    [string]$ChatId,  # The chat ID for which to retrieve messages

    [Parameter(Mandatory = $true)]
    [string]$OutputFilePath,  # The output file where the messages will be saved in JSON format

    [Parameter(Mandatory = $true)]
    [string]$AccessToken  # The manually obtained OAuth 2.0 access token
)

# Define the Graph API URL using the ChatId parameter
$graphApiUrl = "https://graph.microsoft.com/beta/me/chats/$ChatId/messages?$top=50"

# Import the necessary .NET List type
$listType = [System.Collections.Generic.List[System.Object]]

# Create an empty list to store all messages
$allMessages = New-Object $listType

# Function to fetch messages and follow pagination
function Get-Messages {
    param (
        [string]$url
    )

    while ($url) {
        try {
            Write-Host "Fetching messages from: $url" -ForegroundColor Cyan

            # Call the Graph API with manually supplied token
            $response = Invoke-RestMethod -Uri $url -Method Get -Headers @{
                Authorization = "Bearer $AccessToken"
            }

            if ($response -and $response.value) {
                Write-Host "Fetched $($response.value.Count) messages" -ForegroundColor Green

                # Append the current batch of messages
                $allMessages.AddRange($response.value)
            } else {
                Write-Host "No messages found or response is empty" -ForegroundColor Yellow
            }

            # If there is a next page link, update the URL; otherwise, exit the loop
            $url = if ($response.'@odata.nextLink') { 
                Write-Host "Found next page: $($response.'@odata.nextLink')" -ForegroundColor Cyan
                $response.'@odata.nextLink'
            } else { 
                Write-Host "No more pages to fetch." -ForegroundColor Cyan
                $null 
            }
        }
        catch {
            Write-Host "Error fetching messages: $_" -ForegroundColor Red
            break
        }
    }
}

# Ensure the output directory exists
$directory = [System.IO.Path]::GetDirectoryName($OutputFilePath)

if (-not (Test-Path $directory)) {
    Write-Host "Directory $directory does not exist. Creating..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $directory
} else {
    Write-Host "Directory exists: $directory" -ForegroundColor Green
}

# Fetch all messages using the API URL
Get-Messages -url $graphApiUrl

# Debugging: Check if there are any messages
if ($allMessages.Count -eq 0) {
    Write-Host "No messages were fetched. Exiting..." -ForegroundColor Red
    exit
}

# Debugging: Display the total number of fetched messages
Write-Host "Total messages fetched: $($allMessages.Count)" -ForegroundColor Green

# Debugging: Display a sample of the fetched messages (first 1 message, if available)
if ($allMessages.Count -gt 0) {
    Write-Host "Sample message data: " -ForegroundColor Cyan
    $allMessages[0] | ConvertTo-Json -Depth 3 | Write-Host
}

# Convert the messages to JSON and save to the specified file
try {
    $jsonData = $allMessages | ConvertTo-Json -Depth 10

    # Debugging: Ensure the JSON data is not null or empty
    if ([string]::IsNullOrEmpty($jsonData)) {
        Write-Host "Error: JSON data is empty or null!" -ForegroundColor Red
    } else {
        Write-Host "Saving messages to $OutputFilePath" -ForegroundColor Green
        $jsonData | Out-File -FilePath $OutputFilePath -Encoding utf8
        Write-Host "Messages saved successfully!" -ForegroundColor Green
    }
}
catch {
    Write-Host "Error saving JSON data to file: $_" -ForegroundColor Red
}