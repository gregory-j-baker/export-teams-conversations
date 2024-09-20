# Export and Convert Microsoft Teams Conversations

This repository provides PowerShell scripts to export Microsoft Teams conversations in JSON format and optionally convert them to HTML for easy viewing.

## Purpose

The two PowerShell scripts in this repository serve the following purposes:
1. **Get-TeamsChatMessagesWithToken.ps1**: Exports Microsoft Teams conversations to a JSON file by authenticating with a token.
2. **Convert-TeamsMessagesToHtml.ps1**: Converts exported JSON messages into an HTML format for better readability.

## Prerequisites

- PowerShell 5.1 or later
- Admin access to Microsoft Teams API (with necessary permissions)
- Azure AD App Registration for token-based authentication
- Access to Microsoft Graph API
- Basic knowledge of PowerShell

## How to Use

### 1. Export Teams Conversations to JSON

#### Script: `Get-TeamsChatMessagesWithToken.ps1`

1. Clone this repository or download the script file.
2. Run the PowerShell script with the following parameters:
   ```bash
   ./Get-TeamsChatMessagesWithToken.ps1 -ChatId <YourChatId> -AccessToken <YourAccessToken> -OutputFilePath <PathToSaveJson>
   ```
3. This script will authenticate using the provided token and export Teams messages into a JSON file at the specified path.
   
   You can obtain the access token by logging in to [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) with your Microsoft tenant credentials. After logging in, select the necessary permissions (such as `Chat.Read`), and generate the token.

### 2. Convert Teams Messages to HTML

#### Script: `Convert-TeamsMessagesToHtml.ps1`

1. Once you have the JSON file exported from the previous script, you can convert it to HTML.
2. Run the following command:
   ```bash
   ./Convert-TeamsMessagesToHtml.ps1 -JsonPath <PathToJsonFile> -HtmlOutputPath <PathToSaveHtml>
   ```
3. This script will read the JSON file and generate a formatted HTML file with the Teams conversations.

## Example

```bash
# Step 1: Export Teams Messages to JSON
./Get-TeamsChatMessagesWithToken.ps1 -ChatId "19:meeting_NjgxMjA0MGUtZmRiYy00MjY2LTlkYWMtOWJmYmEzYTA3YWZi@thread.v2" -AccessToken "eyJ..." -OutputFilePath "./teamsMessages.json"

# Step 2: Convert JSON to HTML
./Convert-TeamsMessagesToHtml.ps1 -JsonPath "./teamsMessages.json" -HtmlOutputPath "./teamsMessages.html"
```

## Token Generation

To use the `Get-TeamsChatMessagesWithToken.ps1` script, you'll need an access token for Microsoft Teams. This can be obtained through Azure AD or by logging in to [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer). Select the required permissions (e.g., `Chat.Read`), authenticate, and copy the generated token to use in the script.

## Sample Files

- [Sample Teams Messages JSON](samples/sample_teams_messages.json)
- [Sample Teams Messages HTML](samples/sample_teams_messages.html)

These sample files give an idea of the expected results when using the provided scripts to export Teams messages and convert them to HTML.

## License

This project is licensed under the MIT License.
