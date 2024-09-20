<#
.SYNOPSIS
    This script converts Microsoft Teams chat messages stored in a JSON file into an HTML file, 
    representing the conversation in a format similar to what Teams uses, using TailwindCSS for styling.

.PARAMETER JsonFilePath
    The path to the JSON file containing the Teams messages.

.PARAMETER OutputHtmlPath
    The path where the generated HTML file will be saved.

.EXAMPLE
    .\Convert-TeamsMessagesToHtml.ps1 -JsonFilePath "C:\temp\output\messages.json" -OutputHtmlPath "C:\temp\output\conversation.html"

    This command will read the messages from the specified JSON file and generate an HTML conversation file using TailwindCSS.

.NOTES
    Author: Your Name
    Date: September 2024
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$JsonFilePath,   # Path to the JSON file with Teams messages

    [Parameter(Mandatory = $true)]
    [string]$OutputHtmlPath  # Path where the HTML file will be saved
)

# Load the Teams messages from the JSON file
try {
    $messages = Get-Content -Path $JsonFilePath | ConvertFrom-Json
} catch {
    Write-Host "Error: Failed to load or parse JSON file." -ForegroundColor Red
    exit
}

# Prepare the HTML structure with TailwindCSS included
$htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Teams Conversation</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
	body {
    font-size: 0.875rem; /* 14px */
}
pre, code {
    font-size: 0.75rem; /* 12px */
}
        pre {
            background-color: #f3f4f6;
            padding: 10px;
            border-radius: 4px;
            overflow-x: auto;
			width: 100%;
        }
        code {
            font-family: 'Courier New', Courier, monospace;
            background-color: #e5e7eb;
            padding: 2px 4px;
            border-radius: 4px;
			font-size: 0.75em;
			width: 100%;
        }
        img {
            max-width: 100%;
            height: auto;
            margin-top: 10px;
        }
        .avatar {
            width: 40px;
            height: 40px;
            display: inline-block;
            border-radius: 50%;
            font-size: 14px;
            font-weight: bold;
            color: white;
            text-align: center;
            line-height: 40px;
            margin-right: 10px;
            overflow: hidden;
            flex-shrink: 0;
        }
    </style>
</head>
<body class="bg-gray-100 p-6">
    <div class="max-w-6xl mx-auto bg-white shadow-lg rounded-lg overflow-hidden">
        <h1 class="text-2xl font-bold p-6 bg-blue-600 text-white">Teams Conversation</h1>
        <div class="p-4 space-y-4">
"@

# Generate initials from a name (inverted order)
function Get-Initials {
    param (
        [string]$name
    )

    $nameParts = $name -split '\s+'
    $initials = ''
    if ($nameParts.Count -ge 2 -and $nameParts[1].Length -ge 1) {
        $initials += $nameParts[1].Substring(0,1).ToUpper()
    } elseif ($nameParts.Count -ge 1 -and $nameParts[0].Length -ge 1) {
        $initials += $nameParts[0].Substring(0,1).ToUpper()
    }
    if ($nameParts.Count -ge 1 -and $nameParts[0].Length -ge 1) {
        $initials += $nameParts[0].Substring(0,1).ToUpper()
    }
    return $initials
}

# Generate a color for each person (based on a hash of the name)
function Get-UserColor {
    param (
        [string]$name
    )

    # Generate a pseudo-random number from the name (basic hashing)
    $hashValue = 0
    foreach ($char in $name.ToCharArray()) {
        $hashValue = [math]::Abs([int][char]$char + $hashValue)
    }

    # Convert the hash to a color from a predefined list of colors
    $colors = @("bg-red-500", "bg-blue-500", "bg-green-500", "bg-yellow-500", "bg-purple-500", "bg-pink-500", "bg-teal-500")
    $colorIndex = $hashValue % $colors.Count
    return $colors[$colorIndex]
}

# Function to clean up unnecessary HTML, like unwanted <p> tags
function CleanUp-Content {
    param (
        [string]$content
    )

    # Remove any <p> tags that contain only <br> or &nbsp;
    $content = $content -replace '<p.*?><br.*?></p>', ''
    $content = $content -replace '<p.*?>&nbsp;</p>', ''

    return $content
}

# Function to detect and render code blocks, clean up HTML, and handle images
function Format-MessageContent {
    param (
        [string]$content
    )
    
    # First, clean up unnecessary tags
    $content = CleanUp-Content -content $content

    # Detect code blocks, assumed to be wrapped by triple backticks ``` (like in Markdown)
    $content = [regex]::Replace($content, '```(.*?)```', '<pre><code>$1</code></pre>', 'Singleline')

    # Detect inline code (single backticks `) and replace with <code> tags
    $content = $content -replace '`(.+?)`', '<code>$1</code>'

    return $content
}

# Loop through each message and format it into HTML
foreach ($message in $messages) {
    # Ensure the message has content; skip empty or null messages
    if (-not $message.from) {
        continue  # Skip messages with no content
    }

    # Extract necessary fields (adjust based on your JSON structure)
    $sender = $message.from.user.displayName
    $initials = Get-Initials -name $sender
    $color = Get-UserColor -name $sender
    $content = Format-MessageContent -content $message.body.content
    $timestamp = [DateTime]::Parse($message.createdDateTime).ToString("dddd, MMMM dd, yyyy hh:mm tt")

    # Build the HTML for each message with TailwindCSS classes
    $htmlContent += @"
    <div class="chat-message p-4 border-b border-gray-200 flex items-start">
        <div class="avatar $color text-white">$initials</div>
        <div>
            <div class="header text-blue-500 font-semibold">$sender</div>
            <div class="timestamp text-gray-500 text-sm">$timestamp</div>
            <div class="content mt-2 text-gray-700">$content</div>
        </div>
    </div>
"@
}

# Close the HTML structure
$htmlContent += @"
        </div>
    </div>
</body>
</html>
"@

# Save the HTML content to the specified file
try {
    $htmlContent | Out-File -FilePath $OutputHtmlPath -Encoding utf8
    Write-Host "HTML file successfully created at $OutputHtmlPath" -ForegroundColor Green
} catch {
    Write-Host "Error: Failed to save HTML file." -ForegroundColor Red
}
