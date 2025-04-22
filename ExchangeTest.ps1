# Exchange Permission Tool
# This script is designed to be executed directly from GitHub with: irm <url> | iex

Add-Type -AssemblyName PresentationFramework

# Using a base64 encoded string for the XAML content to avoid parsing issues when using irm | iex
$encodedXaml = 'PFdpbmRvdyB4bWxucz0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS93aW5meC8yMDA2L3hhbWwvcHJlc2VudGF0aW9uIg0KICAgICAgICB4bWxuczp4PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL3dpbmZ4LzIwMDYveGFtbCINCiAgICAgICAgVGl0bGU9IkV4Y2hhbmdlIFBlcm1pc3Npb24gVG9vbCIgSGVpZ2h0PSI0NTAiIFdpZHRoPSI2MDAiPg0KICAgIDxHcmlkIE1hcmdpbj0iMTAiPg0KICAgICAgICA8R3JpZC5Sb3dEZWZpbml0aW9ucz4NCiAgICAgICAgICAgIDxSb3dEZWZpbml0aW9uIEhlaWdodD0iQXV0byIvPg0KICAgICAgICAgICAgPFJvd0RlZmluaXRpb24gSGVpZ2h0PSJBdXRvIi8+DQogICAgICAgICAgICA8Um93RGVmaW5pdGlvbiBIZWlnaHQ9IkF1dG8iLz4NCiAgICAgICAgICAgIDxSb3dEZWZpbml0aW9uIEhlaWdodD0iKiIvPg0KICAgICAgICAgICAgPFJvd0RlZmluaXRpb24gSGVpZ2h0PSJBdXRvIi8+DQogICAgICAgIDwvR3JpZC5Sb3dEZWZpbml0aW9ucz4NCiAgICAgICAgPFN0YWNrUGFuZWwgR3JpZC5Sb3c9IjAiIE1hcmdpbj0iMCwwLDAsMTAiPg0KICAgICAgICAgICAgPFRleHRCbG9jayBUZXh0PSJNYWlsYm94IHRvIEFzc2lnbiBQZXJtaXNzaW9ucyBUbzoiIC8+DQogICAgICAgICAgICA8VGV4dEJveCB4Ok5hbWU9Ik1haWxib3hUZXh0Qm94IiBIZWlnaHQ9IjI1Ii8+DQogICAgICAgIDwvU3RhY2tQYW5lbD4NCiAgICAgICAgPFN0YWNrUGFuZWwgR3JpZC5Sb3c9IjEiIE1hcmdpbj0iMCwwLDAsMTAiPg0KICAgICAgICAgICAgPFRleHRCbG9jayBUZXh0PSJGdWxsIEFjY2VzcyBVc2VycyAoY29tbWEtc2VwYXJhdGVkKToiIC8+DQogICAgICAgICAgICA8VGV4dEJveCB4Ok5hbWU9IkZ1bGxBY2Nlc3NUZXh0Qm94IiBIZWlnaHQ9IjI1Ii8+DQogICAgICAgIDwvU3RhY2tQYW5lbD4NCiAgICAgICAgPFN0YWNrUGFuZWwgR3JpZC5Sb3c9IjIiIE1hcmdpbj0iMCwwLDAsMTAiPg0KICAgICAgICAgICAgPFRleHRCbG9jayBUZXh0PSJTZW5kIEFzIFVzZXJzIChjb21tYS1zZXBhcmF0ZWQpOiIgLz4NCiAgICAgICAgICAgIDxUZXh0Qm94IHg6TmFtZT0iU2VuZEFzVGV4dEJveCIgSGVpZ2h0PSIyNSIvPg0KICAgICAgICA8L1N0YWNrUGFuZWw+DQogICAgICAgIDxUZXh0Qm94IHg6TmFtZT0iT3V0cHV0VGV4dEJveCIgR3JpZC5Sb3c9IjMiIElzUmVhZE9ubHk9IlRydWUiDQogICAgICAgICAgICAgICAgIFRleHRXcmFwcGluZz0iV3JhcCIgVmVydGljYWxTY3JvbGxCYXJWaXNpYmlsaXR5PSJBdXRvIiBNYXJnaW49IjAsMCwwLDEwIi8+DQogICAgICAgIDxCdXR0b24gR3JpZC5Sb3c9IjQiIENvbnRlbnQ9IkFwcGx5IFBlcm1pc3Npb25zIiBIZWlnaHQ9IjM1IiBXaWR0aD0iMTUwIg0KICAgICAgICAgICAgICAgIEhvcml6b250YWxBbGlnbm1lbnQ9IkNlbnRlciIgeDpOYW1lPSJBcHBseUJ1dHRvbiIvPg0KICAgIDwvR3JpZD4NCjwvV2luZG93Pg=='

# Convert the base64 string back to XAML
$xaml = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($encodedXaml))

# Check if Exchange module is loaded
$exchangeModuleLoaded = $false
try {
    # Check if we have Exchange cmdlets available
    Get-Command Add-MailboxPermission -ErrorAction Stop | Out-Null
    $exchangeModuleLoaded = $true
} catch {
    # Module not loaded - will notify user in the UI
}

try {
    # Fix XML parsing issues by using StringReader
    $reader = New-Object System.Xml.XmlTextReader ([System.IO.StringReader]::new($xaml))
    # Load XAML and create the window
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Host "Error loading XAML interface: $_" -ForegroundColor Red
    return
} finally {
    if ($reader) {
        $reader.Close()
    }
}

if (-not $window) {
    Write-Host "Failed to create the application window." -ForegroundColor Red
    return
}

# Retrieve controls by name
$MailboxBox = $window.FindName("MailboxTextBox")
$FullBox    = $window.FindName("FullAccessTextBox")
$SendAsBox  = $window.FindName("SendAsTextBox")
$OutputBox  = $window.FindName("OutputTextBox")
$ApplyBtn   = $window.FindName("ApplyButton")

# Button click handler
$Apply_Click = {
    if (-not $exchangeModuleLoaded) {
        $OutputBox.AppendText("Exchange module is not loaded. Please connect to Exchange first.`n")
        $OutputBox.AppendText("Run 'Connect-ExchangeOnline' or import the appropriate Exchange module before using this tool.`n")
        return
    }

    $mailbox = $MailboxBox.Text.Trim()
    $fullUsers = $FullBox.Text.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $sendAsUsers = $SendAsBox.Text.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    
    $OutputBox.Clear()
    
    if (-not $mailbox) {
        $OutputBox.AppendText("Mailbox cannot be empty.`n")
        return
    }
    
    # Verify mailbox exists
    try {
        Get-Mailbox -Identity $mailbox -ErrorAction Stop | Out-Null
        $OutputBox.AppendText("Found mailbox: $mailbox`n")
    } catch {
        $OutputBox.AppendText("Error: Mailbox '$mailbox' not found or access denied. $_`n")
        return
    }
    
    foreach ($user in $fullUsers) {
        try {
            Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping $false
            $OutputBox.AppendText("Granted FullAccess to $user`n")
        } catch {
            $OutputBox.AppendText("Failed FullAccess for $user: $_`n")
        }
    }
    
    foreach ($user in $sendAsUsers) {
        try {
            Add-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false
            $OutputBox.AppendText("Granted SendAs to $user`n")
        } catch {
            $OutputBox.AppendText("Failed SendAs for $user: $_`n")
        }
    }
    
    $OutputBox.AppendText("Operation completed.`n")
}

$ApplyBtn.Add_Click($Apply_Click)

# Add status message at startup
if (-not $exchangeModuleLoaded) {
    $OutputBox.Text = "Warning: Exchange module not detected.`nPlease connect to Exchange first using Connect-ExchangeOnline or the appropriate Exchange connection cmdlet."
} else {
    $OutputBox.Text = "Ready. Exchange module loaded successfully."
}

$window.ShowDialog() | Out-Null
