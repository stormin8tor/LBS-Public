# Exchange Permission Tool - Direct Execution Version
# This script is designed to be executed directly from GitHub with: irm <url> | iex

function Show-ExchangePermissionTool {
    # Check if Exchange module is loaded
    $exchangeModuleLoaded = $false
    try {
        # Check if we have Exchange cmdlets available
        Get-Command Add-MailboxPermission -ErrorAction Stop | Out-Null
        $exchangeModuleLoaded = $true
    } catch {
        # Module not loaded - will handle in the UI
    }

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Windows.Forms
    
    # Create the form without using XAML
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Exchange Permission Tool"
    $form.Size = New-Object System.Drawing.Size(600, 450)
    $form.StartPosition = "CenterScreen"
    
    # Mailbox Label
    $lblMailbox = New-Object System.Windows.Forms.Label
    $lblMailbox.Location = New-Object System.Drawing.Point(10, 20)
    $lblMailbox.Size = New-Object System.Drawing.Size(400, 20)
    $lblMailbox.Text = "Mailbox to Assign Permissions To:"
    $form.Controls.Add($lblMailbox)
    
    # Mailbox Textbox
    $txtMailbox = New-Object System.Windows.Forms.TextBox
    $txtMailbox.Location = New-Object System.Drawing.Point(10, 40)
    $txtMailbox.Size = New-Object System.Drawing.Size(560, 25)
    $form.Controls.Add($txtMailbox)
    
    # FullAccess Label
    $lblFullAccess = New-Object System.Windows.Forms.Label
    $lblFullAccess.Location = New-Object System.Drawing.Point(10, 70)
    $lblFullAccess.Size = New-Object System.Drawing.Size(400, 20)
    $lblFullAccess.Text = "Full Access Users (comma-separated):"
    $form.Controls.Add($lblFullAccess)
    
    # FullAccess Textbox
    $txtFullAccess = New-Object System.Windows.Forms.TextBox
    $txtFullAccess.Location = New-Object System.Drawing.Point(10, 90)
    $txtFullAccess.Size = New-Object System.Drawing.Size(560, 25)
    $form.Controls.Add($txtFullAccess)
    
    # SendAs Label
    $lblSendAs = New-Object System.Windows.Forms.Label
    $lblSendAs.Location = New-Object System.Drawing.Point(10, 120)
    $lblSendAs.Size = New-Object System.Drawing.Size(400, 20)
    $lblSendAs.Text = "Send As Users (comma-separated):"
    $form.Controls.Add($lblSendAs)
    
    # SendAs Textbox
    $txtSendAs = New-Object System.Windows.Forms.TextBox
    $txtSendAs.Location = New-Object System.Drawing.Point(10, 140)
    $txtSendAs.Size = New-Object System.Drawing.Size(560, 25)
    $form.Controls.Add($txtSendAs)
    
    # Output Textbox
    $txtOutput = New-Object System.Windows.Forms.TextBox
    $txtOutput.Location = New-Object System.Drawing.Point(10, 170)
    $txtOutput.Size = New-Object System.Drawing.Size(560, 180)
    $txtOutput.Multiline = $true
    $txtOutput.ScrollBars = "Vertical"
    $txtOutput.ReadOnly = $true
    $form.Controls.Add($txtOutput)
    
    # Apply Button
    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Location = New-Object System.Drawing.Point(225, 360)
    $btnApply.Size = New-Object System.Drawing.Size(150, 35)
    $btnApply.Text = "Apply Permissions"
    $form.Controls.Add($btnApply)
    
    # Add status message at startup
    if (-not $exchangeModuleLoaded) {
        $txtOutput.Text = "Warning: Exchange module not detected.`r`nPlease connect to Exchange first using Connect-ExchangeOnline or the appropriate Exchange connection cmdlet."
    } else {
        $txtOutput.Text = "Ready. Exchange module loaded successfully."
    }
    
    # Button click handler
    $btnApply.Add_Click({
        if (-not $exchangeModuleLoaded) {
            $txtOutput.Text = "Exchange module is not loaded. Please connect to Exchange first.`r`nRun 'Connect-ExchangeOnline' or import the appropriate Exchange module before using this tool."
            return
        }

        $mailbox = $txtMailbox.Text.Trim()
        $fullUsers = $txtFullAccess.Text.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        $sendAsUsers = $txtSendAs.Text.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        
        $txtOutput.Clear()
        
        if (-not $mailbox) {
            $txtOutput.AppendText("Mailbox cannot be empty.`r`n")
            return
        }
        
        # Verify mailbox exists
        try {
            Get-Mailbox -Identity $mailbox -ErrorAction Stop | Out-Null
            $txtOutput.AppendText("Found mailbox: $mailbox`r`n")
        } catch {
            $txtOutput.AppendText("Error: Mailbox '$mailbox' not found or access denied. $_`r`n")
            return
        }
        
        foreach ($user in $fullUsers) {
            try {
                Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping $false
                $txtOutput.AppendText("Granted FullAccess to $user`r`n")
            } catch {
                $txtOutput.AppendText("Failed FullAccess for $user: $_`r`n")
            }
        }
        
        foreach ($user in $sendAsUsers) {
            try {
                Add-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false
                $txtOutput.AppendText("Granted SendAs to $user`r`n")
            } catch {
                $txtOutput.AppendText("Failed SendAs for $user: $_`r`n")
            }
        }
        
        $txtOutput.AppendText("Operation completed.`r`n")
    })
    
    # Show the form
    $form.ShowDialog() | Out-Null
}

# Execute the function to display the form
Show-ExchangePermissionTool
