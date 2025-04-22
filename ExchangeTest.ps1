# Exchange Permission Tool
# Save this entire file to GitHub

function Show-ExchangeTool {
    Add-Type -AssemblyName System.Windows.Forms
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Exchange Permission Tool"
    $form.Size = New-Object System.Drawing.Size(600, 450)
    $form.StartPosition = "CenterScreen"
    
    # Mailbox section
    $lblMailbox = New-Object System.Windows.Forms.Label
    $lblMailbox.Location = New-Object System.Drawing.Point(10, 20)
    $lblMailbox.Size = New-Object System.Drawing.Size(580, 20)
    $lblMailbox.Text = "Mailbox to Assign Permissions To:"
    $form.Controls.Add($lblMailbox)
    
    $txtMailbox = New-Object System.Windows.Forms.TextBox
    $txtMailbox.Location = New-Object System.Drawing.Point(10, 40)
    $txtMailbox.Size = New-Object System.Drawing.Size(560, 25)
    $form.Controls.Add($txtMailbox)
    
    # Full Access section
    $lblFullAccess = New-Object System.Windows.Forms.Label
    $lblFullAccess.Location = New-Object System.Drawing.Point(10, 70)
    $lblFullAccess.Size = New-Object System.Drawing.Size(580, 20)
    $lblFullAccess.Text = "Full Access Users (comma-separated):"
    $form.Controls.Add($lblFullAccess)
    
    $txtFullAccess = New-Object System.Windows.Forms.TextBox
    $txtFullAccess.Location = New-Object System.Drawing.Point(10, 90)
    $txtFullAccess.Size = New-Object System.Drawing.Size(560, 25)
    $form.Controls.Add($txtFullAccess)
    
    # Send As section
    $lblSendAs = New-Object System.Windows.Forms.Label
    $lblSendAs.Location = New-Object System.Drawing.Point(10, 120)
    $lblSendAs.Size = New-Object System.Drawing.Size(580, 20)
    $lblSendAs.Text = "Send As Users (comma-separated):"
    $form.Controls.Add($lblSendAs)
    
    $txtSendAs = New-Object System.Windows.Forms.TextBox
    $txtSendAs.Location = New-Object System.Drawing.Point(10, 140)
    $txtSendAs.Size = New-Object System.Drawing.Size(560, 25)
    $form.Controls.Add($txtSendAs)
    
    # Output section
    $txtOutput = New-Object System.Windows.Forms.TextBox
    $txtOutput.Location = New-Object System.Drawing.Point(10, 170)
    $txtOutput.Size = New-Object System.Drawing.Size(560, 180)
    $txtOutput.Multiline = $true
    $txtOutput.ScrollBars = "Vertical"
    $txtOutput.ReadOnly = $true
    $form.Controls.Add($txtOutput)
    
    # Check for Exchange module
    $hasExchange = $false
    try {
        Get-Command Add-MailboxPermission -ErrorAction Stop | Out-Null
        $hasExchange = $true
        $txtOutput.Text = "Ready. Exchange module loaded successfully."
    } catch {
        $txtOutput.Text = "Warning: Exchange module not detected.`r`nPlease connect to Exchange first using Connect-ExchangeOnline."
    }
    
    # Apply Button
    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Location = New-Object System.Drawing.Point(225, 360)
    $btnApply.Size = New-Object System.Drawing.Size(150, 35)
    $btnApply.Text = "Apply Permissions"
    $form.Controls.Add($btnApply)
    
    # Button click handler
    $btnApply.Add_Click({
        if (-not $hasExchange) {
            $txtOutput.Text = "Exchange module is not loaded. Please connect to Exchange first."
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
    
    [void]$form.ShowDialog()
}

# Run the tool
Show-ExchangeTool
