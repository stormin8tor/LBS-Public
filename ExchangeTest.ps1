$code = @'
# Exchange Permission Tool
function Show-ExchangeTool {
    Add-Type -AssemblyName System.Windows.Forms
    
    # Check for Exchange module
    $hasExchange = $false
    try { Get-Command Get-Mailbox -ErrorAction Stop | Out-Null; $hasExchange = $true } catch {}
    
    # Create form elements
    $form = New-Object -TypeName System.Windows.Forms.Form
    $form.Text = "Exchange Permission Tool"
    $form.Width = 500
    $form.Height = 400
    
    $y = 10
    
    # Mailbox section
    $lblMailbox = New-Object System.Windows.Forms.Label
    $lblMailbox.Text = "Mailbox:"
    $lblMailbox.Location = New-Object System.Drawing.Point(10, $y)
    $lblMailbox.Width = 100
    $form.Controls.Add($lblMailbox)
    
    $txtMailbox = New-Object System.Windows.Forms.TextBox
    $txtMailbox.Location = New-Object System.Drawing.Point(120, $y)
    $txtMailbox.Width = 350
    $form.Controls.Add($txtMailbox)
    $y += 30
    
    # Full Access section
    $lblFullAccess = New-Object System.Windows.Forms.Label
    $lblFullAccess.Text = "Full Access:"
    $lblFullAccess.Location = New-Object System.Drawing.Point(10, $y)
    $lblFullAccess.Width = 100
    $form.Controls.Add($lblFullAccess)
    
    $txtFullAccess = New-Object System.Windows.Forms.TextBox
    $txtFullAccess.Location = New-Object System.Drawing.Point(120, $y)
    $txtFullAccess.Width = 350
    $form.Controls.Add($txtFullAccess)
    $y += 30
    
    # Send As section
    $lblSendAs = New-Object System.Windows.Forms.Label
    $lblSendAs.Text = "Send As:"
    $lblSendAs.Location = New-Object System.Drawing.Point(10, $y)
    $lblSendAs.Width = 100
    $form.Controls.Add($lblSendAs)
    
    $txtSendAs = New-Object System.Windows.Forms.TextBox
    $txtSendAs.Location = New-Object System.Drawing.Point(120, $y)
    $txtSendAs.Width = 350
    $form.Controls.Add($txtSendAs)
    $y += 30
    
    # Results section
    $txtResults = New-Object System.Windows.Forms.TextBox
    $txtResults.Multiline = $true
    $txtResults.ScrollBars = "Vertical"
    $txtResults.Location = New-Object System.Drawing.Point(10, $y)
    $txtResults.Width = 460
    $txtResults.Height = 200
    $txtResults.ReadOnly = $true
    $form.Controls.Add($txtResults)
    $y += 210
    
    # Button 
    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = "Apply Permissions"
    $btnApply.Location = New-Object System.Drawing.Point(185, $y)
    $btnApply.Width = 130
    $form.Controls.Add($btnApply)
    
    # Set initial status message
    if ($hasExchange) {
        $txtResults.Text = "Ready to apply permissions."
    } else {
        $txtResults.Text = "Exchange module not detected. Please connect to Exchange first."
    }
    
    # Button click event
    $btnApply.Add_Click({
        if (-not $hasExchange) {
            $txtResults.Text = "Exchange module not loaded. Please run Connect-ExchangeOnline first."
            return
        }
        
        $mailbox = $txtMailbox.Text.Trim()
        if (-not $mailbox) {
            $txtResults.Text = "Please enter a mailbox name."
            return
        }
        
        $txtResults.Text = ""
        
        try {
            # Check if mailbox exists
            Get-Mailbox -Identity $mailbox -ErrorAction Stop | Out-Null
            $txtResults.AppendText("Found mailbox: $mailbox`r`n")
            
            # Process Full Access users
            $fullUsers = $txtFullAccess.Text.Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            foreach ($user in $fullUsers) {
                try {
                    Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping $false -ErrorAction Stop
                    $txtResults.AppendText("Added Full Access for: $user`r`n")
                } catch {
                    $txtResults.AppendText("Error adding Full Access for $user: $_`r`n") 
                }
            }
            
            # Process Send As users
            $sendAsUsers = $txtSendAs.Text.Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            foreach ($user in $sendAsUsers) {
                try {
                    Add-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                    $txtResults.AppendText("Added Send As for: $user`r`n")
                } catch {
                    $txtResults.AppendText("Error adding Send As for $user: $_`r`n")
                }
            }
            
            $txtResults.AppendText("Operations completed.`r`n")
        } catch {
            $txtResults.AppendText("Error: $_`r`n")
        }
    })
    
    # Show form
    [void]$form.ShowDialog()
}

# Run the tool
Show-ExchangeTool
'@

# Execute the code
Invoke-Expression $code
