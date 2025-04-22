Add-Type -AssemblyName PresentationFramework

# Define XAML as a string (not [xml]!) and load using XmlTextReader
$xamlString = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Exchange Permission Tool" Height="450" Width="600">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,10">
            <TextBlock Text="Mailbox to Assign Permissions To:" />
            <TextBox x:Name="MailboxTextBox" Height="25"/>
        </StackPanel>

        <StackPanel Grid.Row="1" Margin="0,0,0,10">
            <TextBlock Text="Full Access Users (comma-separated):" />
            <TextBox x:Name="FullAccessTextBox" Height="25"/>
        </StackPanel>

        <StackPanel Grid.Row="2" Margin="0,0,0,10">
            <TextBlock Text="Send As Users (comma-separated):" />
            <TextBox x:Name="SendAsTextBox" Height="25"/>
        </StackPanel>

        <TextBox x:Name="OutputTextBox" Grid.Row="3" IsReadOnly="True"
                 TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Margin="0,0,0,10"/>

        <Button x:Name="ApplyButton" Grid.Row="4" Content="Apply Permissions" Height="35" Width="150"
                HorizontalAlignment="Center"/>
    </Grid>
</Window>
"@

# Load the XAML
$reader = New-Object System.Xml.XmlTextReader([System.IO.StringReader]$xamlString)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get references to UI elements
$MailboxBox = $window.FindName("MailboxTextBox")
$FullBox    = $window.FindName("FullAccessTextBox")
$SendAsBox  = $window.FindName("SendAsTextBox")
$OutputBox  = $window.FindName("OutputTextBox")
$ApplyButton = $window.FindName("ApplyButton")

# Define the click handler
$Apply_Click = {
    $mailbox = $MailboxBox.Text.Trim()
    $fullUsers = $FullBox.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $sendAsUsers = $SendAsBox.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }

    if (-not $mailbox) {
        $OutputBox.Text += "Please enter a mailbox.`n"
        return
    }

    $OutputBox.Text = "Applying permissions to: $mailbox`n"

    foreach ($user in $fullUsers) {
        try {
            Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -AutoMapping $false -ErrorAction Stop
            $OutputBox.Text += "Granted FullAccess to $user`n"
        } catch {
            $OutputBox.Text += "Failed to grant FullAccess to $user: $_`n"
        }
    }

    foreach ($user in $sendAsUsers) {
        try {
            Add-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false -ErrorAction Stop
            $OutputBox.Text += "Granted SendAs to $user`n"
        } catch {
            $OutputBox.Text += "Failed to grant SendAs to $user: $_`n"
        }
    }
}

# Hook up the event
$ApplyButton.Add_Click($Apply_Click)

# Show the window
$window.ShowDialog() | Out-Null
