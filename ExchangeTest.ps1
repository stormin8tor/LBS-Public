Add-Type -AssemblyName PresentationFramework

$xaml = @"
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

        <Button Grid.Row="4" Content="Apply Permissions" Height="35" Width="150"
                HorizontalAlignment="Center" x:Name="ApplyButton"/>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlTextReader ([System.IO.StringReader]::new($xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

# Grab controls by name
$MailboxBox = $window.FindName("MailboxTextBox")
$FullBox    = $window.FindName("FullAccessTextBox")
$SendAsBox  = $window.FindName("SendAsTextBox")
$OutputBox  = $window.FindName("OutputTextBox")
$ApplyBtn   = $window.FindName("ApplyButton")

# Define what happens when clicking the button
$Apply_Click = {
    $mailbox = $MailboxBox.Text.Trim()
    $fullUsers = $FullBox.Text.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $sendAsUsers = $SendAsBox.Text.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }

    $OutputBox.Clear()

    if (-not $mailbox) {
        $OutputBox.AppendText("Mailbox cannot be empty.`n")
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
}

$ApplyBtn.Add_Click($Apply_Click)

$window.ShowDialog() | Out-Null
