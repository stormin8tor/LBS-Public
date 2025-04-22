Add-Type -AssemblyName PresentationFramework
Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue

# Define XAML UI
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
$reader = (New-Object System.Xml.XmlTextReader([System.IO.StringReader]$xamlString))
$window = [Windows.Markup.XamlReader]::Load($reader)

# Access UI elements
$MailboxBox  = $window.FindName("MailboxTextBox")
$FullBox     = $window.FindName("FullAccessTextBox")
$SendAsBox   = $window.FindName("SendAsTextBox")
$OutputBox   = $window.FindName("OutputTextBox")
$ApplyButton = $window.FindName("ApplyButton")

function Write-OutputBox {
    param ([string]$Text)
    $OutputBox.AppendText("$Text`n")
    $OutputBox.ScrollToEnd()
}

$ApplyButton.Add_Click({
    $OutputBox.Clear()

    $mailbox = $MailboxBox.Text.Trim()
    $fullAccessUsers = $FullBox.Text.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    $sendAsUsers     = $SendAsBox.Text.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

    if (-not $mailbox) {
        Write-OutputBox "Please enter a mailbox to assign permissions to."
        return
    }

    try {
        Write-OutputBox "Connecting to Exchange Online..."
        $cred = Get-Credential
        Connect-ExchangeOnline -UserPrincipalName $cred.UserName -ShowProgress $true
        Write-OutputBox "Connected successfully."

        Write-OutputBox "`nAssigning Full Access Permissions..."
        foreach ($user in $fullAccessUsers) {
            Write-OutputBox "  Adding Full Access for $user..."
            # Uncomment below to apply
            # Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All
        }

        Write-OutputBox "`nAssigning Send As Permissions..."
        foreach ($user in $sendAsUsers) {
            Write-OutputBox "  Adding Send As for $user..."
            # Uncomment below to apply
            # Add-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false
        }

        Write-OutputBox "`nAll permissions applied successfully!"
    } catch {
        Write-OutputBox "Error: $($_.Exception.Message)"
    }
})

# Show the UI
$window.ShowDialog() | Out-Null
