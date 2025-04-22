Add-Type -AssemblyName PresentationFramework

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
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
                HorizontalAlignment="Center" Click="Apply_Click"/>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get UI elements
$MailboxBox    = $window.FindName("MailboxTextBox")
$FullBox       = $window.FindName("FullAccessTextBox")
$SendAsBox     = $window.FindName("SendAsTextBox")
$OutputBox     = $window.FindName("OutputTextBox")

# Event Handler
$Apply_Click = {
    $mailbox = $MailboxBox.Text.Trim()
    $fullUsers = $FullBox.Text -split "," | ForEach-Object { $_.Trim() }
    $sendAsUsers = $SendAsBox.Text -split "," | ForEach-Object { $_.Trim() }

    $OutputBox.Text = "Connecting to Exchange Online...`n"
    
    try {
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Install-Module ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
        }
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        $cred = Get-Credential
        Connect-ExchangeOnline -UserPrincipalName $cred.UserName

        foreach ($user in $fullUsers) {
            if ($user) {
                $OutputBox.AppendText("Adding Full Access for $user...`n")
                # Uncomment to apply
                # Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All
            }
        }

        foreach ($user in $sendAsUsers) {
            if ($user) {
                $OutputBox.AppendText("Adding Send As for $user...`n")
                # Uncomment to apply
                # Add-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs
            }
        }

        $OutputBox.AppendText("All done.`n")
    } catch {
        $OutputBox.AppendText("Error: $_`n")
    }
}

$window.FindName("Apply_Click").Add_Click($Apply_Click)
$window.ShowDialog()
