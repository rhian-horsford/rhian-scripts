#requires -Version 7.6

[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$Identity,

    [ValidateSet('Both', 'Alias', 'Email')]
    [string]$Lookup = 'Both'
)

Set-StrictMode -Version 2.0

function Show-MailboxLookupWindow {
    param([string]$ScriptPath, [string]$InitialLookup)

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Mailbox Lookup" Width="1080" Height="560" MinWidth="720" MinHeight="420"
        WindowStartupLocation="CenterScreen" Background="{DynamicResource Surface}"
        Foreground="{DynamicResource Text}">
    <Window.Resources>
        <SolidColorBrush x:Key="Surface" Color="#F4F6FA"/>
        <SolidColorBrush x:Key="Field" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="Text" Color="#202A38"/>
        <SolidColorBrush x:Key="Border" Color="#C8CFDA"/>
        <SolidColorBrush x:Key="Error" Color="#B4233B"/>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource Field}"/>
            <Setter Property="Foreground" Value="{DynamicResource Text}"/>
            <Setter Property="CaretBrush" Value="{DynamicResource Text}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
            <Setter Property="Padding" Value="8,6"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="16,7"/>
            <Setter Property="MinWidth" Value="100"/>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="{DynamicResource Field}"/>
            <Setter Property="Foreground" Value="{DynamicResource Text}"/>
            <Setter Property="Padding" Value="8,5"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{DynamicResource Surface}"/>
            <Setter Property="Foreground" Value="{DynamicResource Text}"/>
            <Setter Property="Padding" Value="8"/>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="Padding" Value="6"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>
    </Window.Resources>
    <Grid Margin="22">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Mailbox lookup" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,18"/>
        <Grid Grid.Row="1" Margin="0,0,0,16">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="180"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Margin="0,0,12,0">
                <Label Target="{Binding ElementName=IdentityBox}" Foreground="{DynamicResource Text}" Content="_Alias or full email address"/>
                <TextBox x:Name="IdentityBox" MinHeight="34"/>
            </StackPanel>
            <StackPanel Grid.Column="1" Margin="0,0,12,0">
                <Label Target="{Binding ElementName=LookupBox}" Foreground="{DynamicResource Text}" Content="_Lookup"/>
                <ComboBox x:Name="LookupBox" MinHeight="34" SelectedIndex="0">
                    <ComboBoxItem Content="Alias and email" Tag="Both"/>
                    <ComboBoxItem Content="Alias only" Tag="Alias"/>
                    <ComboBoxItem Content="Email only" Tag="Email"/>
                </ComboBox>
            </StackPanel>
            <Button x:Name="SearchButton" Grid.Column="2" Content="_Search" IsDefault="True"
                    IsEnabled="False" VerticalAlignment="Bottom" Height="34"/>
        </Grid>
        <DataGrid x:Name="ResultsGrid" Grid.Row="2" AutoGenerateColumns="False" IsReadOnly="True"
                  CanUserAddRows="False" CanUserDeleteRows="False" SelectionMode="Single"
                  HeadersVisibility="Column" RowHeaderWidth="0" MinRowHeight="34"
                  Background="{DynamicResource Field}" Foreground="{DynamicResource Text}"
                  RowBackground="{DynamicResource Field}" BorderBrush="{DynamicResource Border}"
                  GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="{DynamicResource Border}"
                  ScrollViewer.HorizontalScrollBarVisibility="Auto">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Display name" Binding="{Binding DisplayName}" Width="160"/>
                <DataGridTextColumn Header="Alias" Binding="{Binding Alias}" Width="130"/>
                <DataGridTextColumn Header="Primary email" Binding="{Binding PrimarySmtpAddress}" Width="220"/>
                <DataGridTextColumn Header="Mailbox type" Binding="{Binding RecipientTypeDetails}" Width="140"/>
                <DataGridTextColumn Header="Email addresses" Binding="{Binding EmailAddresses}" Width="300"/>
            </DataGrid.Columns>
        </DataGrid>
        <TextBlock x:Name="StatusText" Grid.Row="3" Text="Not connected" TextWrapping="Wrap" Margin="0,12,0,10"/>
        <ProgressBar x:Name="Progress" Grid.Row="4" Height="4" Visibility="Collapsed"/>
    </Grid>
</Window>
'@
    $reader = [System.Xml.XmlNodeReader]::new([xml]$xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    $reader.Close()
    $identityBox = $window.FindName('IdentityBox')
    $lookupBox = $window.FindName('LookupBox')
    $searchButton = $window.FindName('SearchButton')
    $resultsGrid = $window.FindName('ResultsGrid')
    $statusText = $window.FindName('StatusText')
    $progress = $window.FindName('Progress')
    foreach ($item in $lookupBox.Items) {
        if ($item.Tag -eq $InitialLookup) { $lookupBox.SelectedItem = $item }
    }

    $uiState = @{ PowerShell = $null; Pending = $null; Theme = $null }
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.ThreadOptions = 'ReuseThread'
    $runspace.Open()
    $worker = {
        param($Path, $SearchIdentity, $SearchLookup)
        $ErrorActionPreference = 'Stop'
        Import-Module ExchangeOnlineManagement -MinimumVersion 3.7.2 -ErrorAction Stop
        $connections = @(Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' })
        if ($connections.Count -eq 0) {
            Connect-ExchangeOnline -DisableWAM -ShowBanner:$false -ErrorAction Stop
        }
        & $Path -Identity $SearchIdentity -Lookup $SearchLookup
    }
    $updateTheme = {
        $preference = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name AppsUseLightTheme -ErrorAction SilentlyContinue
        $light = $null -eq $preference -or [int]$preference.AppsUseLightTheme -ne 0
        if ($uiState.Theme -ne $light) {
            $uiState.Theme = $light
            $colors = if ($light) {
                @{ Surface = '#F4F6FA'; Field = '#FFFFFF'; Text = '#202A38'; Border = '#C8CFDA'; Error = '#B4233B' }
            }
            else {
                @{ Surface = '#131417'; Field = '#17191E'; Text = '#F0F2F6'; Border = '#383C47'; Error = '#FF7185' }
            }
            foreach ($key in $colors.Keys) {
                $window.Resources[$key] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors[$key]))
            }
        }
    }
    & $updateTheme

    $identityBox.Add_TextChanged({
        $searchButton.IsEnabled = $null -eq $uiState.PowerShell -and -not [string]::IsNullOrWhiteSpace($identityBox.Text)
        $resultsGrid.ItemsSource = $null
        $statusText.Text = ''
    })
    $lookupBox.Add_SelectionChanged({
        $resultsGrid.ItemsSource = $null
        $statusText.Text = ''
    })
    $searchButton.Add_Click({
        if ($uiState.PowerShell -or [string]::IsNullOrWhiteSpace($identityBox.Text)) { return }
        $resultsGrid.ItemsSource = $null
        $statusText.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'Text')
        $statusText.Text = 'Connecting / searching...'
        $identityBox.IsEnabled = $false
        $lookupBox.IsEnabled = $false
        $searchButton.IsEnabled = $false
        $progress.Visibility = 'Visible'
        $progress.IsIndeterminate = $true
        try {
            $uiState.PowerShell = [powershell]::Create()
            $uiState.PowerShell.Runspace = $runspace
            $null = $uiState.PowerShell.AddScript($worker.ToString()).AddArgument($ScriptPath).
                AddArgument($identityBox.Text).AddArgument([string]$lookupBox.SelectedItem.Tag)
            $uiState.Pending = $uiState.PowerShell.BeginInvoke()
        }
        catch {
            if ($uiState.PowerShell) { $uiState.PowerShell.Dispose() }
            $uiState.PowerShell = $null
            $uiState.Pending = $null
            $statusText.Text = "Lookup failed: $($_.Exception.Message)"
            $statusText.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'Error')
            $identityBox.IsEnabled = $true
            $lookupBox.IsEnabled = $true
            $searchButton.IsEnabled = $true
            $progress.Visibility = 'Collapsed'
        }
    })

    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(250)
    $timer.Add_Tick({
        & $updateTheme
        if (-not $uiState.Pending -or -not $uiState.Pending.IsCompleted) { return }
        try {
            $rows = @($uiState.PowerShell.EndInvoke($uiState.Pending))
            if ($uiState.PowerShell.HadErrors) {
                throw ($uiState.PowerShell.Streams.Error | Out-String)
            }
            $resultsGrid.ItemsSource = @(foreach ($row in $rows) {
                [pscustomobject]@{
                    DisplayName = $row.DisplayName
                    Alias = $row.Alias
                    PrimarySmtpAddress = [string]$row.PrimarySmtpAddress
                    RecipientTypeDetails = [string]$row.RecipientTypeDetails
                    EmailAddresses = $row.EmailAddresses -join '; '
                }
            })
            $statusText.Text = if ($rows.Count -eq 0) { 'No mailbox found.' } else { "Found $($rows.Count) mailbox(es)." }
        }
        catch {
            $statusText.Text = "Lookup failed: $($_.Exception.Message)"
            $statusText.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'Error')
        }
        finally {
            $uiState.PowerShell.Dispose()
            $uiState.PowerShell = $null
            $uiState.Pending = $null
            $identityBox.IsEnabled = $true
            $lookupBox.IsEnabled = $true
            $searchButton.IsEnabled = -not [string]::IsNullOrWhiteSpace($identityBox.Text)
            $progress.IsIndeterminate = $false
            $progress.Visibility = 'Collapsed'
        }
    })
    $window.Add_Closing({
        param($sender, $eventArgs)
        if ($uiState.PowerShell) {
            $eventArgs.Cancel = $true
            $statusText.Text = 'A lookup is still running. Complete or cancel browser sign-in before closing.'
        }
    })
    $window.Add_ContentRendered({ $null = $identityBox.Focus() })
    $timer.Start()
    try { $null = $window.ShowDialog() }
    finally {
        $timer.Stop()
        $runspace.Close()
        $runspace.Dispose()
    }
}

if (-not $PSBoundParameters.ContainsKey('Identity')) {
    Show-MailboxLookupWindow -ScriptPath $PSCommandPath -InitialLookup $Lookup
    return
}

$searchIdentity = $Identity.Trim()
if ([string]::IsNullOrWhiteSpace($searchIdentity)) {
    throw 'Enter a mailbox alias or a full email address.'
}

try {
    Import-Module ExchangeOnlineManagement -MinimumVersion 3.7.0 -ErrorAction Stop
    $connections = @(Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' })
    if ($connections.Count -eq 0) {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    }

    $escapedIdentity = $searchIdentity.Replace("'", "''")
    $aliasFilter = "Alias -eq '$escapedIdentity'"
    $emailFilter = "EmailAddresses -eq 'smtp:$escapedIdentity'"
    $filter = switch ($Lookup) {
        'Alias' { $aliasFilter }
        'Email' { $emailFilter }
        'Both' { "($aliasFilter) -or ($emailFilter)" }
    }

    $mailboxes = @(Get-EXOMailbox -Filter $filter -ResultSize Unlimited -Properties EmailAddresses -ErrorAction Stop)
    if ($mailboxes.Count -eq 0) {
        Write-Host "No mailbox found for '$searchIdentity' ($Lookup lookup)."
        return
    }

    $mailboxes | Select-Object DisplayName, Alias, PrimarySmtpAddress, RecipientTypeDetails,
        @{ Name = 'EmailAddresses'; Expression = { @($_.EmailAddresses | Where-Object { $_ -like 'smtp:*' }) } }
}
catch {
    throw "Mailbox lookup failed: $($_.Exception.Message)"
}