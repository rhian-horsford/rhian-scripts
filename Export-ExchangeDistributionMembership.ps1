#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$modulePath = Join-Path $PSScriptRoot 'ExchangeDistributionMembership.psm1'

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Exchange Distribution Membership Exporter"
        Height="520" Width="760" MinHeight="500" MinWidth="700"
        WindowStartupLocation="CenterScreen" Background="#F5F7FA">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="MinWidth" Value="110"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="BorderBrush" Value="#AAB4C3"/>
        </Style>
    </Window.Resources>
    <Grid Margin="28">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,24">
            <TextBlock Text="Distribution membership exporter" FontSize="24" FontWeight="SemiBold" Foreground="#172B4D"/>
            <TextBlock Margin="0,6,0,0" Foreground="#53657D" TextWrapping="Wrap"
                       Text="Export a user's direct static distribution-group memberships and current dynamic recipient-filter matches."/>
        </StackPanel>

        <Border Grid.Row="1" Background="White" BorderBrush="#DCE2EA" BorderThickness="1" CornerRadius="6" Padding="18" Margin="0,0,0,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                    <TextBlock Text="Exchange Online connection" FontWeight="SemiBold" Foreground="#172B4D"/>
                    <TextBlock x:Name="ConnectionStatusText" Text="Not connected" Margin="0,5,0,0" Foreground="#B42318"/>
                </StackPanel>
                <Button x:Name="ConnectButton" Grid.Column="1" Content="Connect" Background="#0B63CE" Foreground="White"/>
            </Grid>
        </Border>

        <StackPanel Grid.Row="2" Margin="0,0,0,14">
            <TextBlock Text="User email address or UPN" FontWeight="SemiBold" Foreground="#172B4D" Margin="0,0,0,6"/>
            <TextBox x:Name="IdentifierTextBox" ToolTip="Enter an exact SMTP address or user principal name. Wildcards are not allowed."/>
            <TextBlock x:Name="ValidationText" Foreground="#B42318" Margin="2,5,0,0" Visibility="Collapsed"/>
        </StackPanel>

        <StackPanel Grid.Row="3" Margin="0,0,0,14">
            <TextBlock Text="Recipient search scope" FontWeight="SemiBold" Foreground="#172B4D" Margin="0,0,0,6"/>
            <ComboBox x:Name="LookupModeComboBox" SelectedValuePath="Tag" SelectedValue="Active">
                <ComboBoxItem Content="Active recipients" Tag="Active"/>
                <ComboBoxItem Content="Soft-deleted recipients" Tag="SoftDeleted"/>
                <ComboBoxItem Content="Inactive mailboxes" Tag="InactiveMailbox"/>
                <ComboBoxItem Content="Soft-deleted recipients and inactive mailboxes" Tag="Both"/>
            </ComboBox>
        </StackPanel>

        <StackPanel Grid.Row="4" Margin="0,0,0,16">
            <TextBlock Text="CSV destination" FontWeight="SemiBold" Foreground="#172B4D" Margin="0,0,0,6"/>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="DestinationTextBox" IsReadOnly="True"/>
                <Button x:Name="BrowseButton" Grid.Column="1" Content="Browse..." Margin="10,0,0,0" Background="#E9EEF5" Foreground="#172B4D"/>
            </Grid>
        </StackPanel>

        <Border Grid.Row="5" Background="White" BorderBrush="#DCE2EA" BorderThickness="1" CornerRadius="6" Padding="18">
            <StackPanel>
                <TextBlock Text="Status" FontWeight="SemiBold" Foreground="#172B4D"/>
                <TextBlock x:Name="OperationStatusText" Text="Ready." Margin="0,7,0,10" TextWrapping="Wrap" Foreground="#53657D"/>
                <ProgressBar x:Name="OperationProgressBar" Height="8" Minimum="0" Maximum="100" Value="0"/>
            </StackPanel>
        </Border>

        <Grid Grid.Row="6" Margin="0,20,0,0">
            <TextBlock VerticalAlignment="Center" Foreground="#66758A" Text="No changes are made in Exchange Online."/>
            <Button x:Name="ExportButton" Content="Export CSV" HorizontalAlignment="Right" IsEnabled="False" Background="#157347" Foreground="White" Margin="0"/>
        </Grid>
    </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$connectButton = $window.FindName('ConnectButton')
$connectionStatusText = $window.FindName('ConnectionStatusText')
$identifierTextBox = $window.FindName('IdentifierTextBox')
$lookupModeComboBox = $window.FindName('LookupModeComboBox')
$validationText = $window.FindName('ValidationText')
$destinationTextBox = $window.FindName('DestinationTextBox')
$browseButton = $window.FindName('BrowseButton')
$operationStatusText = $window.FindName('OperationStatusText')
$operationProgressBar = $window.FindName('OperationProgressBar')
$exportButton = $window.FindName('ExportButton')

$state = [hashtable]::Synchronized(@{
    Busy      = $false
    Connected = $false
    Done      = $false
    Status    = 'Ready.'
    Percent   = 0
    Result    = $null
    Error     = $null
    Operation = ''
})

$runspace = [runspacefactory]::CreateRunspace()
$runspace.ApartmentState = 'STA'
$runspace.ThreadOptions = 'ReuseThread'
$runspace.Open()
$activePowerShell = $null
$asyncResult = $null

function Show-ErrorDialog {
    param(
        [string]$Title,
        [string]$Message
    )

    [System.Windows.MessageBox]::Show(
        $window,
        $Message,
        $Title,
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    ) | Out-Null
}

function Update-ExportEnabledState {
    $exportButton.IsEnabled = $state.Connected -and
        -not $state.Busy -and
        -not [string]::IsNullOrWhiteSpace($identifierTextBox.Text) -and
        -not [string]::IsNullOrWhiteSpace($destinationTextBox.Text)
}

function Start-BackgroundOperation {
    param(
        [string]$Operation,
        [scriptblock]$Script,
        [object[]]$Arguments
    )

    if ($state.Busy) {
        return
    }

    $state.Busy = $true
    $state.Done = $false
    $state.Error = $null
    $state.Result = $null
    $state.Operation = $Operation
    $state.Status = if ($Operation -eq 'Connect') { 'Connecting to Exchange Online...' } else { 'Starting export...' }
    $state.Percent = 0

    $connectButton.IsEnabled = $false
    $browseButton.IsEnabled = $false
    $exportButton.IsEnabled = $false
    $identifierTextBox.IsEnabled = $false
    $lookupModeComboBox.IsEnabled = $false
    $operationProgressBar.IsIndeterminate = ($Operation -eq 'Connect')
    if ($Operation -eq 'Connect') {
        $state.Connected = $false
        $connectionStatusText.Text = 'Connecting...'
        $connectionStatusText.Foreground = '#9A6700'
    }

    $activePowerShell = [powershell]::Create()
    $activePowerShell.Runspace = $runspace
    $null = $activePowerShell.AddScript($Script).AddArgument($state).AddArgument($modulePath)
    foreach ($argument in $Arguments) {
        $null = $activePowerShell.AddArgument($argument)
    }
    $script:activePowerShell = $activePowerShell
    $script:asyncResult = $activePowerShell.BeginInvoke()
}

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(150)
$timer.Add_Tick({
    $operationStatusText.Text = [string]$state.Status
    if (-not $operationProgressBar.IsIndeterminate) {
        $operationProgressBar.Value = [double]$state.Percent
    }

    if ($state.Busy -and $state.Done) {
        try {
            if ($script:activePowerShell -and $script:asyncResult) {
                $null = $script:activePowerShell.EndInvoke($script:asyncResult)
            }
        }
        catch {
            if (-not $state.Error) {
                $state.Error = $_.Exception.Message
            }
        }
        finally {
            if ($script:activePowerShell) {
                $script:activePowerShell.Dispose()
            }
            $script:activePowerShell = $null
            $script:asyncResult = $null
        }

        $completedOperation = $state.Operation
        $state.Busy = $false
        $operationProgressBar.IsIndeterminate = $false
        $connectButton.IsEnabled = $true
        $browseButton.IsEnabled = $true
        $identifierTextBox.IsEnabled = $true
        $lookupModeComboBox.IsEnabled = $true

        if ($state.Error) {
            $operationProgressBar.Value = 0
            $operationStatusText.Text = "$completedOperation failed."
            if ($completedOperation -eq 'Connect') {
                $connectionStatusText.Text = 'Not connected'
                $connectionStatusText.Foreground = '#B42318'
            }
            Show-ErrorDialog -Title "$completedOperation failed" -Message ([string]$state.Error)
        }
        elseif ($completedOperation -eq 'Connect') {
            $state.Connected = $true
            $connectionStatusText.Text = 'Connected to Exchange Online'
            $connectionStatusText.Foreground = '#157347'
            $connectButton.Content = 'Reconnect'
            $operationStatusText.Text = 'Connected. Select a destination and export when ready.'
        }
        else {
            $summary = $state.Result
            $operationProgressBar.Value = 100
            $operationStatusText.Text = "Exported $($summary.TotalGroupCount) membership record(s) to $($summary.Path)"
            $message = @"
Export complete.

Recipient: $($summary.RecipientDisplayName) <$($summary.RecipientAddress)>
Search scope: $($summary.LookupMode)
Static direct memberships: $($summary.StaticGroupCount)
Dynamic filter matches: $($summary.DynamicGroupCount) ($($summary.DynamicMembershipStatus))
Total records: $($summary.TotalGroupCount)

CSV: $($summary.Path)
"@
            [System.Windows.MessageBox]::Show(
                $window,
                $message,
                'Export complete',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null
        }

        Update-ExportEnabledState
    }
})
$timer.Start()

$connectScript = {
    param($State, $ModulePath)
    try {
        $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
            Sort-Object Version -Descending |
            Select-Object -First 1
        if (-not $module) {
            throw "The ExchangeOnlineManagement module is not installed. Close this tool and run 'Install-Module ExchangeOnlineManagement -Scope CurrentUser' from a trusted PowerShell Gallery session."
        }

        Import-Module $module.Path -ErrorAction Stop
        Import-Module $ModulePath -Force -ErrorAction Stop
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop

        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $connections = @(Get-ConnectionInformation -ErrorAction Stop)
            if ($connections.Count -eq 0) {
                throw 'Exchange Online did not report an active connection after sign-in.'
            }
        }
        $State.Status = 'Connected to Exchange Online.'
    }
    catch {
        $State.Error = $_.Exception.Message
    }
    finally {
        $State.Done = $true
    }
}

$exportScript = {
    param($State, $ModulePath, $Identifier, $Destination, $LookupMode)
    try {
        Import-Module $ModulePath -Force -ErrorAction Stop
        $statusAction = {
            param($Current, $Total, $Message)
            $State.Percent = if ($Total -gt 0) { [int](($Current / $Total) * 100) } else { 0 }
            $State.Status = $Message
        }
        $State.Result = Export-UserDistributionGroupMembership `
            -Identifier $Identifier `
            -Path $Destination `
            -LookupMode $LookupMode `
            -StatusAction $statusAction `
            -ErrorAction Stop
    }
    catch {
        $State.Error = $_.Exception.Message
    }
    finally {
        $State.Done = $true
    }
}

$connectButton.Add_Click({
    Start-BackgroundOperation -Operation 'Connect' -Script $connectScript -Arguments @()
})

$browseButton.Add_Click({
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Title = 'Choose CSV destination'
    $dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $dialog.DefaultExt = '.csv'
    $dialog.AddExtension = $true
    $dialog.OverwritePrompt = $true
    $dialog.FileName = if ([string]::IsNullOrWhiteSpace($identifierTextBox.Text)) {
        'distribution-memberships.csv'
    }
    else {
        "$($identifierTextBox.Text.Trim().Replace('@', '_at_'))-distribution-memberships.csv"
    }

    if ($dialog.ShowDialog($window)) {
        $destinationTextBox.Text = $dialog.FileName
        Update-ExportEnabledState
    }
})

$identifierTextBox.Add_TextChanged({
    $validationText.Visibility = 'Collapsed'
    Update-ExportEnabledState
})

$exportButton.Add_Click({
    Import-Module $modulePath -Force
    if (-not (Test-RecipientIdentifier -Identifier $identifierTextBox.Text)) {
        $validationText.Text = 'Enter a valid email address or UPN. Wildcards are not allowed.'
        $validationText.Visibility = 'Visible'
        $identifierTextBox.Focus()
        return
    }

    if ([string]::IsNullOrWhiteSpace($destinationTextBox.Text)) {
        Show-ErrorDialog -Title 'Destination required' -Message 'Choose where to save the CSV file.'
        return
    }

    Start-BackgroundOperation `
        -Operation 'Export' `
        -Script $exportScript `
        -Arguments @($identifierTextBox.Text.Trim(), $destinationTextBox.Text, [string]$lookupModeComboBox.SelectedValue)
})

$window.Add_Closing({
    param($sender, $eventArgs)
    if ($state.Busy) {
        $choice = [System.Windows.MessageBox]::Show(
            $window,
            'An operation is still running. Stop it and close the application?',
            'Operation in progress',
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        if ($choice -ne [System.Windows.MessageBoxResult]::Yes) {
            $eventArgs.Cancel = $true
            return
        }
        if ($script:activePowerShell) {
            $script:activePowerShell.Stop()
        }
    }
})

$window.Add_Closed({
    $timer.Stop()
    if ($script:activePowerShell) {
        $script:activePowerShell.Dispose()
    }
    if ($runspace) {
        $runspace.Close()
        $runspace.Dispose()
    }
})

$null = $window.ShowDialog()
