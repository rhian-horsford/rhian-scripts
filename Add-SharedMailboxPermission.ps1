#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

class TrusteeRow {
    [string]$Entry
    [string]$Status
    [object[]]$Options
    [bool]$HasOptions
    [object]$Selected

    TrusteeRow([string]$entry, [string]$status, [object[]]$options) {
        $this.Entry = $entry
        $this.Status = $status
        $this.Options = $options
        $this.HasOptions = $options.Count -gt 0
        if ($options.Count -eq 1) {
            $this.Selected = $options[0]
        }
    }
}

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Shared Mailbox Permission Manager"
        Height="860" Width="920" MinHeight="640" MinWidth="720"
        WindowStartupLocation="CenterScreen" Background="#F5F7FA">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="MinWidth" Value="110"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="BorderBrush" Value="#AAB4C3"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="0,0,24,0"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*" MinHeight="200"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,12">
        <Grid Margin="0,0,0,14">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel VerticalAlignment="Center">
                <TextBlock Text="Shared mailbox permissions" FontSize="20" FontWeight="SemiBold" Foreground="#172B4D"/>
                <TextBlock Margin="0,4,0,0" Foreground="#53657D" TextWrapping="Wrap"
                           Text="Grant Full Access and/or Send As permissions on a shared mailbox to one or more users."/>
            </StackPanel>
            <TextBlock x:Name="ConnectionStatusText" Grid.Column="1" Text="Not connected" VerticalAlignment="Center" Margin="12,0,12,0" Foreground="#B42318"/>
            <Button x:Name="ConnectButton" Grid.Column="2" Content="Connect" VerticalAlignment="Center" Background="#0B63CE" Foreground="White"/>
        </Grid>

        <StackPanel Margin="0,0,0,12">
            <TextBlock Text="Shared mailbox email address" FontWeight="SemiBold" Foreground="#172B4D" Margin="0,0,0,6"/>
            <TextBox x:Name="MailboxTextBox" ToolTip="Exact SMTP address or UPN of the shared mailbox."/>
        </StackPanel>

        <StackPanel Margin="0,0,0,12">
            <TextBlock Text="Users to add: names or email addresses (one per line, or separated by semicolons)" FontWeight="SemiBold" Foreground="#172B4D" Margin="0,0,0,6"/>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="UsersTextBox" Height="70" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                         ToolTip="Full or partial names (e.g. 'Jane Smith', 'Smith, Jane') or exact email addresses. Wildcards are not allowed."/>
                <Button x:Name="CheckUsersButton" Grid.Column="1" Content="Check users" VerticalAlignment="Top" Margin="10,0,0,0"
                        IsEnabled="False" Background="#E9EEF5" Foreground="#172B4D"/>
            </Grid>
        </StackPanel>

        <Border Background="White" BorderBrush="#DCE2EA" BorderThickness="1" CornerRadius="6" Padding="10">
            <StackPanel>
                <TextBlock Text="Permissions" FontWeight="SemiBold" Foreground="#172B4D" Margin="0,0,0,8"/>
                <StackPanel Orientation="Horizontal">
                    <CheckBox x:Name="FullAccessCheckBox" Content="Full Access" IsChecked="True"/>
                    <CheckBox x:Name="SendAsCheckBox" Content="Send As"/>
                    <CheckBox x:Name="AutoMappingCheckBox" Content="Auto-map mailbox in Outlook (Full Access only)" IsChecked="True"/>
                </StackPanel>
                <TextBlock x:Name="ValidationText" Foreground="#B42318" Margin="0,8,0,0" TextWrapping="Wrap" Visibility="Collapsed"/>
            </StackPanel>
        </Border>
        </StackPanel>

        <Border Grid.Row="1" Background="White" BorderBrush="#DCE2EA" BorderThickness="1" CornerRadius="6" Padding="14">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <TextBlock Text="Status" FontWeight="SemiBold" Foreground="#172B4D"/>
                <TextBlock x:Name="OperationStatusText" Grid.Row="1" Text="Ready." Margin="0,7,0,10" TextWrapping="Wrap" Foreground="#53657D"/>
                <ProgressBar x:Name="OperationProgressBar" Grid.Row="2" Height="8" Minimum="0" Maximum="100" Value="0" Margin="0,0,0,10"/>
                <DataGrid x:Name="ReviewGrid" Grid.Row="3" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column"
                          CanUserAddRows="False" GridLinesVisibility="Horizontal" BorderBrush="#DCE2EA" MinRowHeight="32"
                          VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="Entry" Binding="{Binding Entry}" Width="2*"/>
                        <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="2*"/>
                        <DataGridTemplateColumn Header="Match" Width="3*" MinWidth="260">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <ComboBox ItemsSource="{Binding Options}" DisplayMemberPath="Label" IsEnabled="{Binding HasOptions}" Margin="2"
                                              SelectedItem="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>
                    </DataGrid.Columns>
                </DataGrid>
                <DataGrid x:Name="ResultsGrid" Grid.Row="3" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column"
                          CanUserAddRows="False" GridLinesVisibility="Horizontal" BorderBrush="#DCE2EA" MinRowHeight="28" Visibility="Collapsed"
                          VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="User" Binding="{Binding User}" Width="2*"/>
                        <DataGridTextColumn Header="Permission" Binding="{Binding Permission}" Width="*"/>
                        <DataGridTextColumn Header="Result" Binding="{Binding Result}" Width="*"/>
                        <DataGridTextColumn Header="Details" Binding="{Binding Details}" Width="3*">
                            <DataGridTextColumn.ElementStyle>
                                <Style TargetType="TextBlock">
                                    <Setter Property="TextWrapping" Value="Wrap"/>
                                </Style>
                            </DataGridTextColumn.ElementStyle>
                        </DataGridTextColumn>
                    </DataGrid.Columns>
                </DataGrid>
            </Grid>
        </Border>

        <Grid Grid.Row="2" Margin="0,14,0,0">
            <TextBlock VerticalAlignment="Center" Foreground="#66758A" Text="Changes are applied to Exchange Online immediately."/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="CloseButton" Content="Close" Margin="0,0,8,0" Background="#E9EEF5" Foreground="#172B4D"/>
                <Button x:Name="AddButton" Content="Add permissions" IsEnabled="False" Background="#157347" Foreground="White"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$connectButton = $window.FindName('ConnectButton')
$connectionStatusText = $window.FindName('ConnectionStatusText')
$mailboxTextBox = $window.FindName('MailboxTextBox')
$usersTextBox = $window.FindName('UsersTextBox')
$checkUsersButton = $window.FindName('CheckUsersButton')
$reviewGrid = $window.FindName('ReviewGrid')
$fullAccessCheckBox = $window.FindName('FullAccessCheckBox')
$sendAsCheckBox = $window.FindName('SendAsCheckBox')
$autoMappingCheckBox = $window.FindName('AutoMappingCheckBox')
$validationText = $window.FindName('ValidationText')
$operationStatusText = $window.FindName('OperationStatusText')
$operationProgressBar = $window.FindName('OperationProgressBar')
$resultsGrid = $window.FindName('ResultsGrid')
$addButton = $window.FindName('AddButton')
$closeButton = $window.FindName('CloseButton')

$workArea = [System.Windows.SystemParameters]::WorkArea
$window.Height = [Math]::Min($window.Height, $workArea.Height)
$window.Width = [Math]::Min($window.Width, $workArea.Width)

$state = [hashtable]::Synchronized(@{
    Busy              = $false
    Connected         = $false
    Done              = $false
    CompletionHandled = $false
    Status            = 'Ready.'
    Percent           = 0
    Results           = $null
    Error             = $null
    Operation         = ''
    MsalApp           = $null
    MsalAccount       = $null
    AccessToken       = $null
    TokenExpiresOn    = [DateTimeOffset]::MinValue
    UserPrincipalName = ''
    OwnerHandle       = [IntPtr]::Zero
    NativeLibrary     = [IntPtr]::Zero
})

$runspace = [runspacefactory]::CreateRunspace()
$runspace.ApartmentState = 'STA'
$runspace.ThreadOptions = 'ReuseThread'
$runspace.Open()
$activePowerShell = $null
$asyncResult = $null

$identifierPattern = '^[^\s@*?,;]+@[^\s@*?,;]+\.[^\s@*?,;]+$'

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

# Commas are not separators so 'Smith, Jane' stays one entry.
function Get-UserEntries {
    @($usersTextBox.Text -split '[\r\n;]+' |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ } |
        Sort-Object -Unique)
}

function Set-InputsEnabled {
    param([bool]$Enabled)
    $connectButton.IsEnabled = $Enabled
    $mailboxTextBox.IsEnabled = $Enabled
    $usersTextBox.IsEnabled = $Enabled
    $fullAccessCheckBox.IsEnabled = $Enabled
    $sendAsCheckBox.IsEnabled = $Enabled
    $autoMappingCheckBox.IsEnabled = $Enabled -and [bool]$fullAccessCheckBox.IsChecked
    $reviewGrid.IsEnabled = $Enabled
}

function Get-ReviewRows {
    if ($reviewGrid.ItemsSource) { @($reviewGrid.ItemsSource) } else { @() }
}

function Update-ActionButtons {
    $ready = $state.Connected -and -not $state.Busy
    $checkUsersButton.IsEnabled = $ready -and -not [string]::IsNullOrWhiteSpace($usersTextBox.Text)

    $rows = Get-ReviewRows
    $allResolved = $rows.Count -gt 0 -and @($rows | Where-Object { $null -eq $_.Selected }).Count -eq 0
    $addButton.IsEnabled = $ready -and $allResolved -and
        -not [string]::IsNullOrWhiteSpace($mailboxTextBox.Text) -and
        ([bool]$fullAccessCheckBox.IsChecked -or [bool]$sendAsCheckBox.IsChecked)
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
    $state.CompletionHandled = $false
    $state.Error = $null
    $state.Results = $null
    $state.Operation = $Operation
    $state.Status = switch ($Operation) {
        'Connect' { 'Connecting to Exchange Online...' }
        'Check users' { 'Looking up users...' }
        default { 'Validating shared mailbox...' }
    }
    $state.Percent = 0

    Set-InputsEnabled -Enabled $false
    $addButton.IsEnabled = $false
    $checkUsersButton.IsEnabled = $false
    $operationProgressBar.IsIndeterminate = ($Operation -eq 'Connect')
    if ($Operation -eq 'Connect') {
        $state.Connected = $false
        $connectionStatusText.Text = 'Connecting...'
        $connectionStatusText.Foreground = '#9A6700'
    }

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    $null = $ps.AddScript($Script).AddArgument($state)
    foreach ($argument in $Arguments) {
        $null = $ps.AddArgument($argument)
    }
    $script:activePowerShell = $ps
    $script:asyncResult = $ps.BeginInvoke()
}

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(150)
$timer.Add_Tick({
    $operationStatusText.Text = [string]$state.Status
    if (-not $operationProgressBar.IsIndeterminate) {
        $operationProgressBar.Value = [double]$state.Percent
    }

    if ($state.Busy -and $state.Done -and -not $state.CompletionHandled) {
        $state.CompletionHandled = $true
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
        Set-InputsEnabled -Enabled $true

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
            $connectionStatusText.Text = "Connected as $($state.UserPrincipalName)"
            $connectionStatusText.Foreground = '#157347'
            $connectButton.Content = 'Reconnect'
            $operationStatusText.Text = 'Connected. Enter a shared mailbox and users, then select Check users.'
        }
        elseif ($completedOperation -eq 'Check users') {
            $rows = @(foreach ($r in @($state.Results)) {
                [TrusteeRow]::new([string]$r.Entry, [string]$r.Status, [object[]]@($r.Options))
            })
            $reviewGrid.ItemsSource = $rows
            $resultsGrid.Visibility = 'Collapsed'
            $reviewGrid.Visibility = 'Visible'
            $operationProgressBar.Value = 100
            $unresolved = @($rows | Where-Object { $null -eq $_.Selected }).Count
            $operationStatusText.Text = if ($unresolved -eq 0) {
                $connectionStatusText.Text = if ([string]$state.Error -eq 'Sign-in cancelled.') { 'Sign-in cancelled' } else { 'Not connected' }
            }
            else {
            if ([string]$state.Error -ne 'Sign-in cancelled.') {
                Show-ErrorDialog -Title "$completedOperation failed" -Message ([string]$state.Error)
            }
            }
        }
        else {
            $results = @($state.Results)
            $reviewGrid.ItemsSource = $null
            $reviewGrid.Visibility = 'Collapsed'
            $resultsGrid.Visibility = 'Visible'
            $resultsGrid.ItemsSource = $results
            $operationProgressBar.Value = 100
            $failed = @($results | Where-Object { $_.Result -eq 'Failed' }).Count
            $operationStatusText.Text = "Completed: $($results.Count - $failed) succeeded, $failed failed."
            $icon = if ($failed -gt 0) { [System.Windows.MessageBoxImage]::Warning } else { [System.Windows.MessageBoxImage]::Information }
            [System.Windows.MessageBox]::Show(
                $window,
                "Permissions processed.`n`nSucceeded: $($results.Count - $failed)`nFailed: $failed`n`nSee the results table for details.",
                'Operation complete',
                [System.Windows.MessageBoxButton]::OK,
                $icon
            ) | Out-Null
        }

        Update-ActionButtons
    }
})
$timer.Start()

$connectScript = {
    param($State, $OwnerHandle)
    try {
        $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
            Sort-Object Version -Descending |
            Select-Object -First 1
        if (-not $module) {
            throw "The ExchangeOnlineManagement module is not installed. Close this tool and run 'Install-Module ExchangeOnlineManagement -Scope CurrentUser' from a trusted PowerShell Gallery session."
        }

        Import-Module $module.Path -ErrorAction Stop
        $authFolder = Join-Path $module.ModuleBase 'netCore'
        $msalPath = Join-Path $authFolder 'Microsoft.Identity.Client.dll'
        $brokerPath = Join-Path $authFolder 'Microsoft.Identity.Client.Broker.dll'
        if (-not (Test-Path $msalPath) -or -not (Test-Path $brokerPath)) {
            throw 'The installed ExchangeOnlineManagement module is missing its MSAL Windows broker files. Install or update ExchangeOnlineManagement 3.7 or later.'
        }

        [void][Reflection.Assembly]::LoadFrom($msalPath)
        [void][Reflection.Assembly]::LoadFrom($brokerPath)

        $runtimeId = if ([Environment]::Is64BitProcess) { 'win-x64' } else { 'win-x86' }
        $nativeFolder = Join-Path $authFolder (Join-Path 'runtimes' (Join-Path $runtimeId 'native'))
        $nativePath = Join-Path $nativeFolder 'msalruntime.dll'
        if (-not (Test-Path $nativePath)) {
            throw "The MSAL Windows broker runtime was not found at '$nativePath'."
        }
        if ($env:PATH -notlike "*$nativeFolder*") {
            $env:PATH = "$nativeFolder;$env:PATH"
        }
        if ($State.NativeLibrary -eq [IntPtr]::Zero) {
            $State.NativeLibrary = [System.Runtime.InteropServices.NativeLibrary]::Load($nativePath)
        }

        $clientId = 'fb78d390-0c51-40cd-8e17-fdbfab77341b'
        $scopes = [string[]]@('https://outlook.office365.com/.default')
        $brokerOptions = [Microsoft.Identity.Client.BrokerOptions]::new(
            [Microsoft.Identity.Client.BrokerOptions+OperatingSystems]::Windows
        )
        $builder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($clientId)
        $builder = [Microsoft.Identity.Client.Broker.BrokerExtension]::WithBroker($builder, $brokerOptions)
        $app = $builder.WithAuthority('https://login.microsoftonline.com/organizations').Build()
        $authResult = $app.AcquireTokenInteractive($scopes).
            WithParentActivityOrWindow([IntPtr]$OwnerHandle).
            ExecuteAsync().GetAwaiter().GetResult()

        Connect-ExchangeOnline -AccessToken $authResult.AccessToken `
            -UserPrincipalName $authResult.Account.Username `
            -ShowBanner:$false -ErrorAction Stop

        $State.MsalApp = $app
        $State.MsalAccount = $authResult.Account
        $State.AccessToken = $authResult.AccessToken
        $State.TokenExpiresOn = $authResult.ExpiresOn
        $State.UserPrincipalName = $authResult.Account.Username
        $State.OwnerHandle = [IntPtr]$OwnerHandle
        $State.Status = "Connected to Exchange Online as $($authResult.Account.Username)."
    }
    catch {
        $message = $_.Exception.ToString()
        if ($message -match 'authentication_canceled|user canceled|user cancelled') {
            $State.Error = 'Sign-in cancelled.'
        }
        else {
            $State.Error = $_.Exception.Message
        }
    }
    finally {
        $State.Done = $true
    }
}

$ensureConnectionScript = {
    param($State)
    if (-not $State.MsalApp -or -not $State.MsalAccount) {
        throw 'Exchange Online sign-in is missing. Select Connect and sign in again.'
    }

    $scopes = [string[]]@('https://outlook.office365.com/.default')
    try {
        $authResult = $State.MsalApp.AcquireTokenSilent($scopes, $State.MsalAccount).
            ExecuteAsync().GetAwaiter().GetResult()
    }
    catch [Microsoft.Identity.Client.MsalUiRequiredException] {
        try {
            $authResult = $State.MsalApp.AcquireTokenInteractive($scopes).
                WithParentActivityOrWindow([IntPtr]$State.OwnerHandle).
                ExecuteAsync().GetAwaiter().GetResult()
        }
        catch {
            if ($_.Exception.ToString() -match 'authentication_canceled|user canceled|user cancelled') {
                throw 'Sign-in cancelled.'
            }
            throw
        }
    }

    if ($authResult.AccessToken -cne $State.AccessToken) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Connect-ExchangeOnline -AccessToken $authResult.AccessToken `
            -UserPrincipalName $authResult.Account.Username `
            -ShowBanner:$false -ErrorAction Stop
        $State.AccessToken = $authResult.AccessToken
        $State.TokenExpiresOn = $authResult.ExpiresOn
        $State.UserPrincipalName = $authResult.Account.Username
    }
}

$resolveScript = {
    param($State, $Entries, $EmailPattern, $EnsureConnection)
    try {
        & $EnsureConnection $State
        $rows = New-Object System.Collections.Generic.List[object]
        $RawLimit = 200
        $index = 0
        foreach ($entry in $Entries) {
            $index++
            $State.Status = "Looking up '$entry' ($index of $($Entries.Count))..."
            $status = $null
            $found = @()
            try {
                if ($entry -match $EmailPattern) {
                    $recipient = Get-EXORecipient -Identity $entry -ErrorAction Stop
                    if ($recipient.RecipientTypeDetails -eq 'UserMailbox') {
                        $found = @($recipient)
                    }
                    else {
                        $status = "Not a user mailbox ($($recipient.RecipientTypeDetails))"
                    }
                }
                else {
                    $raw = @(Get-Recipient -Anr $entry -RecipientTypeDetails UserMailbox -ResultSize $RawLimit -ErrorAction Stop)
                    $words = @($entry -split '[\s,]+' | Where-Object { $_ })
                    # Guard against the server returning unfiltered results: every typed word must prefix a name/alias/address word.
                    $found = @($raw | Where-Object {
                        $candidateWords = @("$($_.DisplayName) $($_.FirstName) $($_.LastName) $($_.Alias) $($_.PrimarySmtpAddress)" -split '[\s,.@_-]+' | Where-Object { $_ })
                        $missing = @($words | Where-Object {
                            $word = $_
                            -not ($candidateWords | Where-Object { $_.StartsWith($word, [StringComparison]::OrdinalIgnoreCase) })
                        })
                        $missing.Count -eq 0
                    })
                    if ($raw.Count -ge $RawLimit) {
                        $status = 'Too many matches - be more specific'
                        $found = @()
                    }
                }
            }
            catch {
                $status = "Lookup failed: $($_.Exception.Message)"
            }

            if ($found.Count -gt 25) {
                $status = 'Too many matches - be more specific'
                $found = @()
            }
            if (-not $status) {
                $status = if ($found.Count -eq 0) { 'Not found' }
                    elseif ($found.Count -eq 1) { 'Resolved' }
                    else { "Choose a match ($($found.Count) found)" }
            }

            $options = @($found | ForEach-Object {
                [pscustomobject]@{
                    Address = [string]$_.PrimarySmtpAddress
                    Label   = "$($_.DisplayName) <$($_.PrimarySmtpAddress)>"
                }
            })
            $rows.Add([pscustomobject]@{ Entry = $entry; Status = $status; Options = $options })
            $State.Percent = [int](($index / $Entries.Count) * 100)
        }
        $State.Results = $rows.ToArray()
    }
    catch {
        $State.Error = $_.Exception.Message
    }
    finally {
        $State.Done = $true
    }
}

$addScript = {
    param($State, $Mailbox, $Trustees, $GrantFullAccess, $GrantSendAs, $AutoMapping, $EnsureConnection)
    try {
        & $EnsureConnection $State
        $sharedMailbox = Get-EXOMailbox -Identity $Mailbox -Properties RecipientTypeDetails -ErrorAction Stop
        if ($sharedMailbox.RecipientTypeDetails -ne 'SharedMailbox') {
            throw "'$Mailbox' is a $($sharedMailbox.RecipientTypeDetails), not a shared mailbox."
        }
        $target = $sharedMailbox.PrimarySmtpAddress

        $results = New-Object System.Collections.Generic.List[object]
        $permissions = @()
        if ($GrantFullAccess) { $permissions += 'Full Access' }
        if ($GrantSendAs) { $permissions += 'Send As' }
        $total = $Trustees.Count * $permissions.Count
        $step = 0

        foreach ($trustee in $Trustees) {
            $user = $trustee.Label
            foreach ($permission in $permissions) {
                $step++
                $State.Status = "Granting $permission to $user ($step of $total)..."
                try {
                    if ($permission -eq 'Full Access') {
                        $null = Add-MailboxPermission -Identity $target -User $trustee.Address `
                            -AccessRights FullAccess -InheritanceType All -AutoMapping $AutoMapping `
                            -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings
                    }
                    else {
                        $null = Add-RecipientPermission -Identity $target -Trustee $trustee.Address `
                            -AccessRights SendAs -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings
                    }
                    $details = if ($warnings) { ($warnings | ForEach-Object { $_.Message }) -join ' ' } else { 'Granted.' }
                    $results.Add([pscustomobject]@{ User = $user; Permission = $permission; Result = 'Success'; Details = $details })
                }
                catch {
                    $results.Add([pscustomobject]@{ User = $user; Permission = $permission; Result = 'Failed'; Details = $_.Exception.Message })
                }
                $State.Percent = [int](($step / $total) * 100)
            }
        }

        $State.Results = $results.ToArray()
    }
    catch {
        $State.Error = $_.Exception.Message
    }
    finally {
        $State.Done = $true
    }
}

$connectButton.Add_Click({
    $ownerHandle = (New-Object System.Windows.Interop.WindowInteropHelper($window)).Handle
    Start-BackgroundOperation -Operation 'Connect' -Script $connectScript -Arguments @($ownerHandle)
})

$closeButton.Add_Click({
    $window.Close()
})

$fullAccessCheckBox.Add_Click({
    $autoMappingCheckBox.IsEnabled = [bool]$fullAccessCheckBox.IsChecked
    $validationText.Visibility = 'Collapsed'
    Update-ActionButtons
})

$sendAsCheckBox.Add_Click({
    $validationText.Visibility = 'Collapsed'
    Update-ActionButtons
})

$mailboxTextBox.Add_TextChanged({
    $validationText.Visibility = 'Collapsed'
    Update-ActionButtons
})

$usersTextBox.Add_TextChanged({
    $validationText.Visibility = 'Collapsed'
    $reviewGrid.ItemsSource = $null
    Update-ActionButtons
})

$reviewGrid.AddHandler(
    [System.Windows.Controls.Primitives.Selector]::SelectionChangedEvent,
    [System.Windows.Controls.SelectionChangedEventHandler]{ Update-ActionButtons }
)

$checkUsersButton.Add_Click({
    $entries = Get-UserEntries
    $invalid = @($entries | Where-Object { $_ -match '[*?]' -or $_.Length -gt 256 })
    if ($entries.Count -eq 0 -or $invalid.Count -gt 0) {
        $validationText.Text = if ($invalid.Count -gt 0) {
            "Remove wildcards (* or ?) or shorten these entries: $($invalid -join '; ')"
        }
        else {
            'Enter at least one name or email address.'
        }
        $validationText.Visibility = 'Visible'
        return
    }

    $reviewGrid.ItemsSource = $null
    Start-BackgroundOperation `
        -Operation 'Check users' `
        -Script $resolveScript `
        -Arguments @($entries, $identifierPattern, $ensureConnectionScript)
})

$addButton.Add_Click({
    $mailbox = $mailboxTextBox.Text.Trim()
    $rows = Get-ReviewRows

    $errors = @()
    if ($mailbox -notmatch $identifierPattern) {
        $errors += 'Enter a valid shared mailbox email address.'
    }
    if ($rows.Count -eq 0 -or @($rows | Where-Object { $null -eq $_.Selected }).Count -gt 0) {
        $errors += 'Every user must be resolved to a single match. Select Check users and choose any ambiguous matches.'
    }
    if (-not ($fullAccessCheckBox.IsChecked -or $sendAsCheckBox.IsChecked)) {
        $errors += 'Select at least one permission.'
    }
    if ($errors.Count -gt 0) {
        $validationText.Text = $errors -join "`n"
        $validationText.Visibility = 'Visible'
        return
    }

    $permissionText = @(
        if ($fullAccessCheckBox.IsChecked) { 'Full Access' }
        if ($sendAsCheckBox.IsChecked) { 'Send As' }
    ) -join ' and '
    $trustees = @($rows | ForEach-Object { $_.Selected } | Group-Object Address | ForEach-Object { $_.Group[0] })
    $confirm = [System.Windows.MessageBox]::Show(
        $window,
        "Grant $permissionText on $mailbox to $($trustees.Count) user(s)?`n`n$(($trustees | ForEach-Object { $_.Label }) -join "`n")",
        'Confirm changes',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
        return
    }

    $resultsGrid.ItemsSource = $null
    Start-BackgroundOperation `
        -Operation 'Add permissions' `
        -Script $addScript `
        -Arguments @($mailbox, $trustees, [bool]$fullAccessCheckBox.IsChecked, [bool]$sendAsCheckBox.IsChecked, [bool]$autoMappingCheckBox.IsChecked, $ensureConnectionScript)
})

$window.Add_Closing({
    param($windowSender, $windowEventArgs)
    if ($state.Busy) {
        $choice = [System.Windows.MessageBox]::Show(
            $window,
            'An operation is still running. Stop it and close the application?',
            'Operation in progress',
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        if ($choice -ne [System.Windows.MessageBoxResult]::Yes) {
            $windowEventArgs.Cancel = $true
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
