#requires -Version 7.6

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Expanded Distribution Group Exporter"
        Height="820" Width="1120" MinHeight="620" MinWidth="800"
        WindowStartupLocation="CenterScreen" Background="#F5F7FA">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="MinWidth" Value="110"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="BorderBrush" Value="#AAB4C3"/>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*" MinHeight="240"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="0" Margin="0,0,0,14">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel VerticalAlignment="Center">
                <TextBlock Text="Expanded distribution group members" FontSize="22" FontWeight="SemiBold" Foreground="#172B4D"/>
                <TextBlock Text="Expand nested and dynamic groups, deduplicate recipients, and export the source details." Margin="0,4,0,0" Foreground="#53657D"/>
            </StackPanel>
            <TextBlock x:Name="ConnectionStatusText" Grid.Column="1" Text="Not connected" VerticalAlignment="Center" Margin="12,0,12,0" Foreground="#B42318"/>
            <Button x:Name="ConnectButton" Grid.Column="2" Content="Connect" VerticalAlignment="Center" Background="#0B63CE" Foreground="White"/>
        </Grid>

        <Border Grid.Row="1" Background="White" BorderBrush="#DCE2EA" BorderThickness="1" CornerRadius="6" Padding="14" Margin="0,0,0,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Margin="0,0,14,0">
                    <TextBlock Text="Parent distribution group" FontWeight="SemiBold" Foreground="#172B4D" Margin="0,0,0,6"/>
                    <TextBox x:Name="ParentGroupTextBox" ToolTip="Enter the group's exact email address, alias, or identity."/>
                </StackPanel>
                <StackPanel Grid.Column="1" Margin="0,0,14,0">
                    <TextBlock Text="CSV export folder" FontWeight="SemiBold" Foreground="#172B4D" Margin="0,0,0,6"/>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="ExportFolderTextBox" IsReadOnly="True"/>
                        <Button x:Name="BrowseButton" Grid.Column="1" Content="Browse..." Margin="8,0,0,0" Background="#E9EEF5" Foreground="#172B4D"/>
                    </Grid>
                </StackPanel>
                <Button x:Name="ExportButton" Grid.Column="2" Content="Expand and export" VerticalAlignment="Bottom" IsEnabled="False" Background="#157347" Foreground="White"/>
            </Grid>
        </Border>

        <Border Grid.Row="2" Background="White" BorderBrush="#DCE2EA" BorderThickness="1" CornerRadius="6" Padding="14">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <TextBlock Text="Export results" FontWeight="SemiBold" Foreground="#172B4D"/>
                <Grid Grid.Row="1" Margin="0,6,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="220"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock x:Name="OperationStatusText" Text="Connect to Exchange Online to begin." TextWrapping="Wrap" Foreground="#53657D" VerticalAlignment="Center"/>
                    <ProgressBar x:Name="OperationProgressBar" Grid.Column="1" Height="8" Minimum="0" Maximum="100" Value="0" VerticalAlignment="Center"/>
                </Grid>
                <TabControl Grid.Row="2">
                    <TabItem Header="Recipients">
                        <DataGrid x:Name="ResultsGrid" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column"
                                  CanUserAddRows="False" GridLinesVisibility="Horizontal" BorderBrush="#DCE2EA" MinRowHeight="28"
                                  VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Display name" Binding="{Binding DisplayName}" Width="1.4*"/>
                                <DataGridTextColumn Header="Email" Binding="{Binding PrimarySmtpAddress}" Width="1.5*"/>
                                <DataGridTextColumn Header="Recipient type" Binding="{Binding RecipientTypeDetails}" Width="1.2*"/>
                                <DataGridTextColumn Header="Source groups" Binding="{Binding SourceGroups}" Width="2*"/>
                                <DataGridTextColumn Header="Membership types" Binding="{Binding MembershipTypes}" Width="1.2*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </TabItem>
                    <TabItem Header="Warnings">
                        <ListBox x:Name="WarningsList" BorderThickness="0" Padding="8"/>
                    </TabItem>
                </TabControl>
            </Grid>
        </Border>

        <Grid Grid.Row="3" Margin="0,14,0,0">
            <TextBlock VerticalAlignment="Center" Foreground="#66758A" Text="This tool reads Exchange Online and writes a local CSV."/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="CloseButton" Content="Close" Margin="0,0,8,0" Background="#E9EEF5" Foreground="#172B4D"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
'@

$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader($xaml)))
$connectButton = $window.FindName('ConnectButton')
$connectionStatusText = $window.FindName('ConnectionStatusText')
$parentGroupTextBox = $window.FindName('ParentGroupTextBox')
$exportFolderTextBox = $window.FindName('ExportFolderTextBox')
$browseButton = $window.FindName('BrowseButton')
$exportButton = $window.FindName('ExportButton')
$operationStatusText = $window.FindName('OperationStatusText')
$operationProgressBar = $window.FindName('OperationProgressBar')
$resultsGrid = $window.FindName('ResultsGrid')
$warningsList = $window.FindName('WarningsList')
$closeButton = $window.FindName('CloseButton')
$exportFolderTextBox.Text = Join-Path $HOME 'OneDrive - Slalom\Desktop'

$workArea = [System.Windows.SystemParameters]::WorkArea
$window.Height = [Math]::Min($window.Height, $workArea.Height)
$window.Width = [Math]::Min($window.Width, $workArea.Width)

$state = [hashtable]::Synchronized(@{
    Busy              = $false
    Connected         = $false
    Done              = $false
    CompletionHandled = $false
    Status            = 'Connect to Exchange Online to begin.'
    Percent           = 0
    Results           = $null
    Warnings          = $null
    Error             = $null
    Operation         = ''
    OutputPath        = ''
    MsalApp           = $null
    MsalAccount       = $null
    AccessToken       = $null
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

function Set-InputsEnabled {
    param([bool]$Enabled)
    $connectButton.IsEnabled = $Enabled
    $parentGroupTextBox.IsEnabled = $Enabled
    $browseButton.IsEnabled = $Enabled
    $exportButton.IsEnabled = $Enabled -and $state.Connected -and
        -not [string]::IsNullOrWhiteSpace($parentGroupTextBox.Text) -and
        -not [string]::IsNullOrWhiteSpace($exportFolderTextBox.Text)
}

function Update-ExportEnabledState {
    $exportButton.IsEnabled = $state.Connected -and -not $state.Busy -and
        -not [string]::IsNullOrWhiteSpace($parentGroupTextBox.Text) -and
        -not [string]::IsNullOrWhiteSpace($exportFolderTextBox.Text)
}

function Start-BackgroundOperation {
    param(
        [string]$Operation,
        [scriptblock]$Script,
        [object[]]$Arguments
    )

    if ($state.Busy) { return }

    $state.Busy = $true
    $state.Done = $false
    $state.CompletionHandled = $false
    $state.Error = $null
    $state.Results = $null
    $state.Operation = $Operation
    $state.Status = if ($Operation -eq 'Connect') { 'Opening Microsoft sign-in...' } else { 'Resolving group and expanding members...' }
    $state.Percent = 0
    Set-InputsEnabled -Enabled $false
    $connectButton.IsEnabled = $false
    $operationProgressBar.IsIndeterminate = $true
    if ($Operation -eq 'Connect') {
        $state.Connected = $false
        $connectionStatusText.Text = 'Signing in...'
        $connectionStatusText.Foreground = '#9A6700'
    }

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    $null = $ps.AddScript($Script).AddArgument($state)
    foreach ($argument in $Arguments) { $null = $ps.AddArgument($argument) }
    $script:activePowerShell = $ps
    $script:asyncResult = $ps.BeginInvoke()
}

$connectScript = {
    param($State, $OwnerHandle)
    try {
        $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
            Sort-Object Version -Descending |
            Select-Object -First 1
        if (-not $module) {
            throw "ExchangeOnlineManagement is not installed. Install it with 'Install-Module ExchangeOnlineManagement -Scope CurrentUser'."
        }

        Import-Module $module.Path -ErrorAction Stop
        $authFolder = Join-Path $module.ModuleBase 'netCore'
        $msalPath = Join-Path $authFolder 'Microsoft.Identity.Client.dll'
        $brokerPath = Join-Path $authFolder 'Microsoft.Identity.Client.Broker.dll'
        if (-not (Test-Path $msalPath) -or -not (Test-Path $brokerPath)) {
            throw 'The installed ExchangeOnlineManagement module is missing its MSAL Windows broker files. Install version 3.7 or later.'
        }

        [void][Reflection.Assembly]::LoadFrom($msalPath)
        [void][Reflection.Assembly]::LoadFrom($brokerPath)
        $runtimeId = if ([Environment]::Is64BitProcess) { 'win-x64' } else { 'win-x86' }
        $nativeFolder = Join-Path $authFolder (Join-Path 'runtimes' (Join-Path $runtimeId 'native'))
        $nativePath = Join-Path $nativeFolder 'msalruntime.dll'
        if (-not (Test-Path $nativePath)) { throw "MSAL broker runtime not found: $nativePath" }
        if ($env:PATH -notlike "*$nativeFolder*") { $env:PATH = "$nativeFolder;$env:PATH" }
        if ($State.NativeLibrary -eq [IntPtr]::Zero) {
            $State.NativeLibrary = [System.Runtime.InteropServices.NativeLibrary]::Load($nativePath)
        }

        $options = [Microsoft.Identity.Client.BrokerOptions]::new(
            [Microsoft.Identity.Client.BrokerOptions+OperatingSystems]::Windows
        )
        $builder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create('fb78d390-0c51-40cd-8e17-fdbfab77341b')
        $builder = [Microsoft.Identity.Client.Broker.BrokerExtension]::WithBroker($builder, $options)
        $app = $builder.WithAuthority('https://login.microsoftonline.com/organizations').Build()
        $scopes = [string[]]@('https://outlook.office365.com/.default')
        $authResult = $app.AcquireTokenInteractive($scopes).
            WithParentActivityOrWindow([IntPtr]$OwnerHandle).
            ExecuteAsync().GetAwaiter().GetResult()

        Connect-ExchangeOnline -AccessToken $authResult.AccessToken `
            -UserPrincipalName $authResult.Account.Username -ShowBanner:$false -ErrorAction Stop
        $State.MsalApp = $app
        $State.MsalAccount = $authResult.Account
        $State.AccessToken = $authResult.AccessToken
        $State.UserPrincipalName = $authResult.Account.Username
        $State.OwnerHandle = [IntPtr]$OwnerHandle
        $State.Status = "Connected as $($authResult.Account.Username)."
    }
    catch {
        if ($_.Exception.ToString() -match 'authentication_canceled|user canceled|user cancelled') {
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
        $authResult = $State.MsalApp.AcquireTokenInteractive($scopes).
            WithParentActivityOrWindow([IntPtr]$State.OwnerHandle).
            ExecuteAsync().GetAwaiter().GetResult()
    }

    if ($authResult.AccessToken -cne $State.AccessToken) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Connect-ExchangeOnline -AccessToken $authResult.AccessToken `
            -UserPrincipalName $authResult.Account.Username -ShowBanner:$false -ErrorAction Stop
        $State.AccessToken = $authResult.AccessToken
        $State.UserPrincipalName = $authResult.Account.Username
    }
}

$expandScript = {
    param($State, $ParentGroupIdentity, $ExportFolder, $EnsureConnection)
    try {
        & $EnsureConnection $State

        $script:ProcessedGroups = @{}
        $script:ExpandedRecipients = New-Object 'System.Collections.Generic.List[object]'
        $script:ExpansionWarnings = New-Object 'System.Collections.Generic.List[string]'

        function Add-ExpandedRecipient {
            param([object]$Recipient, [string]$SourceGroup, [string]$MembershipType)
            if (-not $Recipient.PrimarySmtpAddress) { return }
            $script:ExpandedRecipients.Add([PSCustomObject]@{
                DisplayName          = $Recipient.DisplayName
                Alias                = $Recipient.Alias
                PrimarySmtpAddress   = $Recipient.PrimarySmtpAddress.ToString()
                RecipientTypeDetails = $Recipient.RecipientTypeDetails
                SourceGroup          = $SourceGroup
                MembershipType       = $MembershipType
            })
        }

        function Expand-MailGroup {
            param([string]$GroupIdentity)
            $dynamicGroup = Get-DynamicDistributionGroup -Identity $GroupIdentity -ErrorAction SilentlyContinue
            if ($dynamicGroup) {
                $groupKey = "Dynamic:$($dynamicGroup.Guid)"
                if ($script:ProcessedGroups.ContainsKey($groupKey)) {
                    $script:ExpansionWarnings.Add("Already processed dynamic group: $($dynamicGroup.DisplayName)")
                    return
                }
                $script:ProcessedGroups[$groupKey] = $true
                $State.Status = "Evaluating dynamic group $($dynamicGroup.DisplayName)..."
                try {
                    $query = @{
                        RecipientPreviewFilter = $dynamicGroup.RecipientFilter
                        ResultSize             = 'Unlimited'
                        ErrorAction            = 'Stop'
                    }
                    if ($dynamicGroup.RecipientContainer) { $query.OrganizationalUnit = $dynamicGroup.RecipientContainer }
                    $dynamicMembers = @(Get-Recipient @query)
                    foreach ($member in $dynamicMembers) {
                        Add-ExpandedRecipient -Recipient $member -SourceGroup $dynamicGroup.DisplayName -MembershipType 'Dynamic'
                    }
                }
                catch {
                    $script:ExpansionWarnings.Add("Could not evaluate dynamic group '$($dynamicGroup.DisplayName)': $($_.Exception.Message)")
                }
                return
            }

            $regularGroup = Get-DistributionGroup -Identity $GroupIdentity -ErrorAction SilentlyContinue
            if (-not $regularGroup) {
                $script:ExpansionWarnings.Add("'$GroupIdentity' could not be identified as a regular or dynamic distribution group.")
                return
            }

            $groupKey = "Regular:$($regularGroup.Guid)"
            if ($script:ProcessedGroups.ContainsKey($groupKey)) {
                $script:ExpansionWarnings.Add("Already processed group: $($regularGroup.DisplayName)")
                return
            }
            $script:ProcessedGroups[$groupKey] = $true
            $State.Status = "Expanding group $($regularGroup.DisplayName)..."

            try {
                $members = @(Get-DistributionGroupMember -Identity $regularGroup.Identity -ResultSize Unlimited -ErrorAction Stop)
            }
            catch {
                $script:ExpansionWarnings.Add("Could not retrieve members of '$($regularGroup.DisplayName)': $($_.Exception.Message)")
                return
            }

            foreach ($member in $members) {
                if (-not $member.RecipientTypeDetails) {
                    $script:ExpansionWarnings.Add("Could not identify the type for member '$($member.DisplayName)'.")
                    continue
                }
                $recipientType = $member.RecipientTypeDetails.ToString()
                if ($recipientType -in @('DynamicDistributionGroup', 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup')) {
                    $identity = if ($member.PrimarySmtpAddress) { $member.PrimarySmtpAddress.ToString() } else { $member.Identity.ToString() }
                    Expand-MailGroup -GroupIdentity $identity
                }
                else {
                    Add-ExpandedRecipient -Recipient $member -SourceGroup $regularGroup.DisplayName -MembershipType 'Direct / Exception'
                }
            }
        }

        $parentGroup = Get-DynamicDistributionGroup -Identity $ParentGroupIdentity -ErrorAction SilentlyContinue
        if (-not $parentGroup) { $parentGroup = Get-DistributionGroup -Identity $ParentGroupIdentity -ErrorAction Stop }
        $State.Status = "Expanding $($parentGroup.DisplayName)..."
        $parentIdentity = if ($parentGroup.PrimarySmtpAddress) { $parentGroup.PrimarySmtpAddress.ToString() } else { $ParentGroupIdentity }
        Expand-MailGroup -GroupIdentity $parentIdentity

        $finalMembers = @(
            $script:ExpandedRecipients |
                Group-Object -Property PrimarySmtpAddress |
                ForEach-Object {
                    $recipientRows = $_.Group
                    $firstRecipient = $recipientRows | Select-Object -First 1
                    [PSCustomObject]@{
                        DisplayName          = $firstRecipient.DisplayName
                        Alias                = $firstRecipient.Alias
                        PrimarySmtpAddress   = $firstRecipient.PrimarySmtpAddress
                        RecipientTypeDetails = $firstRecipient.RecipientTypeDetails
                        SourceGroups         = (($recipientRows.SourceGroup | Sort-Object -Unique) -join '; ')
                        MembershipTypes      = (($recipientRows.MembershipType | Sort-Object -Unique) -join '; ')
                    }
                } |
                Sort-Object DisplayName
        )

        if ($script:ExpansionWarnings.Count -gt 0) {
            $State.Warnings = $script:ExpansionWarnings.ToArray()
        }
        else {
            $State.Warnings = @()
        }

        if ($State.Warnings.Count -gt 0) {
            throw "Expansion was incomplete. Review the Warnings tab; no CSV was written."
        }

        if (-not (Test-Path -LiteralPath $ExportFolder -PathType Container)) {
            throw "The export folder does not exist: $ExportFolder"
        }
        $safeGroupName = $parentGroup.DisplayName -replace '[\\/:*?"<>|]', '_'
        $outputPath = Join-Path $ExportFolder "$safeGroupName-Expanded-Members.csv"
        $State.Status = "Writing $($finalMembers.Count) recipient(s) to CSV..."
        $finalMembers | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        $State.Results = $finalMembers
        $State.OutputPath = $outputPath
        $State.ProcessedGroupCount = $script:ProcessedGroups.Count
        $State.Percent = 100
        $State.Status = "Exported $($finalMembers.Count) unique recipient(s) from $($script:ProcessedGroups.Count) group(s)."
    }
    catch {
        $State.Error = $_.Exception.Message
    }
    finally {
        $State.Done = $true
    }
}

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(150)
$timer.Add_Tick({
    $operationStatusText.Text = [string]$state.Status
    if (-not $operationProgressBar.IsIndeterminate) { $operationProgressBar.Value = [double]$state.Percent }

    if ($state.Busy -and $state.Done -and -not $state.CompletionHandled) {
        $state.CompletionHandled = $true
        try {
            if ($script:activePowerShell -and $script:asyncResult) { $null = $script:activePowerShell.EndInvoke($script:asyncResult) }
        }
        catch {
            if (-not $state.Error) { $state.Error = $_.Exception.Message }
        }
        finally {
            if ($script:activePowerShell) { $script:activePowerShell.Dispose() }
            $script:activePowerShell = $null
            $script:asyncResult = $null
        }

        $completedOperation = $state.Operation
        $state.Busy = $false
        $operationProgressBar.IsIndeterminate = $false
        Set-InputsEnabled -Enabled $true

        if ($state.Error) {
            $operationStatusText.Text = "$completedOperation failed: $($state.Error)"
            if ($completedOperation -eq 'Connect') {
                $connectionStatusText.Text = 'Not connected'
                $connectionStatusText.Foreground = '#B42318'
            }
            if ($state.Warnings) { $warningsList.ItemsSource = @($state.Warnings) }
            if ($state.Error -ne 'Sign-in cancelled.') {
                [System.Windows.MessageBox]::Show($window, [string]$state.Error, "$completedOperation failed",
                    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            }
        }
        elseif ($completedOperation -eq 'Connect') {
            $state.Connected = $true
            $connectionStatusText.Text = "Connected as $($state.UserPrincipalName)"
            $connectionStatusText.Foreground = '#157347'
            $connectButton.Content = 'Reconnect'
            $operationStatusText.Text = 'Connected. Enter a parent group and export folder.'
        }
        else {
            $resultsGrid.ItemsSource = @($state.Results)
            $warningsList.ItemsSource = @($state.Warnings)
            $operationProgressBar.Value = 100
            $operationStatusText.Text = "$($state.Status) CSV: $($state.OutputPath)"
            [System.Windows.MessageBox]::Show($window,
                "Export complete.`n`nUnique recipients: $(@($state.Results).Count)`nGroups processed: $($state.ProcessedGroupCount)`n`nCSV: $($state.OutputPath)",
                'Export complete', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        }

        Update-ExportEnabledState
    }
})
$timer.Start()

$connectButton.Add_Click({
    $ownerHandle = (New-Object System.Windows.Interop.WindowInteropHelper($window)).Handle
    Start-BackgroundOperation -Operation 'Connect' -Script $connectScript -Arguments @($ownerHandle)
})

$browseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Choose the folder for the expanded-members CSV.'
    $dialog.SelectedPath = if (Test-Path -LiteralPath $exportFolderTextBox.Text) { $exportFolderTextBox.Text } else { $HOME }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $exportFolderTextBox.Text = $dialog.SelectedPath
        Update-ExportEnabledState
    }
})

$parentGroupTextBox.Add_TextChanged({ Update-ExportEnabledState })
$exportButton.Add_Click({
    $identity = $parentGroupTextBox.Text.Trim()
    $folder = $exportFolderTextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($identity)) {
        $operationStatusText.Text = 'Enter a parent distribution group identity.'
        return
    }
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        $operationStatusText.Text = 'Choose an existing export folder.'
        return
    }

    $resultsGrid.ItemsSource = $null
    $warningsList.ItemsSource = $null
    Start-BackgroundOperation -Operation 'Expand and export' -Script $expandScript `
        -Arguments @($identity, $folder, $ensureConnectionScript)
})

$closeButton.Add_Click({ $window.Close() })
$window.Add_Closing({
    param($sender, $eventArgs)
    if ($state.Busy) {
        $choice = [System.Windows.MessageBox]::Show($window,
            'An operation is still running. Stop it and close the application?', 'Operation in progress',
            [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($choice -ne [System.Windows.MessageBoxResult]::Yes) {
            $eventArgs.Cancel = $true
            return
        }
        if ($script:activePowerShell) { $script:activePowerShell.Stop() }
    }
})
$window.Add_Closed({
    $timer.Stop()
    if ($script:activePowerShell) { $script:activePowerShell.Dispose() }
    if ($runspace) { $runspace.Close(); $runspace.Dispose() }
})

$null = $window.ShowDialog()