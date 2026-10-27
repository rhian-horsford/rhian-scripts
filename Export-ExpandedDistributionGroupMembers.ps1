#requires -Version 7.6

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# The Windows app-theme preference is read again when Windows notifies us of a change.
$systemTheme = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name AppsUseLightTheme -ErrorAction SilentlyContinue
$isLightTheme = $null -eq $systemTheme -or [int]$systemTheme.AppsUseLightTheme -ne 0
$lightThemeColors = [ordered]@{
    '#131417' = '#F4F6FA'; '#30333C' = '#E0E4EB'; '#737B89' = '#697386'
    '#17191E' = '#FFFFFF'; '#F0F2F6' = '#202A38'; '#383C47' = '#C8CFDA'
    '#1E2027' = '#FFFFFF'; '#E9ECF2' = '#263243'; '#23262E' = '#F1F3F6'
    '#30343E' = '#D9DEE6'; '#252831' = '#E8ECF2'; '#B6BFCD' = '#455368'
    '#343843' = '#CDD4DE'; '#2A2E38' = '#E4EAF3'; '#173A78' = '#315D9D'
    '#1264FF' = '#0B5CE0'; '#F3F5F8' = '#202A38'; '#282B34' = '#E9EDF4'
    '#A6AFBD' = '#58667A'; '#19D8DF' = '#009AA6'; '#FF4568' = '#DB3656'
    '#B9A3FF' = '#8665D9'; '#D0FF47' = '#8AAF18'; '#F2F4F8' = '#1D2939'
    '#9AA4B3' = '#5B6778'; '#20232A' = '#F4F6FA'; '#FFB454' = '#8C5200'
    '#343842' = '#CDD4DE'; '#2B2F38' = '#D5DBE4'; '#2A2D36' = '#E8ECF2'
    '#E7EAF0' = '#263243'
    '#858E9C' = '#5C6879'; '#FF7185' = '#B4233B'; '#76E0A8' = '#08743C'
}

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Expanded Distribution Group Exporter"
        Height="820" Width="1120" MinHeight="620" MinWidth="800"
        WindowStartupLocation="CenterScreen" Background="#131417">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="MinWidth" Value="110"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="ButtonBorder" Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="True"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.88"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.72"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="ButtonBorder" Property="Background" Value="#30333C"/>
                                <Setter Property="Foreground" Value="#737B89"/>
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="1"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="Background" Value="#17191E"/>
            <Setter Property="Foreground" Value="#F0F2F6"/>
            <Setter Property="CaretBrush" Value="#F0F2F6"/>
            <Setter Property="BorderBrush" Value="#383C47"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#1E2027"/>
            <Setter Property="Foreground" Value="#E9ECF2"/>
            <Setter Property="RowBackground" Value="#1E2027"/>
            <Setter Property="AlternatingRowBackground" Value="#23262E"/>
            <Setter Property="BorderBrush" Value="#30343E"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#30343E"/>
            <Setter Property="VerticalGridLinesBrush" Value="#30343E"/>
            <Setter Property="RowHeaderWidth" Value="0"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#252831"/>
            <Setter Property="Foreground" Value="#B6BFCD"/>
            <Setter Property="BorderBrush" Value="#343843"/>
            <Setter Property="Padding" Value="9,8"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style TargetType="DataGridRow">
            <Setter Property="Foreground" Value="#E9ECF2"/>
            <Setter Property="BorderBrush" Value="#30343E"/>
            <Setter Property="MinHeight" Value="34"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#2A2E38"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#173A78"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="Foreground" Value="#E9ECF2"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="8,6"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#173A78"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="#1E2027"/>
            <Setter Property="BorderBrush" Value="#30343E"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="#A6AFBD"/>
            <Setter Property="Background" Value="#1E2027"/>
            <Setter Property="Padding" Value="13,8"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="TabBorder" Background="{TemplateBinding Background}" BorderBrush="Transparent" BorderThickness="0,0,0,2" Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#1264FF"/>
                                <Setter Property="Foreground" Value="#F3F5F8"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="#282B34"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ListBox">
            <Setter Property="Background" Value="#1E2027"/>
            <Setter Property="Foreground" Value="#E9ECF2"/>
            <Setter Property="BorderBrush" Value="#30343E"/>
        </Style>
        <Style TargetType="ListBoxItem">
            <Setter Property="Foreground" Value="#E9ECF2"/>
            <Setter Property="Padding" Value="8,6"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#173A78"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#2A2E38"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="ProgressBar">
            <Setter Property="Background" Value="#30333C"/>
            <Setter Property="Foreground" Value="#1264FF"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="5"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="1.2*"/>
                <ColumnDefinition Width="1.2*"/>
                <ColumnDefinition Width="1.2*"/>
                <ColumnDefinition Width="1.2*"/>
                <ColumnDefinition Width="1.2*"/>
            </Grid.ColumnDefinitions>
            <Border Background="#1264FF"/>
            <Border Grid.Column="1" Background="#19D8DF"/>
            <Border Grid.Column="2" Background="#FF4568"/>
            <Border Grid.Column="3" Background="#B9A3FF"/>
            <Border Grid.Column="4" Background="#D0FF47"/>
        </Grid>
        <Grid Grid.Row="1" Margin="22">
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
                <TextBlock Text="Expanded group members" FontSize="22" FontWeight="SemiBold" Foreground="#F2F4F8"/>
                <TextBlock Text="Expand distribution and dynamic groups plus Microsoft 365 members and owners; deduplicate and export source details." Margin="0,4,0,0" Foreground="#9AA4B3"/>
            </StackPanel>
            <Border Grid.Column="1" Background="#20232A" BorderBrush="#343842" BorderThickness="1" CornerRadius="6" Padding="10,7" Margin="12,0,12,0" VerticalAlignment="Center">
                <TextBlock x:Name="ConnectionStatusText" Text="Not connected" Foreground="#FFB454"/>
            </Border>
            <Button x:Name="ConnectButton" Grid.Column="2" Content="Connect" VerticalAlignment="Center" Background="#0B63CE" Foreground="White"/>
        </Grid>

        <Border Grid.Row="1" Background="#1E2027" BorderBrush="#2B2F38" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Margin="0,0,14,0">
                    <TextBlock Text="Parent group" FontWeight="SemiBold" Foreground="#E9ECF2" Margin="0,0,0,6"/>
                    <TextBox x:Name="ParentGroupTextBox" ToolTip="Enter the exact email address, alias, or identity for a distribution, dynamic, mail-enabled security, or Microsoft 365 group."/>
                </StackPanel>
                <StackPanel Grid.Column="1" Margin="0,0,14,0">
                    <TextBlock Text="CSV export folder" FontWeight="SemiBold" Foreground="#E9ECF2" Margin="0,0,0,6"/>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="ExportFolderTextBox" IsReadOnly="True"/>
                        <Button x:Name="BrowseButton" Grid.Column="1" Content="Browse..." Margin="8,0,0,0" Background="#2A2D36" Foreground="#E7EAF0"/>
                    </Grid>
                </StackPanel>
                <Button x:Name="ExportButton" Grid.Column="2" Content="Expand and export" VerticalAlignment="Bottom" IsEnabled="False" Background="#1264FF" Foreground="White"/>
            </Grid>
        </Border>

        <Border Grid.Row="2" Background="#1E2027" BorderBrush="#2B2F38" BorderThickness="1" CornerRadius="8" Padding="14">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <TextBlock Text="Export results" FontWeight="SemiBold" Foreground="#E9ECF2"/>
                <Grid Grid.Row="1" Margin="0,6,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="220"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock x:Name="OperationStatusText" Text="Connect to Exchange Online to begin." TextWrapping="Wrap" Foreground="#9AA4B3" VerticalAlignment="Center"/>
                    <ProgressBar x:Name="OperationProgressBar" Grid.Column="1" Height="8" Minimum="0" Maximum="100" Value="0" VerticalAlignment="Center"/>
                </Grid>
                <TabControl Grid.Row="2">
                    <TabItem Header="Recipients">
                        <DataGrid x:Name="ResultsGrid" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column"
                                  CanUserAddRows="False" GridLinesVisibility="Horizontal" BorderBrush="#30343E" MinRowHeight="28"
                                  VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Display name" Binding="{Binding DisplayName}" Width="1.4*"/>
                                <DataGridTextColumn Header="Email" Binding="{Binding PrimarySmtpAddress}" Width="1.5*"/>
                                <DataGridTextColumn Header="Recipient type" Binding="{Binding RecipientTypeDetails}" Width="1.2*"/>
                                <DataGridTextColumn Header="Source group types" Binding="{Binding SourceGroupTypes}" Width="1.5*"/>
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
            <TextBlock VerticalAlignment="Center" Foreground="#858E9C" Text="This tool reads Exchange Online and writes a local CSV."/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="CloseButton" Content="Close" Margin="0,0,8,0" Background="#2A2D36" Foreground="#E7EAF0"/>
            </StackPanel>
        </Grid>
        </Grid>
    </Grid>
</Window>
'@

$themeHexColors = @([regex]::Matches($xaml, '#[0-9A-Fa-f]{6}') | ForEach-Object { $_.Value } | Sort-Object -Unique)
$themeBrushMarkup = ''
foreach ($color in $themeHexColors) {
    $key = 'ThemeColor' + $color.Substring(1).ToUpperInvariant()
    $themeBrushMarkup += "        <SolidColorBrush x:Key=`"$key`" Color=`"$color`"/>`n"
    $xaml = $xaml.Replace($color, "{DynamicResource $key}")
}
$xaml = $xaml.Replace('<Window.Resources>', "<Window.Resources>`n$themeBrushMarkup")
[xml]$xaml = $xaml

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
$script:window = $window
$script:connectionStatusText = $connectionStatusText
$exportFolderTextBox.Text = Join-Path $HOME 'OneDrive - Slalom\Desktop'

function Get-SystemLightTheme {
    $setting = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name AppsUseLightTheme -ErrorAction SilentlyContinue
    return ($null -eq $setting -or [int]$setting.AppsUseLightTheme -ne 0)
}

function Set-WindowTheme {
    param([bool]$LightTheme)
    foreach ($color in $script:themeHexColors) {
        $key = 'ThemeColor' + $color.Substring(1).ToUpperInvariant()
        $targetColor = if ($LightTheme -and $script:lightThemeColors.Contains($color)) { $script:lightThemeColors[$color] } else { $color }
        $newBrush = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($targetColor)
        )
        $script:window.Resources[$key] = $newBrush
    }
    $script:isLightTheme = $LightTheme
}

function Set-ConnectionStatusTone {
    param([string]$Color)
    $key = 'ThemeColor' + $Color.TrimStart('#').ToUpperInvariant()
    $script:connectionStatusText.Foreground = $script:window.Resources[$key]
}

Set-WindowTheme -LightTheme $isLightTheme

$themeTimer = New-Object System.Windows.Threading.DispatcherTimer
$themeTimer.Interval = [TimeSpan]::FromSeconds(1)
$themeTimer.Add_Tick({
    $lightTheme = Get-SystemLightTheme
    if ($lightTheme -ne $script:isLightTheme) {
        Set-WindowTheme -LightTheme $lightTheme
    }
})
$themeTimer.Start()

$workArea = [System.Windows.SystemParameters]::WorkArea
$window.Height = [Math]::Min($window.Height, $workArea.Height)
$window.Width = [Math]::Min($window.Width, $workArea.Width)

# This synchronized state is the only data channel shared with the worker runspace.
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
        Set-ConnectionStatusTone -Color '#FFB454'
    }

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    $null = $ps.AddScript($Script).AddArgument($state)
    foreach ($argument in $Arguments) { $null = $ps.AddArgument($argument) }
    $script:activePowerShell = $ps
    $script:asyncResult = $ps.BeginInvoke()
}

# The broker prompt uses the WPF handle as its owner instead of opening a browser tab.
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

# Renew the delegated token before issuing Exchange commands from the worker.
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
            param([object]$Recipient, [string]$SourceGroup, [string]$SourceGroupType, [string]$MembershipType)
            if (-not $Recipient.PrimarySmtpAddress) { return }
            $script:ExpandedRecipients.Add([PSCustomObject]@{
                DisplayName          = $Recipient.DisplayName
                Alias                = $Recipient.Alias
                PrimarySmtpAddress   = $Recipient.PrimarySmtpAddress.ToString()
                RecipientTypeDetails = $Recipient.RecipientTypeDetails
                SourceGroup          = $SourceGroup
                SourceGroupType      = $SourceGroupType
                MembershipType       = $MembershipType
            })
        }

        function Add-GroupOwners {
            param([object]$Group, [string]$SourceGroup, [string]$SourceGroupType)
            foreach ($ownerIdentity in @($Group.ManagedBy)) {
                if (-not $ownerIdentity) { continue }
                try {
                    $owner = Get-Recipient -Identity $ownerIdentity.ToString() -ErrorAction Stop
                    Add-ExpandedRecipient -Recipient $owner -SourceGroup $SourceGroup `
                        -SourceGroupType $SourceGroupType -MembershipType 'Owner'
                }
                catch {
                    $script:ExpansionWarnings.Add("Could not resolve an owner of '$SourceGroup': $($_.Exception.Message)")
                }
            }
        }

        function Expand-MailGroup {
            param([string]$GroupIdentity)
            $m365Group = Get-UnifiedGroup -Identity $GroupIdentity -ErrorAction SilentlyContinue
            if ($m365Group) {
                $groupKey = "M365:$($m365Group.Guid)"
                if ($script:ProcessedGroups.ContainsKey($groupKey)) { return }
                $script:ProcessedGroups[$groupKey] = $true
                $State.Status = "Expanding Microsoft 365 group $($m365Group.DisplayName) members and owners..."
                try {
                    $m365Identity = if ($m365Group.Identity) { $m365Group.Identity.ToString() } else { $GroupIdentity }
                    $m365Members = @(Get-UnifiedGroupLinks -Identity $m365Identity -LinkType Members -ResultSize Unlimited -ErrorAction Stop)
                    $m365Owners = @(Get-UnifiedGroupLinks -Identity $m365Identity -LinkType Owners -ResultSize Unlimited -ErrorAction Stop)
                    foreach ($member in $m365Members) {
                        Add-ExpandedRecipient -Recipient $member -SourceGroup $m365Group.DisplayName `
                            -SourceGroupType 'Microsoft 365 Group' -MembershipType 'Member'
                    }
                    foreach ($owner in $m365Owners) {
                        Add-ExpandedRecipient -Recipient $owner -SourceGroup $m365Group.DisplayName `
                            -SourceGroupType 'Microsoft 365 Group' -MembershipType 'Owner'
                    }
                }
                catch {
                    $script:ExpansionWarnings.Add("Could not retrieve members and owners of Microsoft 365 group '$($m365Group.DisplayName)': $($_.Exception.Message)")
                }
                return
            }

            $dynamicGroup = Get-DynamicDistributionGroup -Identity $GroupIdentity -ErrorAction SilentlyContinue
            if ($dynamicGroup) {
                $groupKey = "Dynamic:$($dynamicGroup.Guid)"
                if ($script:ProcessedGroups.ContainsKey($groupKey)) { return }
                $script:ProcessedGroups[$groupKey] = $true
                $State.Status = "Evaluating dynamic group $($dynamicGroup.DisplayName)..."
                Add-GroupOwners -Group $dynamicGroup -SourceGroup $dynamicGroup.DisplayName `
                    -SourceGroupType 'Dynamic Distribution Group'
                try {
                    # Exchange evaluates the dynamic group's saved OPATH filter at export time.
                    $query = @{
                        RecipientPreviewFilter = $dynamicGroup.RecipientFilter
                        ResultSize             = 'Unlimited'
                        ErrorAction            = 'Stop'
                    }
                    if ($dynamicGroup.RecipientContainer) { $query.OrganizationalUnit = $dynamicGroup.RecipientContainer }
                    $dynamicMembers = @(Get-Recipient @query)
                    foreach ($member in $dynamicMembers) {
                        Add-ExpandedRecipient -Recipient $member -SourceGroup $dynamicGroup.DisplayName `
                            -SourceGroupType 'Dynamic Distribution Group' -MembershipType 'Dynamic'
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
            if ($script:ProcessedGroups.ContainsKey($groupKey)) { return }
            $script:ProcessedGroups[$groupKey] = $true
            $State.Status = "Expanding group $($regularGroup.DisplayName)..."
            Add-GroupOwners -Group $regularGroup -SourceGroup $regularGroup.DisplayName `
                -SourceGroupType ([string]$regularGroup.RecipientTypeDetails)

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
                if ($recipientType -in @('GroupMailbox', 'DynamicDistributionGroup', 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup')) {
                    $identity = if ($member.PrimarySmtpAddress) { $member.PrimarySmtpAddress.ToString() } else { $member.Identity.ToString() }
                    Expand-MailGroup -GroupIdentity $identity
                }
                else {
                    Add-ExpandedRecipient -Recipient $member -SourceGroup $regularGroup.DisplayName `
                        -SourceGroupType ([string]$regularGroup.RecipientTypeDetails) -MembershipType 'Member'
                }
            }
        }

        $parentGroup = Get-UnifiedGroup -Identity $ParentGroupIdentity -ErrorAction SilentlyContinue
        if (-not $parentGroup) { $parentGroup = Get-DynamicDistributionGroup -Identity $ParentGroupIdentity -ErrorAction SilentlyContinue }
        if (-not $parentGroup) { $parentGroup = Get-DistributionGroup -Identity $ParentGroupIdentity -ErrorAction Stop }
        $State.Status = "Expanding $($parentGroup.DisplayName)..."
        $parentIdentity = if ($parentGroup.PrimarySmtpAddress) { $parentGroup.PrimarySmtpAddress.ToString() } else { $ParentGroupIdentity }
        Expand-MailGroup -GroupIdentity $parentIdentity

        # Group by primary SMTP so overlapping paths retain one row with all source details.
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
                        SourceGroupTypes     = (($recipientRows.SourceGroupType | Sort-Object -Unique) -join '; ')
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

        # A partial expansion must not produce a CSV that looks complete.
        if ($State.Warnings.Count -gt 0) {
            throw "Expansion was incomplete. Review the Warnings tab; no CSV was written."
        }

        if (-not (Test-Path -LiteralPath $ExportFolder -PathType Container)) {
            throw "The export folder does not exist: $ExportFolder"
        }
        $safeGroupName = $parentGroup.DisplayName -replace '[\\/:*?"<>|]', '_'
        $outputPath = Join-Path $ExportFolder "$safeGroupName-Expanded-Members-and-Owners.csv"
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

# The dispatcher timer applies worker progress and results on the WPF UI thread.
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
                Set-ConnectionStatusTone -Color '#FF7185'
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
            Set-ConnectionStatusTone -Color '#76E0A8'
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
    $themeTimer.Stop()
    $timer.Stop()
    if ($script:activePowerShell) { $script:activePowerShell.Dispose() }
    if ($runspace) { $runspace.Close(); $runspace.Dispose() }
})

$null = $window.ShowDialog()