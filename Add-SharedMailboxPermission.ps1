#requires -Version 7.6

[CmdletBinding()]
param([switch]$CreateSharedMailbox)

Set-StrictMode -Version 2.0
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

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

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Shared Mailbox Permission Manager"
        Height="860" Width="920" MinHeight="640" MinWidth="720"
        WindowStartupLocation="CenterScreen" Background="#131417">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="16,8"/>
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
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="0,0,24,0"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Foreground" Value="#E9ECF2"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="#252831"/>
            <Setter Property="Foreground" Value="#F0F2F6"/>
            <Setter Property="BorderBrush" Value="#414550"/>
            <Setter Property="Padding" Value="6,4"/>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="#252831"/>
            <Setter Property="Foreground" Value="#F0F2F6"/>
            <Setter Property="Padding" Value="8,5"/>
            <Style.Triggers>
                <Trigger Property="IsHighlighted" Value="True">
                    <Setter Property="Background" Value="#173A78"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
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
                <TextBlock x:Name="HeadingText" Text="Shared mailbox permissions" FontSize="20" FontWeight="SemiBold" Foreground="#F2F4F8"/>
                <TextBlock x:Name="DescriptionText" Margin="0,4,0,0" Foreground="#9AA4B3" TextWrapping="Wrap"
                           Text="Grant or remove Full Access and/or Send As permissions on a shared mailbox."/>
            </StackPanel>
            <Border Grid.Column="1" Background="#20232A" BorderBrush="#343842" BorderThickness="1" CornerRadius="6" Padding="10,7" Margin="12,0,12,0" VerticalAlignment="Center">
                <TextBlock x:Name="ConnectionStatusText" Text="Not connected" Foreground="#FFB454"/>
            </Border>
            <Button x:Name="ConnectButton" Grid.Column="2" Content="Connect" VerticalAlignment="Center" Background="#0B63CE" Foreground="White"/>
        </Grid>

        <Grid x:Name="CreationFields" Visibility="Collapsed" Margin="0,0,0,12">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel Margin="0,0,12,0">
                <TextBlock Text="Display name" FontWeight="SemiBold" Foreground="#E9ECF2" Margin="0,0,0,6"/>
                <TextBox x:Name="NameTextBox" MaxLength="64"/>
            </StackPanel>
            <StackPanel Grid.Column="1">
                <TextBlock Text="Alias" FontWeight="SemiBold" Foreground="#E9ECF2" Margin="0,0,0,6"/>
                <TextBox x:Name="AliasTextBox" MaxLength="64"/>
            </StackPanel>
        </Grid>

        <StackPanel Margin="0,0,0,12">
            <TextBlock Text="Shared mailbox email address" FontWeight="SemiBold" Foreground="#E9ECF2" Margin="0,0,0,6"/>
            <TextBox x:Name="MailboxTextBox" ToolTip="Exact SMTP address or UPN of the shared mailbox."/>
        </StackPanel>

        <StackPanel Margin="0,0,0,12">
            <TextBlock Text="Users: names or email addresses (one per line, or separated by semicolons)" FontWeight="SemiBold" Foreground="#E9ECF2" Margin="0,0,0,6"/>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="UsersTextBox" Height="70" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                         ToolTip="Full or partial names (e.g. 'Jane Smith', 'Smith, Jane') or exact email addresses. Wildcards are not allowed."/>
                <Button x:Name="CheckUsersButton" Grid.Column="1" Content="Check users" VerticalAlignment="Top" Margin="10,0,0,0"
                    IsEnabled="False" Background="#2A2D36" Foreground="#E7EAF0"/>
            </Grid>
        </StackPanel>

        <Border Background="#1E2027" BorderBrush="#2B2F38" BorderThickness="1" CornerRadius="8" Padding="10">
            <StackPanel>
            <TextBlock Text="Permissions" FontWeight="SemiBold" Foreground="#E9ECF2" Margin="0,0,0,8"/>
                <StackPanel Orientation="Horizontal">
                    <CheckBox x:Name="FullAccessCheckBox" Content="Full Access" IsChecked="True"/>
                    <CheckBox x:Name="SendAsCheckBox" Content="Send As"/>
                    <CheckBox x:Name="AutoMappingCheckBox" Content="Auto-map in Outlook (adding Full Access only)" IsChecked="True"/>
                </StackPanel>
                <TextBlock x:Name="PermissionHintText" Foreground="#9AA4B3" Margin="0,8,0,0" TextWrapping="Wrap"
                           Text="Removal checks direct permissions first. Inherited or group-based access is not removed. Auto-mapping applies only when adding."/>
                <TextBlock x:Name="ValidationText" Foreground="#FF7185" Margin="0,8,0,0" TextWrapping="Wrap" Visibility="Collapsed"/>
            </StackPanel>
        </Border>
        </StackPanel>

        <Border Grid.Row="1" Background="#1E2027" BorderBrush="#2B2F38" BorderThickness="1" CornerRadius="8" Padding="14">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <TextBlock Text="Status" FontWeight="SemiBold" Foreground="#E9ECF2"/>
                <TextBlock x:Name="OperationStatusText" Grid.Row="1" Text="Ready." Margin="0,7,0,10" TextWrapping="Wrap" Foreground="#9AA4B3"/>
                <ProgressBar x:Name="OperationProgressBar" Grid.Row="2" Height="8" Minimum="0" Maximum="100" Value="0" Margin="0,0,0,10"/>
                <DataGrid x:Name="ReviewGrid" Grid.Row="3" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column"
                          CanUserAddRows="False" GridLinesVisibility="Horizontal" BorderBrush="#30343E" MinRowHeight="32"
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
                          CanUserAddRows="False" GridLinesVisibility="Horizontal" BorderBrush="#30343E" MinRowHeight="28" Visibility="Collapsed"
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
            <TextBlock VerticalAlignment="Center" Foreground="#858E9C" Text="Changes apply immediately."/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="CloseButton" Content="Close" Margin="0,0,8,0" Background="#2A2D36" Foreground="#E7EAF0"/>
                <Button x:Name="RemoveButton" Content="Remove permissions" IsEnabled="False" Margin="0,0,8,0" Background="#FF4568" Foreground="White"/>
                <Button x:Name="AddButton" Content="Add permissions" IsEnabled="False" Background="#1264FF" Foreground="White"/>
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
$removeButton = $window.FindName('RemoveButton')
$closeButton = $window.FindName('CloseButton')
$nameTextBox = $window.FindName('NameTextBox')
$aliasTextBox = $window.FindName('AliasTextBox')
if ($CreateSharedMailbox) {
    $window.Title = 'Create Shared Mailbox'
    $window.FindName('HeadingText').Text = 'Create shared mailbox'
    $window.FindName('DescriptionText').Text = ''
    $window.FindName('PermissionHintText').Text = ''
    $window.FindName('CreationFields').Visibility = 'Visible'
    $removeButton.Visibility = 'Collapsed'
    $addButton.Content = 'Create mailbox'
}
$script:window = $window
$script:connectionStatusText = $connectionStatusText

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
    CreatedMailbox    = ''
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
    $mailboxTextBox.IsEnabled = $Enabled -and -not ($CreateSharedMailbox -and $state.CreatedMailbox)
    $nameTextBox.IsEnabled = $mailboxTextBox.IsEnabled
    $aliasTextBox.IsEnabled = $mailboxTextBox.IsEnabled
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

    $rows = @(Get-ReviewRows)
    $allResolved = $rows.Count -gt 0 -and @($rows | Where-Object { $null -eq $_.Selected }).Count -eq 0
    $addButton.IsEnabled = $ready -and $allResolved -and
        -not [string]::IsNullOrWhiteSpace($mailboxTextBox.Text) -and
        ([bool]$fullAccessCheckBox.IsChecked -or [bool]$sendAsCheckBox.IsChecked)
    $removeButton.IsEnabled = $addButton.IsEnabled
    if ($CreateSharedMailbox) {
        $addButton.IsEnabled = $addButton.IsEnabled -and -not $state.CreatedMailbox -and
            -not [string]::IsNullOrWhiteSpace($nameTextBox.Text) -and
            -not [string]::IsNullOrWhiteSpace($aliasTextBox.Text)
        $removeButton.IsEnabled = $false
    }
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
    $removeButton.IsEnabled = $false
    $checkUsersButton.IsEnabled = $false
    $operationProgressBar.IsIndeterminate = ($Operation -eq 'Connect')
    if ($Operation -eq 'Connect') {
        $state.Connected = $false
        $connectionStatusText.Text = 'Connecting...'
        Set-ConnectionStatusTone -Color '#FFB454'
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
                Set-ConnectionStatusTone -Color '#FF7185'
            }
            Show-ErrorDialog -Title "$completedOperation failed" -Message ([string]$state.Error)
        }
        elseif ($completedOperation -eq 'Connect') {
            $state.Connected = $true
            $connectionStatusText.Text = "Connected as $($state.UserPrincipalName)"
            Set-ConnectionStatusTone -Color '#76E0A8'
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
                if ($CreateSharedMailbox) { "All $($rows.Count) user(s) resolved. Review the matches before creating the mailbox." }
                else { "All $($rows.Count) user(s) resolved. Review the matches and select Add permissions or Remove permissions." }
            }
            else {
                "$unresolved entry/entries need attention. Choose a match or correct the input and check again."
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
            $warnings = @($results | Where-Object { $_.Result -eq 'Warning' }).Count
            $operationStatusText.Text = "Completed: $($results.Count - $failed - $warnings) processed or unchanged, $warnings warning(s), $failed failed."
            if ($CreateSharedMailbox -and $state.CreatedMailbox) {
                $operationStatusText.Text += " Mailbox $($state.CreatedMailbox) was created. To retry failed permissions, use Add-SharedMailboxPermission.ps1 after provisioning completes."
            }
            $icon = if ($failed -gt 0 -or $warnings -gt 0) { [System.Windows.MessageBoxImage]::Warning } else { [System.Windows.MessageBoxImage]::Information }
            [System.Windows.MessageBox]::Show(
                $window,
                "$($operationStatusText.Text)`n`nSee the results table for details. Successful changes are not rolled back if another permission fails.",
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
    param($State, $Mailbox, $Trustees, $GrantFullAccess, $GrantSendAs, $AutoMapping, $EnsureConnection,
        [ValidateSet('Add', 'Remove')][string]$Action = 'Add', [bool]$CompleteState = $true)
    try {
        & $EnsureConnection $State
        if (-not ($GrantFullAccess -or $GrantSendAs) -or $Trustees.Count -eq 0) {
            throw 'Select at least one permission and resolve at least one user.'
        }
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
                $State.Status = "$Action $permission for $user ($step of $total)..."
                try {
                    $warnings = @()
                    $result = 'Success'
                    if ($Action -eq 'Remove') {
                        $State.Status = "Checking $permission for $user ($step of $total)..."
                        $assignments = if ($permission -eq 'Full Access') {
                            @(Get-MailboxPermission -Identity $target -User $trustee.Address -ErrorAction Stop |
                                Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.Deny })
                        }
                        else {
                            @(Get-RecipientPermission -Identity $target -Trustee $trustee.Address -AccessRights SendAs `
                                -ResultSize Unlimited -ErrorAction Stop |
                                Where-Object { $_.AccessRights -contains 'SendAs' -and $_.AccessControlType -eq 'Allow' })
                        }
                        $direct = @($assignments | Where-Object { -not $_.IsInherited })
                        $inherited = @($assignments | Where-Object { $_.IsInherited })
                        if ($direct.Count -eq 0) {
                            if ($inherited.Count -gt 0) {
                                $result = 'Warning'
                                $details = 'Only inherited permission was found; it cannot be removed here. Change the source assignment.'
                            }
                            else {
                                $result = 'Not present'
                                $details = 'No direct allow permission found; no change. Group-based access is not evaluated or removed.'
                            }
                        }
                        else {
                            if ($permission -eq 'Full Access') {
                                $null = Remove-MailboxPermission -Identity $target -User $trustee.Address `
                                    -AccessRights FullAccess -InheritanceType All -Confirm:$false `
                                    -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings
                            }
                            else {
                                $null = Remove-RecipientPermission -Identity $target -Trustee $trustee.Address `
                                    -AccessRights SendAs -Confirm:$false -ErrorAction Stop `
                                    -WarningAction SilentlyContinue -WarningVariable warnings
                            }
                            $details = 'Direct permission removed. Group-based access, if any, is unchanged.'
                            if ($inherited.Count -gt 0) {
                                $result = 'Warning'
                                $details += ' Inherited permission remains; change the source assignment to remove it.'
                            }
                        }
                    }
                    elseif ($permission -eq 'Full Access') {
                        $null = Add-MailboxPermission -Identity $target -User $trustee.Address `
                            -AccessRights FullAccess -InheritanceType All -AutoMapping $AutoMapping `
                            -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings
                        $details = 'Granted.'
                    }
                    else {
                        $null = Add-RecipientPermission -Identity $target -Trustee $trustee.Address `
                            -AccessRights SendAs -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings
                        $details = 'Granted.'
                    }
                    if ($warnings.Count -gt 0) {
                        $result = 'Warning'
                        $details += ' ' + (($warnings | ForEach-Object { $_.Message }) -join ' ')
                    }
                    $results.Add([pscustomobject]@{ User = $user; Permission = $permission; Result = $result; Details = $details })
                }
                catch {
                    $results.Add([pscustomobject]@{ User = $user; Permission = $permission; Result = 'Failed'; Details = $_.Exception.Message })
                }
                $State.Results = $results.ToArray()
                $State.Percent = [int](($step / $total) * 100)
            }
        }

        $State.Results = $results.ToArray()
    }
    catch {
        $State.Error = $_.Exception.Message
    }
    finally {
        if ($CompleteState) { $State.Done = $true }
    }
}

$createScript = {
    param($State, $Name, $Alias, $Mailbox, $Trustees, $GrantFullAccess, $GrantSendAs, $AutoMapping,
        $EnsureConnection, $GrantPermissions, $EmailPattern)
    $creationResult = $null
    try {
        if ($State.CreatedMailbox) {
            throw "Mailbox $($State.CreatedMailbox) was already created. Use the permission manager to retry grants."
        }
        if ([string]::IsNullOrWhiteSpace($Name) -or $Name.Length -gt 64) {
            throw 'Enter a display name of 1 to 64 characters.'
        }
        if ($Alias -notmatch '^[a-zA-Z0-9._-]{1,64}$') {
            throw 'Enter an alias of 1 to 64 letters, numbers, dots, underscores, or hyphens.'
        }
        if ($Mailbox -notmatch $EmailPattern) { throw 'Enter a valid shared mailbox email address.' }
        if (-not ($GrantFullAccess -or $GrantSendAs) -or $Trustees.Count -eq 0 -or
            @($Trustees | Where-Object { -not $_.Address }).Count -gt 0) {
            throw 'Select at least one permission and resolve at least one user.'
        }
        & $EnsureConnection $State
        $State.Status = "Creating shared mailbox $Mailbox..."
        $warnings = @()
        $null = New-Mailbox -Shared -Name $Name -DisplayName $Name -Alias $Alias -PrimarySmtpAddress $Mailbox `
            -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings
        $State.CreatedMailbox = $Mailbox
        $creationResult = [pscustomobject]@{
            User = $Mailbox; Permission = 'Create mailbox'; Result = 'Success'; Details = 'Shared mailbox created.'
        }
        if ($warnings.Count -gt 0) {
            $creationResult.Result = 'Warning'
            $creationResult.Details += ' ' + (($warnings | ForEach-Object { $_.Message }) -join ' ')
        }
        $State.Results = @()
        & $GrantPermissions $State $Mailbox $Trustees $GrantFullAccess $GrantSendAs $AutoMapping {} Add $false
        if ($State.Error) {
            $State.Results = @($State.Results) + [pscustomobject]@{
                User = $Mailbox; Permission = 'Permission setup'; Result = 'Failed'; Details = [string]$State.Error
            }
            $State.Error = $null
        }
    }
    catch {
        $State.Error = $_.Exception.Message
    }
    finally {
        if ($null -ne $creationResult) {
            if ($State.Error) {
                $State.Results = @($State.Results) + [pscustomobject]@{
                    User = $Mailbox; Permission = 'Permission setup'; Result = 'Failed'; Details = [string]$State.Error
                }
                $State.Error = $null
            }
            $State.Results = @($creationResult) + @($State.Results)
            $State.Percent = 100
            $State.Status = "Mailbox $Mailbox created. Review permission results."
        }
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

$nameTextBox.Add_TextChanged({ $validationText.Visibility = 'Collapsed'; Update-ActionButtons })
$aliasTextBox.Add_TextChanged({ $validationText.Visibility = 'Collapsed'; Update-ActionButtons })

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
    $entries = @(Get-UserEntries)
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

function Invoke-PermissionAction {
    param([ValidateSet('Add', 'Remove')][string]$Action)
    $mailbox = $mailboxTextBox.Text.Trim()
    $rows = @(Get-ReviewRows)

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
    if ($CreateSharedMailbox) {
        if ($state.CreatedMailbox) { $errors += 'This mailbox was already created. Use the permission manager to retry grants.' }
        if ([string]::IsNullOrWhiteSpace($nameTextBox.Text) -or $nameTextBox.Text.Trim().Length -gt 64) {
            $errors += 'Enter a display name of 1 to 64 characters.'
        }
        if ($aliasTextBox.Text.Trim() -notmatch '^[a-zA-Z0-9._-]{1,64}$') {
            $errors += 'Enter an alias of 1 to 64 letters, numbers, dots, underscores, or hyphens.'
        }
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
    $verb = if ($Action -eq 'Remove') { 'Remove' } else { 'Grant' }
    $message = "$verb $permissionText on $mailbox for $($trustees.Count) user(s)?`n`n$(($trustees | ForEach-Object { $_.Label }) -join "`n")"
    if ($CreateSharedMailbox) {
        $message = "Create shared mailbox '$($nameTextBox.Text.Trim())' with alias '$($aliasTextBox.Text.Trim())' and email $mailbox, then grant $permissionText to $($trustees.Count) user(s)?`n`n$(($trustees | ForEach-Object { $_.Label }) -join "`n")"
        if ($fullAccessCheckBox.IsChecked) { $message += "`n`nAuto-map in Outlook: $([bool]$autoMappingCheckBox.IsChecked)." }
        $message += "`n`nThe mailbox will remain created even if permission grants fail. Retry failed grants with Add-SharedMailboxPermission.ps1 after provisioning completes."
    }
    if ($Action -eq 'Remove') {
        $message += "`n`nOnly existing direct allow permissions will be removed. Inherited permissions, deny entries, and group-based access are not removed. Auto-mapping is ignored."
    }
    $message += "`n`nSuccessful changes are not rolled back if another permission fails."
    $confirm = [System.Windows.MessageBox]::Show(
        $window,
        $message,
        'Confirm changes',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
        return
    }

    $resultsGrid.ItemsSource = $null
    if ($CreateSharedMailbox) {
        Start-BackgroundOperation -Operation 'Create mailbox' -Script $createScript `
            -Arguments @($nameTextBox.Text.Trim(), $aliasTextBox.Text.Trim(), $mailbox, $trustees,
                [bool]$fullAccessCheckBox.IsChecked, [bool]$sendAsCheckBox.IsChecked,
                [bool]$autoMappingCheckBox.IsChecked, $ensureConnectionScript, $addScript, $identifierPattern)
        return
    }
    Start-BackgroundOperation `
        -Operation "$Action permissions" `
        -Script $addScript `
        -Arguments @($mailbox, $trustees, [bool]$fullAccessCheckBox.IsChecked, [bool]$sendAsCheckBox.IsChecked, [bool]$autoMappingCheckBox.IsChecked, $ensureConnectionScript, $Action)
}
$addButton.Add_Click({ Invoke-PermissionAction -Action Add })
$removeButton.Add_Click({ Invoke-PermissionAction -Action Remove })

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
    $themeTimer.Stop()
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
