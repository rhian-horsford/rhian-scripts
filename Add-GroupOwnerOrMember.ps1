#requires -Version 7.6

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

class GroupUserRow {
    [string]$Entry
    [string]$Status
    [object[]]$Options
    [bool]$HasOptions
    [object]$Selected

    GroupUserRow([string]$entry, [string]$status, [object[]]$options) {
        $this.Entry = $entry
        $this.Status = $status
        $this.Options = $options
        $this.HasOptions = $options.Count -gt 0
        if ($options.Count -eq 1) {
            $this.Selected = $options[0]
        }
    }
}

$connectScript = {
    param($State, $OwnerHandle)
    try {
        $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
            Sort-Object Version -Descending | Select-Object -First 1
        if (-not $module -or $module.Version -lt [version]'3.7') {
            throw "ExchangeOnlineManagement 3.7 or later is required. Close this tool and run 'Install-Module ExchangeOnlineManagement -Scope CurrentUser' from a trusted PowerShell Gallery session."
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
        $scopes = [string[]]@('https://outlook.office365.com/.default')
        $brokerOptions = [Microsoft.Identity.Client.BrokerOptions]::new(
            [Microsoft.Identity.Client.BrokerOptions+OperatingSystems]::Windows
        )
        $builder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create('fb78d390-0c51-40cd-8e17-fdbfab77341b')
        $builder = [Microsoft.Identity.Client.Broker.BrokerExtension]::WithBroker($builder, $brokerOptions)
        $app = $builder.WithAuthority('https://login.microsoftonline.com/organizations').Build()
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
        $State.Status = "Connected to Exchange Online as $($authResult.Account.Username)."
    }
    catch {
        $State.Error = if ($_.Exception.ToString() -match 'authentication_canceled|user canceled|user cancelled') {
            'Sign-in cancelled.'
        }
        else { $_.Exception.Message }
    }
    finally { $State.Done = $true }
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
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Connect-ExchangeOnline -AccessToken $authResult.AccessToken `
            -UserPrincipalName $authResult.Account.Username -ShowBanner:$false -ErrorAction Stop
        $State.AccessToken = $authResult.AccessToken
        $State.UserPrincipalName = $authResult.Account.Username
    }
}

$resolveScript = {
    param($State, $Entries, $EmailPattern, $EnsureConnection)
    try {
        & $EnsureConnection $State
        $rows = [System.Collections.Generic.List[object]]::new()
        $rawLimit = 200
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
                    else { $status = "Not a user mailbox ($($recipient.RecipientTypeDetails))" }
                }
                else {
                    $raw = @(Get-Recipient -Anr $entry -RecipientTypeDetails UserMailbox -ResultSize $rawLimit -ErrorAction Stop)
                    $words = @($entry -split '[\s,]+' | Where-Object { $_ })
                    # Require every typed word to match a name, alias, or address prefix.
                    $found = @($raw | Where-Object {
                        $candidateWords = @("$($_.DisplayName) $($_.FirstName) $($_.LastName) $($_.Alias) $($_.PrimarySmtpAddress)" -split '[\s,.@_-]+' | Where-Object { $_ })
                        $missing = @($words | Where-Object {
                            $word = $_
                            -not ($candidateWords | Where-Object { $_.StartsWith($word, [StringComparison]::OrdinalIgnoreCase) })
                        })
                        $missing.Count -eq 0
                    })
                    if ($raw.Count -ge $rawLimit) {
                        $status = 'Too many matches - be more specific'
                        $found = @()
                    }
                }
            }
            catch { $status = "Lookup failed: $($_.Exception.Message)" }
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
                    Label = "$($_.DisplayName) <$($_.PrimarySmtpAddress)>"
                }
            })
            $rows.Add([pscustomobject]@{ Entry = $entry; Status = $status; Options = $options })
            $State.Percent = [int](($index / $Entries.Count) * 100)
        }
        $State.Results = $rows.ToArray()
    }
    catch { $State.Error = $_.Exception.Message }
    finally { $State.Done = $true }
}

$addScript = {
    param($State, $Group, $Users, $AddMember, $AddOwner, $BypassManagerCheck, $EnsureConnection)
    try {
        & $EnsureConnection $State
        if (-not ($AddMember -or $AddOwner) -or $Users.Count -eq 0) {
            throw 'Select at least one role and resolve at least one user.'
        }
        $recipient = Get-Recipient -Identity $Group -ErrorAction Stop
        $isUnified = $recipient.RecipientTypeDetails -eq 'GroupMailbox'
        if (-not $isUnified -and $recipient.RecipientTypeDetails -notin @('MailUniversalDistributionGroup', 'MailUniversalSecurityGroup')) {
            throw "'$Group' is a $($recipient.RecipientTypeDetails), not a supported static distribution list or Microsoft 365 group. Dynamic group membership cannot be edited."
        }
        $target = [string]$recipient.PrimarySmtpAddress
        $groupObject = if ($isUnified) {
            Get-UnifiedGroup -Identity $target -ErrorAction Stop
        }
        else { Get-DistributionGroup -Identity $target -ErrorAction Stop }
        if ($groupObject.PSObject.Properties['IsDirSynced'] -and $groupObject.IsDirSynced) {
            throw "'$target' is synchronized from on-premises. Make this change in the source directory."
        }

        $members = @{}
        $owners = @{}
        $needsMembers = $AddMember -or ($isUnified -and $AddOwner)
        $State.Status = "Reading existing roles on $target..."
        if ($needsMembers) {
            $existingMembers = if ($isUnified) {
                @(Get-UnifiedGroupLinks -Identity $target -LinkType Members -ResultSize Unlimited -ErrorAction Stop)
            }
            else { @(Get-DistributionGroupMember -Identity $target -ResultSize Unlimited -ErrorAction Stop) }
            foreach ($member in $existingMembers) {
                $members[[string]$member.PrimarySmtpAddress] = $true
            }
        }
        if ($AddOwner) {
            $existingOwners = if ($isUnified) {
                @(Get-UnifiedGroupLinks -Identity $target -LinkType Owners -ResultSize Unlimited -ErrorAction Stop)
            }
            else {
                @(foreach ($owner in $groupObject.ManagedBy) {
                    Get-Recipient -Identity $owner -ErrorAction Stop
                })
            }
            foreach ($owner in $existingOwners) {
                $owners[[string]$owner.PrimarySmtpAddress] = $true
            }
        }

        $roles = @(
            if ($needsMembers) { 'Member' }
            if ($AddOwner) { 'Owner' }
        )
        $results = [System.Collections.Generic.List[object]]::new()
        $total = $Users.Count * $roles.Count
        $step = 0
        foreach ($user in $Users) {
            foreach ($role in $roles) {
                $step++
                $State.Status = "Adding $role on $target for $($user.Label) ($step of $total)..."
                $roleSet = if ($role -eq 'Member') { $members } else { $owners }
                $details = ''
                $result = 'Success'
                try {
                    if ($roleSet.ContainsKey($user.Address)) {
                        $result = 'Already present'
                        $details = "Already a $($role.ToLowerInvariant()); no change."
                    }
                    else {
                        if ($isUnified -and $role -eq 'Owner' -and -not $members.ContainsKey($user.Address)) {
                            throw 'Ownership was not added because the required membership could not be added. See the Member result.'
                        }
                        $warnings = @()
                        if ($isUnified) {
                            $linkType = if ($role -eq 'Member') { 'Members' } else { 'Owners' }
                            $null = Add-UnifiedGroupLinks -Identity $target -LinkType $linkType -Links $user.Address `
                                -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings
                        }
                        elseif ($role -eq 'Member') {
                            $null = Add-DistributionGroupMember -Identity $target -Member $user.Address `
                                -BypassSecurityGroupManagerCheck:$BypassManagerCheck -Confirm:$false `
                                -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings
                        }
                        else {
                            $null = Set-DistributionGroup -Identity $target -ManagedBy @{ Add = $user.Address } `
                                -BypassSecurityGroupManagerCheck:$BypassManagerCheck -Confirm:$false `
                                -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings
                        }
                        $roleSet[$user.Address] = $true
                        $details = 'Added.'
                        if ($warnings.Count -gt 0) {
                            $result = 'Warning'
                            $details = ($warnings | ForEach-Object { $_.Message }) -join ' '
                        }
                    }
                }
                catch {
                    $result = 'Failed'
                    $details = $_.Exception.Message
                }
                if ($isUnified -and $role -eq 'Member' -and $AddOwner -and -not $AddMember) {
                    $details = "Required for Microsoft 365 ownership. $details"
                }
                $results.Add([pscustomobject]@{
                    User = $user.Label; Role = $role; Result = $result; Details = $details
                })
                $State.Results = $results.ToArray()
                $State.Percent = [int](($step / $total) * 100)
            }
        }
    }
    catch { $State.Error = $_.Exception.Message }
    finally { $State.Done = $true }
}

$lightThemeColors = [ordered]@{
    '#131417' = '#F4F6FA'; '#30333C' = '#E0E4EB'; '#737B89' = '#697386'
    '#17191E' = '#FFFFFF'; '#F0F2F6' = '#202A38'; '#383C47' = '#C8CFDA'
    '#1E2027' = '#FFFFFF'; '#E9ECF2' = '#263243'; '#23262E' = '#F1F3F6'
    '#30343E' = '#D9DEE6'; '#252831' = '#E8ECF2'; '#B6BFCD' = '#455368'
    '#343843' = '#CDD4DE'; '#2A2E38' = '#E4EAF3'; '#173A78' = '#315D9D'
    '#1264FF' = '#0B5CE0'; '#B9A3FF' = '#8665D9'; '#D0FF47' = '#8AAF18'
    '#F2F4F8' = '#1D2939'; '#9AA4B3' = '#5B6778'; '#20232A' = '#F4F6FA'
    '#FFB454' = '#8C5200'; '#343842' = '#CDD4DE'; '#2B2F38' = '#D5DBE4'
    '#2A2D36' = '#E8ECF2'; '#E7EAF0' = '#263243'; '#858E9C' = '#5C6879'
    '#FF7185' = '#B4233B'; '#76E0A8' = '#08743C'; '#19D8DF' = '#009AA6'
    '#FF4568' = '#DB3656'
}

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Group Owner and Member Manager" Height="860" Width="920" MinHeight="640" MinWidth="720"
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
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="ButtonBorder" Property="Opacity" Value="0.88"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter TargetName="ButtonBorder" Property="Opacity" Value="0.72"/></Trigger>
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
            <Setter Property="Padding" Value="8,6"/><Setter Property="Background" Value="#17191E"/>
            <Setter Property="Foreground" Value="#F0F2F6"/><Setter Property="CaretBrush" Value="#F0F2F6"/>
            <Setter Property="BorderBrush" Value="#383C47"/><Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="0,0,24,0"/><Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Foreground" Value="#E9ECF2"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="#252831"/><Setter Property="Foreground" Value="#F0F2F6"/>
            <Setter Property="BorderBrush" Value="#414550"/><Setter Property="Padding" Value="6,4"/>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="#252831"/><Setter Property="Foreground" Value="#F0F2F6"/>
            <Setter Property="Padding" Value="8,5"/>
            <Style.Triggers>
                <Trigger Property="IsHighlighted" Value="True">
                    <Setter Property="Background" Value="#173A78"/><Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#1E2027"/><Setter Property="Foreground" Value="#E9ECF2"/>
            <Setter Property="RowBackground" Value="#1E2027"/><Setter Property="AlternatingRowBackground" Value="#23262E"/>
            <Setter Property="BorderBrush" Value="#30343E"/><Setter Property="HorizontalGridLinesBrush" Value="#30343E"/>
            <Setter Property="VerticalGridLinesBrush" Value="#30343E"/><Setter Property="RowHeaderWidth" Value="0"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#252831"/><Setter Property="Foreground" Value="#B6BFCD"/>
            <Setter Property="BorderBrush" Value="#343843"/><Setter Property="Padding" Value="9,8"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style TargetType="DataGridRow">
            <Setter Property="Foreground" Value="#E9ECF2"/><Setter Property="BorderBrush" Value="#30343E"/>
            <Setter Property="MinHeight" Value="34"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#2A2E38"/></Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#173A78"/><Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="Foreground" Value="#E9ECF2"/><Setter Property="BorderThickness" Value="0"/><Setter Property="Padding" Value="8,6"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#173A78"/><Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="ProgressBar">
            <Setter Property="Background" Value="#30333C"/><Setter Property="Foreground" Value="#1264FF"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions><RowDefinition Height="5"/><RowDefinition Height="*"/></Grid.RowDefinitions>
        <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="1.2*"/><ColumnDefinition Width="1.2*"/><ColumnDefinition Width="1.2*"/>
                <ColumnDefinition Width="1.2*"/><ColumnDefinition Width="1.2*"/>
            </Grid.ColumnDefinitions>
            <Border Background="#1264FF"/><Border Grid.Column="1" Background="#19D8DF"/>
            <Border Grid.Column="2" Background="#FF4568"/><Border Grid.Column="3" Background="#B9A3FF"/>
            <Border Grid.Column="4" Background="#D0FF47"/>
        </Grid>
        <Grid Grid.Row="1" Margin="22">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/><RowDefinition Height="*" MinHeight="200"/><RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <StackPanel Grid.Row="0" Margin="0,0,0,12">
                <Grid Margin="0,0,0,14">
                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock Text="Group owners and members" FontSize="20" FontWeight="SemiBold" Foreground="#F2F4F8"/>
                        <TextBlock Margin="0,4,0,0" Foreground="#9AA4B3" TextWrapping="Wrap"
                                   Text="Add members and/or owners to a distribution list or Microsoft 365 group."/>
                    </StackPanel>
                    <Border Grid.Column="1" Background="#20232A" BorderBrush="#343842" BorderThickness="1" CornerRadius="6" Padding="10,7" Margin="12,0,12,0" VerticalAlignment="Center">
                        <TextBlock x:Name="ConnectionStatusText" Text="Not connected" Foreground="#FFB454"/>
                    </Border>
                    <Button x:Name="ConnectButton" Grid.Column="2" Content="Connect" VerticalAlignment="Center" Background="#0B63CE" Foreground="White"/>
                </Grid>
                <StackPanel Margin="0,0,0,12">
                    <TextBlock Text="Group email address" FontWeight="SemiBold" Foreground="#E9ECF2" Margin="0,0,0,6"/>
                    <TextBox x:Name="GroupTextBox" ToolTip="Exact SMTP address of a static distribution list, mail-enabled security group, or Microsoft 365 group."/>
                </StackPanel>
                <StackPanel Margin="0,0,0,12">
                    <TextBlock Text="Users to add: names or email addresses (one per line, or separated by semicolons)" FontWeight="SemiBold" Foreground="#E9ECF2" Margin="0,0,0,6"/>
                    <Grid>
                        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                        <TextBox x:Name="UsersTextBox" Height="70" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                                 ToolTip="Full or partial names (e.g. 'Jane Smith', 'Smith, Jane') or exact email addresses of user mailboxes. Wildcards are not allowed."/>
                        <Button x:Name="CheckUsersButton" Grid.Column="1" Content="Check users" VerticalAlignment="Top" Margin="10,0,0,0"
                                IsEnabled="False" Background="#2A2D36" Foreground="#E7EAF0"/>
                    </Grid>
                </StackPanel>
                <Border Background="#1E2027" BorderBrush="#2B2F38" BorderThickness="1" CornerRadius="8" Padding="10">
                    <StackPanel>
                        <TextBlock Text="Roles" FontWeight="SemiBold" Foreground="#E9ECF2" Margin="0,0,0,8"/>
                        <StackPanel Orientation="Horizontal">
                            <CheckBox x:Name="MemberCheckBox" Content="Member" IsChecked="True"/>
                            <CheckBox x:Name="OwnerCheckBox" Content="Owner"/>
                            <CheckBox x:Name="BypassCheckBox" Content="Admin override (distribution lists only)"
                                      ToolTip="Bypass the group manager check. Requires Organization Management or the Security Group Creation and Membership role."/>
                        </StackPanel>
                        <TextBlock Foreground="#9AA4B3" Margin="0,8,0,0" TextWrapping="Wrap"
                                   Text="Microsoft 365 owners are also added as members. Distribution-list owners are not automatically added as members."/>
                        <TextBlock x:Name="ValidationText" Foreground="#FF7185" Margin="0,8,0,0" TextWrapping="Wrap" Visibility="Collapsed"/>
                    </StackPanel>
                </Border>
            </StackPanel>
            <Border Grid.Row="1" Background="#1E2027" BorderBrush="#2B2F38" BorderThickness="1" CornerRadius="8" Padding="14">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <TextBlock Text="Status" FontWeight="SemiBold" Foreground="#E9ECF2"/>
                    <TextBlock x:Name="OperationStatusText" Grid.Row="1" Text="Ready." Margin="0,7,0,10" TextWrapping="Wrap" Foreground="#9AA4B3"/>
                    <ProgressBar x:Name="OperationProgressBar" Grid.Row="2" Height="8" Minimum="0" Maximum="100" Value="0" Margin="0,0,0,10"/>
                    <DataGrid x:Name="ReviewGrid" Grid.Row="3" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column"
                              CanUserAddRows="False" MinRowHeight="32" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
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
                              CanUserAddRows="False" MinRowHeight="28" Visibility="Collapsed"
                              VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="User" Binding="{Binding User}" Width="2*"/>
                            <DataGridTextColumn Header="Role" Binding="{Binding Role}" Width="*"/>
                            <DataGridTextColumn Header="Result" Binding="{Binding Result}" Width="*"/>
                            <DataGridTextColumn Header="Details" Binding="{Binding Details}" Width="3*">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock"><Setter Property="TextWrapping" Value="Wrap"/></Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </Border>
            <Grid Grid.Row="2" Margin="0,14,0,0">
                <TextBlock VerticalAlignment="Center" Foreground="#858E9C" Text="Changes are applied to Exchange Online immediately."/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="CloseButton" Content="Close" Margin="0,0,8,0" Background="#2A2D36" Foreground="#E7EAF0"/>
                    <Button x:Name="AddButton" Content="Add roles" IsEnabled="False" Background="#1264FF" Foreground="White"/>
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
$reader = [System.Xml.XmlNodeReader]::new([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$connectButton = $window.FindName('ConnectButton')
$connectionStatusText = $window.FindName('ConnectionStatusText')
$groupTextBox = $window.FindName('GroupTextBox')
$usersTextBox = $window.FindName('UsersTextBox')
$checkUsersButton = $window.FindName('CheckUsersButton')
$memberCheckBox = $window.FindName('MemberCheckBox')
$ownerCheckBox = $window.FindName('OwnerCheckBox')
$bypassCheckBox = $window.FindName('BypassCheckBox')
$validationText = $window.FindName('ValidationText')
$operationStatusText = $window.FindName('OperationStatusText')
$operationProgressBar = $window.FindName('OperationProgressBar')
$reviewGrid = $window.FindName('ReviewGrid')
$resultsGrid = $window.FindName('ResultsGrid')
$addButton = $window.FindName('AddButton')
$closeButton = $window.FindName('CloseButton')

function Get-SystemLightTheme {
    $setting = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name AppsUseLightTheme -ErrorAction SilentlyContinue
    return ($null -eq $setting -or [int]$setting.AppsUseLightTheme -ne 0)
}

function Set-WindowTheme {
    param([bool]$LightTheme)
    foreach ($color in $script:themeHexColors) {
        $key = 'ThemeColor' + $color.Substring(1).ToUpperInvariant()
        $targetColor = if ($LightTheme -and $script:lightThemeColors.Contains($color)) { $script:lightThemeColors[$color] } else { $color }
        $script:window.Resources[$key] = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($targetColor)
        )
    }
    $script:isLightTheme = $LightTheme
}

function Set-ConnectionStatusTone {
    param([string]$Color)
    $key = 'ThemeColor' + $Color.TrimStart('#').ToUpperInvariant()
    $script:connectionStatusText.Foreground = $script:window.Resources[$key]
}

Set-WindowTheme -LightTheme (Get-SystemLightTheme)
$themeTimer = [System.Windows.Threading.DispatcherTimer]::new()
$themeTimer.Interval = [TimeSpan]::FromSeconds(1)
$themeTimer.Add_Tick({
    $lightTheme = Get-SystemLightTheme
    if ($lightTheme -ne $script:isLightTheme) { Set-WindowTheme -LightTheme $lightTheme }
})
$themeTimer.Start()
$workArea = [System.Windows.SystemParameters]::WorkArea
$window.Height = [Math]::Min($window.Height, $workArea.Height)
$window.Width = [Math]::Min($window.Width, $workArea.Width)

$state = [hashtable]::Synchronized(@{
    Busy = $false; Connected = $false; Done = $false; CompletionHandled = $false
    Status = 'Ready.'; Percent = 0; Results = $null; Error = $null; Operation = ''
    MsalApp = $null; MsalAccount = $null; AccessToken = $null; UserPrincipalName = ''
    OwnerHandle = [IntPtr]::Zero; NativeLibrary = [IntPtr]::Zero
})
$runspace = [runspacefactory]::CreateRunspace()
$runspace.ApartmentState = 'STA'
$runspace.ThreadOptions = 'ReuseThread'
$runspace.Open()
$activePowerShell = $null
$asyncResult = $null
$identifierPattern = '^[^\s@*?,;]+@[^\s@*?,;]+\.[^\s@*?,;]+$'

function Show-ErrorDialog {
    param([string]$Title, [string]$Message)
    [System.Windows.MessageBox]::Show($window, $Message, $Title,
        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
}

function Get-UserEntries {
    # Commas are not separators so 'Smith, Jane' stays one entry.
    @($usersTextBox.Text -split '[\r\n;]+' | ForEach-Object { $_.Trim() } |
        Where-Object { $_ } | Sort-Object -Unique)
}

function Get-ReviewRows {
    if ($reviewGrid.ItemsSource) { @($reviewGrid.ItemsSource) } else { @() }
}

function Set-InputsEnabled {
    param([bool]$Enabled)
    foreach ($control in @($connectButton, $groupTextBox, $usersTextBox, $memberCheckBox, $ownerCheckBox, $bypassCheckBox, $reviewGrid)) {
        $control.IsEnabled = $Enabled
    }
}

function Update-ActionButtons {
    $ready = $state.Connected -and -not $state.Busy
    $checkUsersButton.IsEnabled = $ready -and -not [string]::IsNullOrWhiteSpace($usersTextBox.Text)
    $rows = @(Get-ReviewRows)
    $allResolved = $rows.Count -gt 0 -and @($rows | Where-Object { $null -eq $_.Selected }).Count -eq 0
    $addButton.IsEnabled = $ready -and $allResolved -and
        ($groupTextBox.Text.Trim() -match $identifierPattern) -and
        ([bool]$memberCheckBox.IsChecked -or [bool]$ownerCheckBox.IsChecked)
}

function Start-BackgroundOperation {
    param([string]$Operation, [scriptblock]$Script, [object[]]$Arguments)
    if ($state.Busy) { return }
    $state.Busy = $true
    $state.Done = $false
    $state.CompletionHandled = $false
    $state.Error = $null
    $state.Results = $null
    $state.Operation = $Operation
    $state.Status = switch ($Operation) {
        'Connect' { 'Connecting to Exchange Online...' }
        'Check users' { 'Looking up users...' }
        default { 'Validating group...' }
    }
    $state.Percent = 0
    Set-InputsEnabled -Enabled $false
    Update-ActionButtons
    $operationProgressBar.IsIndeterminate = ($Operation -eq 'Connect')
    if ($Operation -eq 'Connect') {
        $state.Connected = $false
        $connectionStatusText.Text = 'Connecting...'
        Set-ConnectionStatusTone -Color '#FFB454'
    }
    try {
        $script:activePowerShell = [powershell]::Create()
        $script:activePowerShell.Runspace = $runspace
        $null = $script:activePowerShell.AddScript($Script).AddArgument($state)
        foreach ($argument in $Arguments) { $null = $script:activePowerShell.AddArgument($argument) }
        $script:asyncResult = $script:activePowerShell.BeginInvoke()
    }
    catch {
        $state.Error = $_.Exception.Message
        $state.Done = $true
    }
}

$timer = [System.Windows.Threading.DispatcherTimer]::new()
$timer.Interval = [TimeSpan]::FromMilliseconds(150)
$timer.Add_Tick({
    $operationStatusText.Text = [string]$state.Status
    if (-not $operationProgressBar.IsIndeterminate) { $operationProgressBar.Value = [double]$state.Percent }
    if ($state.Busy -and -not $state.CompletionHandled -and
        ($state.Done -or ($script:asyncResult -and $script:asyncResult.IsCompleted))) {
        $state.CompletionHandled = $true
        try {
            if ($script:activePowerShell -and $script:asyncResult) {
                $null = $script:activePowerShell.EndInvoke($script:asyncResult)
                if ($script:activePowerShell.HadErrors -and -not $state.Error) {
                    $state.Error = ($script:activePowerShell.Streams.Error | ForEach-Object { $_.ToString() }) -join "`n"
                }
            }
        }
        catch { if (-not $state.Error) { $state.Error = $_.Exception.Message } }
        finally {
            if ($script:activePowerShell) { $script:activePowerShell.Dispose() }
            $script:activePowerShell = $null
            $script:asyncResult = $null
        }
        $state.Busy = $false
        $operationProgressBar.IsIndeterminate = $false
        Set-InputsEnabled -Enabled $true
        if ($state.Error) {
            $state.Status = "$($state.Operation) failed."
            $operationProgressBar.Value = 0
            if ($state.Operation -eq 'Connect') {
                $connectionStatusText.Text = 'Not connected'
                Set-ConnectionStatusTone -Color '#FF7185'
            }
            Show-ErrorDialog -Title "$($state.Operation) failed" -Message ([string]$state.Error)
        }
        elseif ($state.Operation -eq 'Connect') {
            $state.Connected = $true
            $connectionStatusText.Text = "Connected as $($state.UserPrincipalName)"
            Set-ConnectionStatusTone -Color '#76E0A8'
            $connectButton.Content = 'Reconnect'
            $reviewGrid.ItemsSource = $null
            $state.Status = 'Connected. Enter a group and users, then select Check users.'
        }
        elseif ($state.Operation -eq 'Check users') {
            $rows = @(foreach ($row in @($state.Results)) {
                [GroupUserRow]::new([string]$row.Entry, [string]$row.Status, [object[]]@($row.Options))
            })
            $reviewGrid.ItemsSource = $rows
            $reviewGrid.Visibility = 'Visible'
            $resultsGrid.Visibility = 'Collapsed'
            $operationProgressBar.Value = 100
            $unresolved = @($rows | Where-Object { $null -eq $_.Selected }).Count
            $state.Status = if ($unresolved -eq 0) { "All $($rows.Count) user(s) resolved. Review the matches and select Add roles." }
                else { "$unresolved entry/entries need attention. Choose a match or correct the input and check again." }
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
            $state.Status = "Completed: $($results.Count - $failed - $warnings) added or already present, $warnings warning(s), $failed failed."
            $icon = if ($failed -gt 0 -or $warnings -gt 0) { [System.Windows.MessageBoxImage]::Warning }
                else { [System.Windows.MessageBoxImage]::Information }
            [System.Windows.MessageBox]::Show($window,
                "$($state.Status)`n`nSee the results table for details. Successful changes are not rolled back if another role fails.",
                'Operation complete', [System.Windows.MessageBoxButton]::OK, $icon) | Out-Null
        }
        $operationStatusText.Text = [string]$state.Status
        Update-ActionButtons
    }
})
$timer.Start()

$connectButton.Add_Click({
    $ownerHandle = [System.Windows.Interop.WindowInteropHelper]::new($window).Handle
    Start-BackgroundOperation -Operation 'Connect' -Script $connectScript -Arguments @($ownerHandle)
})
$closeButton.Add_Click({ $window.Close() })
foreach ($checkBox in @($memberCheckBox, $ownerCheckBox, $bypassCheckBox)) {
    $checkBox.Add_Click({ $validationText.Visibility = 'Collapsed'; Update-ActionButtons })
}
$groupTextBox.Add_TextChanged({ $validationText.Visibility = 'Collapsed'; Update-ActionButtons })
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
        else { 'Enter at least one name or email address.' }
        $validationText.Visibility = 'Visible'
        return
    }
    $reviewGrid.ItemsSource = $null
    Start-BackgroundOperation -Operation 'Check users' -Script $resolveScript `
        -Arguments @($entries, $identifierPattern, $ensureConnectionScript)
})
$addButton.Add_Click({
    $group = $groupTextBox.Text.Trim()
    $rows = @(Get-ReviewRows)
    $errors = @()
    if ($group -notmatch $identifierPattern) { $errors += 'Enter a valid group email address.' }
    if ($rows.Count -eq 0 -or @($rows | Where-Object { $null -eq $_.Selected }).Count -gt 0) {
        $errors += 'Every user must be resolved to a single match. Select Check users and choose any ambiguous matches.'
    }
    if (-not ($memberCheckBox.IsChecked -or $ownerCheckBox.IsChecked)) { $errors += 'Select at least one role.' }
    if ($errors.Count -gt 0) {
        $validationText.Text = $errors -join "`n"
        $validationText.Visibility = 'Visible'
        return
    }
    $roleText = @(
        if ($memberCheckBox.IsChecked) { 'Member' }
        if ($ownerCheckBox.IsChecked) { 'Owner' }
    ) -join ' and '
    $users = @($rows | ForEach-Object { $_.Selected } | Group-Object Address | ForEach-Object { $_.Group[0] })
    $message = "Add $roleText on $group for $($users.Count) user(s)?`n`n$(($users | ForEach-Object { $_.Label }) -join "`n")"
    if ($ownerCheckBox.IsChecked) {
        $message += "`n`nFor Microsoft 365 groups, missing membership will be added before ownership. Distribution-list ownership does not add membership."
    }
    if ($bypassCheckBox.IsChecked) { $message += "`n`nThe distribution-list group manager check will be bypassed using your admin permissions." }
    $message += "`n`nSuccessful changes are not rolled back if another role fails."
    $confirm = [System.Windows.MessageBox]::Show($window, $message, 'Confirm changes',
        [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }
    $resultsGrid.ItemsSource = $null
    Start-BackgroundOperation -Operation 'Add roles' -Script $addScript `
        -Arguments @($group, $users, [bool]$memberCheckBox.IsChecked, [bool]$ownerCheckBox.IsChecked,
            [bool]$bypassCheckBox.IsChecked, $ensureConnectionScript)
})
$window.Add_Closing({
    param($windowSender, $windowEventArgs)
    if ($state.Busy) {
        $choice = [System.Windows.MessageBox]::Show($window,
            'An operation is still running. Stop it and close the application? Changes already applied will remain.',
            'Operation in progress', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($choice -ne [System.Windows.MessageBoxResult]::Yes) {
            $windowEventArgs.Cancel = $true
            return
        }
        if ($script:activePowerShell) { $script:activePowerShell.Stop() }
    }
})
$window.Add_Closed({
    $themeTimer.Stop()
    $timer.Stop()
    if ($script:activePowerShell) { $script:activePowerShell.Dispose() }
    $runspace.Close()
    $runspace.Dispose()
})
$null = $window.ShowDialog()
