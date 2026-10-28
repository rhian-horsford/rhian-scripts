#requires -Version 7.6

Set-StrictMode -Version 2.0

$scriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Add-SharedMailboxPermission.ps1'
$tokens = $null
$parseErrors = $null
$script:mailboxAst = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors.Message -join "`n") }

function Get-AssignedExpression {
    param([string]$Name)
    $assignment = $script:mailboxAst.Find({
        param($node)
        $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left.Extent.Text -eq "`$$Name"
    }, $true)
    $assignment.Right.Expression
}

$script:createMailbox = (Get-AssignedExpression 'createScript').ScriptBlock.GetScriptBlock()
$script:grantPermissions = (Get-AssignedExpression 'addScript').ScriptBlock.GetScriptBlock()
$script:emailPattern = (Get-AssignedExpression 'identifierPattern').Value

function New-Mailbox {
    [CmdletBinding()]
    param([switch]$Shared, $Name, $DisplayName, $Alias, $PrimarySmtpAddress)
    throw 'Exchange command must be mocked.'
}
function Get-EXOMailbox {
    [CmdletBinding()]
    param($Identity, $Properties)
    throw 'Exchange command must be mocked.'
}
function Add-MailboxPermission {
    [CmdletBinding(SupportsShouldProcess)]
    param($Identity, $User, $AccessRights, $InheritanceType, $AutoMapping)
    if ($PSCmdlet.ShouldProcess($Identity, 'Test stub')) { throw 'Exchange command must be mocked.' }
}
function Add-RecipientPermission {
    [CmdletBinding(SupportsShouldProcess)]
    param($Identity, $Trustee, $AccessRights)
    if ($PSCmdlet.ShouldProcess($Identity, 'Test stub')) { throw 'Exchange command must be mocked.' }
}

Describe 'Shared mailbox creation and permissions' {
    BeforeEach {
        $script:state = @{ CreatedMailbox = ''; Error = $null; Results = @(); Done = $false; Status = ''; Percent = 0 }
        $script:users = @([pscustomobject]@{ Address = 'jane@example.com'; Label = 'Jane <jane@example.com>' })
        $script:calls = [System.Collections.Generic.List[string]]::new()
        $script:failCreate = $false
        $script:failLookup = $false
        $script:failFullAccess = $false
        $script:warnCreate = $false
        Mock New-Mailbox {
            $script:calls.Add('Create')
            if ($script:failCreate) { throw 'Address already in use' }
            if ($script:warnCreate) { Write-Warning 'Provisioning warning' }
        }
        Mock Get-EXOMailbox {
            if ($script:failLookup) { throw 'Mailbox not yet available' }
            [pscustomobject]@{ RecipientTypeDetails = 'SharedMailbox'; PrimarySmtpAddress = 'team@example.com' }
        }
        Mock Add-MailboxPermission {
            if ($script:state.Done) { throw 'Worker completed too early' }
            $script:calls.Add('FullAccess')
            if ($script:failFullAccess) { throw 'Full Access denied' }
        }
        Mock Add-RecipientPermission { $script:calls.Add('SendAs') }
    }

    It 'creates a shared mailbox before granting both permissions' {
        & $script:createMailbox $script:state 'Team' 'team' 'team@example.com' $script:users $true $true $true {} $script:grantPermissions $script:emailPattern
        $script:state.Error | Should BeNullOrEmpty
        ($script:calls -join ',') | Should Be 'Create,FullAccess,SendAs'
        $script:state.CreatedMailbox | Should Be 'team@example.com'
        $script:state.Results.Count | Should Be 3
        @($script:state.Results | Where-Object Result -eq 'Failed').Count | Should Be 0
        $script:state.Done | Should Be $true
        Assert-MockCalled New-Mailbox -Times 1 -Exactly -Scope It -ParameterFilter {
            $Shared -and $Name -eq 'Team' -and $DisplayName -eq 'Team' -and $Alias -eq 'team' -and $PrimarySmtpAddress -eq 'team@example.com'
        }
        Assert-MockCalled Add-MailboxPermission -Times 1 -Exactly -Scope It -ParameterFilter {
            $Identity -eq 'team@example.com' -and $User -eq 'jane@example.com' -and
            $AccessRights -eq 'FullAccess' -and $InheritanceType -eq 'All' -and $AutoMapping
        }
        Assert-MockCalled Add-RecipientPermission -Times 1 -Exactly -Scope It -ParameterFilter {
            $Identity -eq 'team@example.com' -and $Trustee -eq 'jane@example.com' -and $AccessRights -eq 'SendAs'
        }
    }

    It 'supports Send As only without Full Access' {
        & $script:createMailbox $script:state 'Team' 'team' 'team@example.com' $script:users $false $true $true {} $script:grantPermissions $script:emailPattern
        ($script:calls -join ',') | Should Be 'Create,SendAs'
        $script:state.Results.Count | Should Be 2
    }

    It 'supports Full Access only with auto-mapping disabled' {
        & $script:createMailbox $script:state 'Team' 'team' 'team@example.com' $script:users $true $false $false {} $script:grantPermissions $script:emailPattern
        ($script:calls -join ',') | Should Be 'Create,FullAccess'
        Assert-MockCalled Add-MailboxPermission -Times 1 -Exactly -Scope It -ParameterFilter { -not $AutoMapping }
    }

    It 'does not grant permissions when creation fails or conflicts' {
        $script:failCreate = $true
        & $script:createMailbox $script:state 'Team' 'team' 'team@example.com' $script:users $true $true $true {} $script:grantPermissions $script:emailPattern
        $script:state.Error | Should Match 'Address already in use'
        $script:state.CreatedMailbox | Should BeNullOrEmpty
        ($script:calls -join ',') | Should Be 'Create'
        Assert-MockCalled Get-EXOMailbox -Times 0 -Exactly -Scope It
    }

    It 'preserves creation and Send As success when Full Access fails' {
        $script:failFullAccess = $true
        & $script:createMailbox $script:state 'Team' 'team' 'team@example.com' $script:users $true $true $true {} $script:grantPermissions $script:emailPattern
        $script:state.Error | Should BeNullOrEmpty
        $script:state.Results[0].Result | Should Be 'Success'
        $script:state.Results[1].Result | Should Be 'Failed'
        $script:state.Results[2].Result | Should Be 'Success'
        $script:state.CreatedMailbox | Should Be 'team@example.com'
    }

    It 'reports provisioning lookup failures without hiding mailbox creation' {
        $script:failLookup = $true
        & $script:createMailbox $script:state 'Team' 'team' 'team@example.com' $script:users $true $true $true {} $script:grantPermissions $script:emailPattern
        $script:state.Error | Should BeNullOrEmpty
        $script:state.Results[0].Permission | Should Be 'Create mailbox'
        $script:state.Results[1].Result | Should Be 'Failed'
        $script:state.Results[1].Details | Should Match 'not yet available'
        ($script:calls -join ',') | Should Be 'Create'
    }

    It 'blocks repeat creation after success' {
        $script:state.CreatedMailbox = 'team@example.com'
        & $script:createMailbox $script:state 'Team' 'team' 'team@example.com' $script:users $true $true $true {} $script:grantPermissions $script:emailPattern
        $script:state.Error | Should Match 'already created'
        $script:calls.Count | Should Be 0
    }

    It 'rejects invalid aliases before creating anything' {
        & $script:createMailbox $script:state 'Team' 'bad alias' 'team@example.com' $script:users $true $true $true {} $script:grantPermissions $script:emailPattern
        $script:state.Error | Should Match 'alias'
        $script:calls.Count | Should Be 0
    }

    It 'requires a resolved user and at least one permission before creation' {
        & $script:createMailbox $script:state 'Team' 'team' 'team@example.com' @() $true $true $true {} $script:grantPermissions $script:emailPattern
        $script:state.Error | Should Match 'resolve at least one user'
        & $script:createMailbox $script:state 'Team' 'team' 'team@example.com' $script:users $false $false $true {} $script:grantPermissions $script:emailPattern
        $script:state.Error | Should Match 'Select at least one permission'
        $script:calls.Count | Should Be 0
    }

    It 'does not create anything when sign-in fails' {
        & $script:createMailbox $script:state 'Team' 'team' 'team@example.com' $script:users $true $true $true { throw 'Sign-in cancelled' } $script:grantPermissions $script:emailPattern
        $script:state.Error | Should Match 'Sign-in cancelled'
        $script:calls.Count | Should Be 0
    }

    It 'retains Exchange creation warnings in the results' {
        $script:warnCreate = $true
        & $script:createMailbox $script:state 'Team' 'team' 'team@example.com' $script:users $false $true $true {} $script:grantPermissions $script:emailPattern
        $script:state.Results[0].Result | Should Be 'Warning'
        $script:state.Results[0].Details | Should Match 'Provisioning warning'
    }

    It 'preserves standalone permission worker completion' {
        & $script:grantPermissions $script:state 'team@example.com' $script:users $true $false $true {}
        $script:state.Done | Should Be $true
        $script:state.Error | Should BeNullOrEmpty
        ($script:calls -join ',') | Should Be 'FullAccess'
    }

    It 'loads creation fields in the existing WPF layout' {
        Add-Type -AssemblyName PresentationFramework
        $window = [Windows.Markup.XamlReader]::Parse((Get-AssignedExpression 'xaml').Value)
        try {
            foreach ($name in 'CreationFields', 'NameTextBox', 'AliasTextBox', 'HeadingText') {
                ($null -ne $window.FindName($name)) | Should Be $true
            }
            $window.FindName('CreationFields').Visibility.ToString() | Should Be 'Collapsed'
        }
        finally { $window.Close() }
    }

    It 'opens the creation launcher with the expected controls and closes without sign-in' {
        Add-Type -AssemblyName PresentationFramework
        $probe = @{ Seen = $false; Error = $null }
        $timer = [Windows.Threading.DispatcherTimer]::new()
        $timer.Interval = [TimeSpan]::FromMilliseconds(150)
        $timer.Add_Tick({
            foreach ($source in @([Windows.PresentationSource]::CurrentSources)) {
                $dialog = $source.RootVisual
                if ($dialog -isnot [Windows.Window] -or $dialog.Title -ne 'Create Shared Mailbox') { continue }
                try {
                    $probe.Seen = $true
                    if ($dialog.FindName('CreationFields').Visibility -ne 'Visible') { throw 'Creation fields hidden' }
                    if ($dialog.FindName('RemoveButton').Visibility -ne 'Collapsed') { throw 'Removal offered during creation' }
                    $dialog.FindName('NameTextBox').Text = 'Team'
                    $dialog.FindName('AliasTextBox').Text = 'team'
                    $dialog.FindName('MailboxTextBox').Text = 'team@example.com'
                    if ($dialog.FindName('AddButton').IsEnabled) { throw 'Creation allowed before sign-in and user resolution' }
                    if ($dialog.FindName('AddButton').Content -ne 'Create mailbox') { throw 'Wrong primary action' }
                }
                catch { $probe.Error = $_.Exception.Message }
                finally { $timer.Stop(); $dialog.Close() }
            }
        }.GetNewClosure())
        $timer.Start()
        try { & (Join-Path (Split-Path $PSScriptRoot -Parent) 'New-SharedMailbox.ps1') }
        finally { $timer.Stop() }
        $probe.Seen | Should Be $true
        $probe.Error | Should BeNullOrEmpty
    }

    It 'keeps the normal permission-manager mode unchanged' {
        Add-Type -AssemblyName PresentationFramework
        $probe = @{ Seen = $false; Error = $null }
        $timer = [Windows.Threading.DispatcherTimer]::new()
        $timer.Interval = [TimeSpan]::FromMilliseconds(150)
        $timer.Add_Tick({
            foreach ($source in @([Windows.PresentationSource]::CurrentSources)) {
                $dialog = $source.RootVisual
                if ($dialog -isnot [Windows.Window] -or $dialog.Title -ne 'Shared Mailbox Permission Manager') { continue }
                try {
                    $probe.Seen = $true
                    if ($dialog.FindName('CreationFields').Visibility -ne 'Collapsed') { throw 'Creation fields shown in permissions mode' }
                    if ($dialog.FindName('RemoveButton').Visibility -ne 'Visible') { throw 'Remove button hidden' }
                    $dialog.FindName('MailboxTextBox').Text = 'team@example.com'
                    if ($dialog.FindName('AddButton').Content -ne 'Add permissions') { throw 'Normal action changed' }
                }
                catch { $probe.Error = $_.Exception.Message }
                finally { $timer.Stop(); $dialog.Close() }
            }
        }.GetNewClosure())
        $timer.Start()
        try { & (Join-Path (Split-Path $PSScriptRoot -Parent) 'Add-SharedMailboxPermission.ps1') }
        finally { $timer.Stop() }
        $probe.Seen | Should Be $true
        $probe.Error | Should BeNullOrEmpty
    }
}