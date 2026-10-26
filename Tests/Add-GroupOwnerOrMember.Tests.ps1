#requires -Version 7.6

Set-StrictMode -Version 2.0

$scriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Add-GroupOwnerOrMember.ps1'
$tokens = $null
$parseErrors = $null
$script:groupAst = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors.Message -join "`n") }

function Get-AssignedExpression {
    param([string]$Name)
    $assignment = $script:groupAst.Find({
        param($node)
        $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left.Extent.Text -eq "`$$Name"
    }, $true)
    $assignment.Right.Expression
}

$script:addRoles = (Get-AssignedExpression 'addScript').ScriptBlock.GetScriptBlock()
$script:resolveUsers = (Get-AssignedExpression 'resolveScript').ScriptBlock.GetScriptBlock()
$script:emailPattern = (Get-AssignedExpression 'identifierPattern').Value
$script:importCsvUsers = $script:groupAst.Find({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-CsvUserEntries'
}, $true).Body.GetScriptBlock()

# Local command signatures allow mocking without ExchangeOnlineManagement or a tenant.
function Get-Recipient {
    [CmdletBinding()]
    param($Identity, $Anr, $RecipientTypeDetails, $ResultSize)
}
function Get-EXORecipient {
    [CmdletBinding()]
    param($Identity)
}
function Get-UnifiedGroup {
    [CmdletBinding()]
    param($Identity)
}
function Get-DistributionGroup {
    [CmdletBinding()]
    param($Identity)
}
function Get-UnifiedGroupLinks {
    [CmdletBinding()]
    param($Identity, $LinkType, $ResultSize)
}
function Get-DistributionGroupMember {
    [CmdletBinding()]
    param($Identity, $ResultSize)
}
function Add-UnifiedGroupLinks {
    [CmdletBinding(SupportsShouldProcess)]
    param($Identity, $LinkType, $Links)
    if ($PSCmdlet.ShouldProcess($Identity, 'Test stub')) { throw 'Exchange command must be mocked.' }
}
function Add-DistributionGroupMember {
    [CmdletBinding(SupportsShouldProcess)]
    param($Identity, $Member, [switch]$BypassSecurityGroupManagerCheck)
    if ($PSCmdlet.ShouldProcess($Identity, 'Test stub')) { throw 'Exchange command must be mocked.' }
}
function Set-DistributionGroup {
    [CmdletBinding(SupportsShouldProcess)]
    param($Identity, $ManagedBy, [switch]$BypassSecurityGroupManagerCheck)
    if ($PSCmdlet.ShouldProcess($Identity, 'Test stub')) { throw 'Exchange command must be mocked.' }
}
function Remove-UnifiedGroupLinks {
    [CmdletBinding(SupportsShouldProcess)]
    param($Identity, $LinkType, $Links)
    if ($PSCmdlet.ShouldProcess($Identity, 'Test stub')) { throw 'Exchange command must be mocked.' }
}
function Remove-DistributionGroupMember {
    [CmdletBinding(SupportsShouldProcess)]
    param($Identity, $Member, [switch]$BypassSecurityGroupManagerCheck)
    if ($PSCmdlet.ShouldProcess($Identity, 'Test stub')) { throw 'Exchange command must be mocked.' }
}

Describe 'Standalone group tool' {
    BeforeEach {
        $script:state = @{ Error = $null; Results = @(); Done = $false; Status = ''; Percent = 0 }
        $script:user = [pscustomobject]@{ Address = 'jane@example.com'; Label = 'Jane <jane@example.com>' }
        $script:groupType = 'GroupMailbox'
        $script:synced = $false
        $script:members = @()
        $script:owners = @()
        $script:failMember = $false
        $script:failOwner = $false
        $script:failReads = $false
        $script:warnOnWrite = $false
        $script:failAddress = ''
        $script:dlOwners = @('existing-owner')
        $script:calls = [System.Collections.Generic.List[string]]::new()
        Mock Get-Recipient {
            if ($Identity -eq 'existing-owner') {
                return [pscustomobject]@{ PrimarySmtpAddress = 'existing@example.com' }
            }
            [pscustomobject]@{ PrimarySmtpAddress = 'team@example.com'; RecipientTypeDetails = $script:groupType }
        }
        Mock Get-UnifiedGroup { [pscustomobject]@{ IsDirSynced = $script:synced } }
        Mock Get-DistributionGroup { [pscustomobject]@{ IsDirSynced = $script:synced; ManagedBy = $script:dlOwners } }
        Mock Get-UnifiedGroupLinks {
            if ($script:failReads) { throw 'Read denied' }
            if ($LinkType -eq 'Members') { $script:members } else { $script:owners }
        }
        Mock Get-DistributionGroupMember { $script:members }
        Mock Add-UnifiedGroupLinks {
            if ($script:failMember -and $LinkType -eq 'Members') { throw 'Membership denied' }
            if ($script:failOwner -and $LinkType -eq 'Owners') { throw 'Ownership denied' }
            if ($Links -eq $script:failAddress) { throw 'User denied' }
            if ($script:warnOnWrite) { Write-Warning 'Review this change' }
            $script:calls.Add("$LinkType`:$Links")
        }
        Mock Add-DistributionGroupMember { $script:calls.Add("Member:$Member") }
        Mock Set-DistributionGroup {
            if ($ManagedBy.ContainsKey('Remove')) { $script:calls.Add("RemoveOwner:$($ManagedBy.Remove)") }
            else { $script:calls.Add("Owner:$($ManagedBy.Add)") }
        }
        Mock Remove-UnifiedGroupLinks {
            if ($script:failReads) { throw 'Read denied' }
            if ($script:failOwner -and $LinkType -eq 'Owners') { throw 'Ownership denied' }
            if ($script:failMember -and $LinkType -eq 'Members') { throw 'Membership denied' }
            if ($script:warnOnWrite) { Write-Warning 'Review this removal' }
            $script:calls.Add("Remove$LinkType`:$Links")
        }
        Mock Remove-DistributionGroupMember { $script:calls.Add("RemoveMember:$Member") }
    }

    It 'skips absent roles without issuing removal commands' {
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $true $false {} Remove
        $script:state.Error | Should BeNullOrEmpty
        @($script:state.Results | Where-Object Result -eq 'Not present').Count | Should Be 2
        $script:calls.Count | Should Be 0
        Assert-MockCalled Remove-UnifiedGroupLinks -Times 0 -Exactly -Scope It
    }

    It 'removes M365 ownership before membership even when only Member is selected' {
        $script:members = @([pscustomobject]@{ PrimarySmtpAddress = 'jane@example.com' })
        $script:owners = @(
            [pscustomobject]@{ PrimarySmtpAddress = 'JANE@example.com' }
            [pscustomobject]@{ PrimarySmtpAddress = 'other@example.com' }
        )
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $false $false {} Remove
        ($script:calls -join ',') | Should Be 'RemoveOwners:jane@example.com,RemoveMembers:jane@example.com'
        $script:state.Results[0].Role | Should Be 'Owner'
        $script:state.Results[0].Details | Should Match 'Checked before'
        $script:state.Results[1].Details | Should Be 'Removed.'
        $script:state.Percent | Should Be 100
    }

    It 'removes membership for a non-owner without an owner write' {
        $script:members = @([pscustomobject]@{ PrimarySmtpAddress = 'jane@example.com' })
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $false $false {} Remove
        ($script:calls -join ',') | Should Be 'RemoveMembers:jane@example.com'
        $script:state.Results[0].Result | Should Be 'Not present'
    }

    It 'leaves membership unchanged when only Owner is removed' {
        $script:owners = @(
            [pscustomobject]@{ PrimarySmtpAddress = 'jane@example.com' }
            [pscustomobject]@{ PrimarySmtpAddress = 'other@example.com' }
        )
        & $script:addRoles $script:state 'team@example.com' @($script:user) $false $true $false {} Remove
        ($script:calls -join ',') | Should Be 'RemoveOwners:jane@example.com'
        $script:state.Results.Count | Should Be 1
        Assert-MockCalled Get-UnifiedGroupLinks -Times 0 -Exactly -Scope It -ParameterFilter { $LinkType -eq 'Members' }
    }

    It 'blocks membership removal when prerequisite ownership removal fails' {
        $script:members = @([pscustomobject]@{ PrimarySmtpAddress = 'jane@example.com' })
        $script:owners = @(
            [pscustomobject]@{ PrimarySmtpAddress = 'jane@example.com' }
            [pscustomobject]@{ PrimarySmtpAddress = 'other@example.com' }
        )
        $script:failOwner = $true
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $false $false {} Remove
        $script:state.Results[0].Result | Should Be 'Failed'
        $script:state.Results[1].Details | Should Match 'ownership could not be removed'
        Assert-MockCalled Remove-UnifiedGroupLinks -Times 0 -Exactly -Scope It -ParameterFilter { $LinkType -eq 'Members' }
    }

    It 'preserves the last owner and its membership' {
        $script:members = @([pscustomobject]@{ PrimarySmtpAddress = 'jane@example.com' })
        $script:owners = $script:members
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $true $false {} Remove
        $script:state.Results[0].Details | Should Match 'last owner'
        $script:state.Results[1].Result | Should Be 'Failed'
        $script:calls.Count | Should Be 0
    }

    It 'keeps the final owner when multiple owners are selected for removal' {
        $second = [pscustomobject]@{ Address = 'other@example.com'; Label = 'Other' }
        $script:owners = @(
            [pscustomobject]@{ PrimarySmtpAddress = 'jane@example.com' }
            [pscustomobject]@{ PrimarySmtpAddress = 'other@example.com' }
        )
        & $script:addRoles $script:state 'team@example.com' @($script:user, $second) $false $true $false {} Remove
        $script:state.Results[0].Result | Should Be 'Success'
        $script:state.Results[1].Details | Should Match 'last owner'
        $script:calls.Count | Should Be 1
    }

    It 'removes DL membership with explicit admin override' {
        $script:groupType = 'MailUniversalDistributionGroup'
        $script:members = @([pscustomobject]@{ PrimarySmtpAddress = 'jane@example.com' })
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $false $true {} Remove
        ($script:calls -join ',') | Should Be 'RemoveMember:jane@example.com'
        Assert-MockCalled Remove-DistributionGroupMember -Times 1 -Exactly -Scope It -ParameterFilter { $BypassSecurityGroupManagerCheck }
        Assert-MockCalled Set-DistributionGroup -Times 0 -Exactly -Scope It
    }

    It 'uses a subtractive DL owner update without touching membership or other owners' {
        $script:groupType = 'MailUniversalSecurityGroup'
        $script:dlOwners = @('existing-owner', 'team@example.com')
        $script:user = [pscustomobject]@{ Address = 'existing@example.com'; Label = 'Existing' }
        & $script:addRoles $script:state 'team@example.com' @($script:user) $false $true $true {} Remove
        $script:state.Error | Should BeNullOrEmpty
        ($script:calls -join ',') | Should Be 'RemoveOwner:existing@example.com'
        Assert-MockCalled Set-DistributionGroup -Times 1 -Exactly -Scope It -ParameterFilter {
            $ManagedBy.Count -eq 1 -and $ManagedBy.Remove -eq 'existing@example.com' -and $BypassSecurityGroupManagerCheck
        }
        Assert-MockCalled Remove-DistributionGroupMember -Times 0 -Exactly -Scope It
    }

    It 'blocks removal of the last DL owner' {
        $script:groupType = 'MailUniversalDistributionGroup'
        $script:user = [pscustomobject]@{ Address = 'existing@example.com'; Label = 'Existing' }
        & $script:addRoles $script:state 'team@example.com' @($script:user) $false $true $false {} Remove
        $script:state.Results[0].Details | Should Match 'last owner'
        $script:calls.Count | Should Be 0
    }

    It 'does not remove anything if the existing role lookup fails' {
        $script:failReads = $true
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $true $false {} Remove
        $script:state.Error | Should Be 'Read denied'
        $script:calls.Count | Should Be 0
    }

    It 'reports partial removal failures without rolling back successful removals' {
        $script:members = @([pscustomobject]@{ PrimarySmtpAddress = 'jane@example.com' })
        $script:owners = @(
            [pscustomobject]@{ PrimarySmtpAddress = 'jane@example.com' }
            [pscustomobject]@{ PrimarySmtpAddress = 'other@example.com' }
        )
        $script:failMember = $true
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $true $false {} Remove
        $script:state.Results[0].Result | Should Be 'Success'
        $script:state.Results[1].Result | Should Be 'Failed'
        ($script:calls -join ',') | Should Be 'RemoveOwners:jane@example.com'
    }

    It 'surfaces removal warnings' {
        $script:members = @([pscustomobject]@{ PrimarySmtpAddress = 'jane@example.com' })
        $script:warnOnWrite = $true
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $false $false {} Remove
        $script:state.Results[1].Result | Should Be 'Warning'
        $script:state.Results[1].Details | Should Match 'Review this removal'
    }

    It 'adds required M365 membership before ownership and reports both roles' {
        & $script:addRoles $script:state 'team@example.com' @($script:user) $false $true $false {}
        $script:state.Error | Should BeNullOrEmpty
        $script:state.Done | Should Be $true
        ($script:calls -join ',') | Should Be 'Members:jane@example.com,Owners:jane@example.com'
        $script:state.Results.Count | Should Be 2
        $script:state.Results[0].Details | Should Match 'Required for Microsoft 365 ownership'
        $script:state.Percent | Should Be 100
    }

    It 'does not add ownership if required membership fails' {
        $script:failMember = $true
        & $script:addRoles $script:state 'team@example.com' @($script:user) $false $true $false {}
        $script:state.Results[0].Result | Should Be 'Failed'
        $script:state.Results[1].Result | Should Be 'Failed'
        $script:state.Results[1].Details | Should Match 'required membership'
        Assert-MockCalled Add-UnifiedGroupLinks -Times 0 -Exactly -Scope It -ParameterFilter { $LinkType -eq 'Owners' }
    }

    It 'preserves a successful membership result when ownership fails' {
        $script:failOwner = $true
        & $script:addRoles $script:state 'team@example.com' @($script:user) $false $true $false {}
        $script:state.Results[0].Result | Should Be 'Success'
        $script:state.Results[1].Result | Should Be 'Failed'
        $script:state.Results[1].Details | Should Be 'Ownership denied'
    }

    It 'skips case-insensitive existing membership and ownership' {
        $script:members = @([pscustomobject]@{ PrimarySmtpAddress = 'JANE@example.com' })
        $script:owners = @([pscustomobject]@{ PrimarySmtpAddress = 'JANE@example.com' })
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $true $false {}
        @($script:state.Results | Where-Object Result -eq 'Already present').Count | Should Be 2
        Assert-MockCalled Add-UnifiedGroupLinks -Times 0 -Exactly -Scope It
    }

    It 'adds only membership when Member is selected' {
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $false $false {}
        $script:state.Results.Count | Should Be 1
        ($script:calls -join ',') | Should Be 'Members:jane@example.com'
        Assert-MockCalled Get-UnifiedGroupLinks -Times 0 -Exactly -Scope It -ParameterFilter { $LinkType -eq 'Owners' }
    }

    It 'requests complete membership and ownership snapshots' {
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $true $false {}
        Assert-MockCalled Get-UnifiedGroupLinks -Times 2 -Exactly -Scope It -ParameterFilter { $ResultSize -eq 'Unlimited' }
    }

    It 'rejects a request with no selected role' {
        & $script:addRoles $script:state 'team@example.com' @($script:user) $false $false $false {}
        $script:state.Error | Should Match 'Select at least one role'
        $script:calls.Count | Should Be 0
    }

    It 'uses an additive owner update without adding DL membership' {
        $script:groupType = 'MailUniversalDistributionGroup'
        & $script:addRoles $script:state 'team@example.com' @($script:user) $false $true $false {}
        $script:state.Error | Should BeNullOrEmpty
        $script:state.Results.Count | Should Be 1
        ($script:calls -join ',') | Should Be 'Owner:jane@example.com'
        Assert-MockCalled Set-DistributionGroup -Times 1 -Exactly -Scope It -ParameterFilter {
            $ManagedBy.Count -eq 1 -and $ManagedBy.Add -eq 'jane@example.com' -and -not $BypassSecurityGroupManagerCheck
        }
        Assert-MockCalled Add-DistributionGroupMember -Times 0 -Exactly -Scope It
    }

    It 'skips an existing DL owner resolved from ManagedBy' {
        $script:groupType = 'MailUniversalDistributionGroup'
        $script:user = [pscustomobject]@{ Address = 'existing@example.com'; Label = 'Existing' }
        & $script:addRoles $script:state 'team@example.com' @($script:user) $false $true $false {}
        $script:state.Results[0].Result | Should Be 'Already present'
        Assert-MockCalled Set-DistributionGroup -Times 0 -Exactly -Scope It
    }

    It 'supports mail-enabled security groups and explicit manager override' {
        $script:groupType = 'MailUniversalSecurityGroup'
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $true $true {}
        $script:state.Error | Should BeNullOrEmpty
        Assert-MockCalled Add-DistributionGroupMember -Times 1 -Exactly -Scope It -ParameterFilter { $BypassSecurityGroupManagerCheck }
        Assert-MockCalled Set-DistributionGroup -Times 1 -Exactly -Scope It -ParameterFilter { $BypassSecurityGroupManagerCheck }
    }

    It 'rejects dynamic groups before any write' {
        $script:groupType = 'DynamicDistributionGroup'
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $true $false {}
        $script:state.Error | Should Match 'not a supported static distribution list'
        $script:calls.Count | Should Be 0
    }

    It 'rejects synchronized groups before any write' {
        $script:synced = $true
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $true $false {}
        $script:state.Error | Should Match 'source directory'
        $script:calls.Count | Should Be 0
    }

    It 'does not mutate when the existing role snapshot fails' {
        $script:failReads = $true
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $true $false {}
        $script:state.Error | Should Be 'Read denied'
        $script:calls.Count | Should Be 0
    }

    It 'surfaces Exchange warnings rather than reporting unqualified success' {
        $script:warnOnWrite = $true
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $false $false {}
        $script:state.Results[0].Result | Should Be 'Warning'
        $script:state.Results[0].Details | Should Match 'Review this change'
    }

    It 'continues to the next user after a failed write' {
        $script:failAddress = 'jane@example.com'
        $second = [pscustomobject]@{ Address = 'john@example.com'; Label = 'John' }
        & $script:addRoles $script:state 'team@example.com' @($script:user, $second) $true $false $false {}
        $script:state.Results[0].Result | Should Be 'Failed'
        $script:state.Results[1].Result | Should Be 'Success'
    }

    It 'reports connection failures without writing' {
        & $script:addRoles $script:state 'team@example.com' @($script:user) $true $false $false { throw 'Sign-in cancelled.' }
        $script:state.Error | Should Be 'Sign-in cancelled.'
        $script:state.Done | Should Be $true
        $script:calls.Count | Should Be 0
    }

    It 'resolves an exact user mailbox address' {
        Mock Get-EXORecipient {
            [pscustomobject]@{ RecipientTypeDetails = 'UserMailbox'; PrimarySmtpAddress = 'jane@example.com'; DisplayName = 'Jane' }
        }
        & $script:resolveUsers $script:state @('jane@example.com') $script:emailPattern {}
        $script:state.Results[0].Status | Should Be 'Resolved'
        $script:state.Results[0].Options[0].Address | Should Be 'jane@example.com'
    }

    It 'does not offer contacts or guests as user mailboxes' {
        Mock Get-EXORecipient { [pscustomobject]@{ RecipientTypeDetails = 'MailContact' } }
        & $script:resolveUsers $script:state @('contact@example.com') $script:emailPattern {}
        $script:state.Results[0].Status | Should Match 'Not a user mailbox'
        $script:state.Results[0].Options.Count | Should Be 0
    }

    It 'keeps ambiguous name matches for explicit selection and filters unrelated names' {
        Mock Get-Recipient {
            @(
                [pscustomobject]@{ DisplayName = 'Jane Smith'; FirstName = 'Jane'; LastName = 'Smith'; Alias = 'jsmith'; PrimarySmtpAddress = 'jane@example.com' }
                [pscustomobject]@{ DisplayName = 'Janet Smith'; FirstName = 'Janet'; LastName = 'Smith'; Alias = 'jansmith'; PrimarySmtpAddress = 'janet@example.com' }
                [pscustomobject]@{ DisplayName = 'John Jones'; FirstName = 'John'; LastName = 'Jones'; Alias = 'jjones'; PrimarySmtpAddress = 'john@example.com' }
            )
        }
        & $script:resolveUsers $script:state @('Smith, Jan') $script:emailPattern {}
        $script:state.Results[0].Status | Should Be 'Choose a match (2 found)'
        $script:state.Results[0].Options.Count | Should Be 2
    }

    It 'shows failed lookups as unresolved rows' {
        Mock Get-EXORecipient { throw 'Recipient not found' }
        & $script:resolveUsers $script:state @('missing@example.com') $script:emailPattern {}
        $script:state.Results[0].Status | Should Match 'Lookup failed: Recipient not found'
        $script:state.Results[0].Options.Count | Should Be 0
    }

    It 'requires a more specific name if too many matches are returned' {
        Mock Get-Recipient {
            1..26 | ForEach-Object {
                [pscustomobject]@{ DisplayName = "Jane $_"; FirstName = 'Jane'; LastName = "$_"; Alias = "jane$_"; PrimarySmtpAddress = "jane$_@example.com" }
            }
        }
        & $script:resolveUsers $script:state @('Jane') $script:emailPattern {}
        $script:state.Results[0].Status | Should Be 'Too many matches - be more specific'
        $script:state.Results[0].Options.Count | Should Be 0
    }

    It 'splits users by newline or semicolon but preserves commas' {
        $function = $script:groupAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-UserEntries'
        }, $true)
        $script:usersTextBox = [pscustomobject]@{ Text = "Smith, Jane; john@example.com`njohn@example.com" }
        $entries = @(& $function.Body.GetScriptBlock())
        $entries.Count | Should Be 2
        ($entries -contains 'Smith, Jane') | Should Be $true
    }

    It 'imports and deduplicates PrimarySmtpAddress values from the exporter CSV' {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) "$([guid]::NewGuid().ToString('N')).csv"
        @(
            'DisplayName,PrimarySmtpAddress,MembershipTypes'
            'Jane Smith,jane@example.com,Member'
            'Jane Smith,jane@example.com,Owner'
            'John Jones,john@example.com,Member'
        ) | Set-Content -LiteralPath $path -Encoding utf8
        try {
            $entries = @(& $script:importCsvUsers -Path $path)
            $entries.Count | Should Be 2
            ($entries -contains 'jane@example.com') | Should Be $true
            ($entries -contains 'john@example.com') | Should Be $true
        }
        finally { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
    }

    It 'rejects CSV files without a supported email column' {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) "$([guid]::NewGuid().ToString('N')).csv"
        @('DisplayName,Department', 'Jane,Sales') | Set-Content -LiteralPath $path -Encoding utf8
        try {
            { & $script:importCsvUsers -Path $path } | Should Throw
        }
        finally { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
    }

    It 'loads the themed WPF window and all named controls without connecting' {
        Add-Type -AssemblyName PresentationFramework
        $xaml = (Get-AssignedExpression 'xaml').Value
        $colors = @([regex]::Matches($xaml, '#[0-9A-Fa-f]{6}') | ForEach-Object { $_.Value } | Sort-Object -Unique)
        $brushes = ''
        foreach ($color in $colors) {
            $key = 'ThemeColor' + $color.Substring(1).ToUpperInvariant()
            $brushes += "<SolidColorBrush x:Key=`"$key`" Color=`"$color`"/>"
            $xaml = $xaml.Replace($color, "{DynamicResource $key}")
        }
        $xaml = $xaml.Replace('<Window.Resources>', "<Window.Resources>$brushes")
        $reader = [System.Xml.XmlNodeReader]::new([xml]$xaml)
        $window = [Windows.Markup.XamlReader]::Load($reader)
        try {
            $window.Title | Should Be 'Group Owner and Member Manager'
            $namedControls = $script:groupAst.FindAll({
                param($node)
                $node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Member.Value -eq 'FindName'
            }, $true)
            foreach ($control in $namedControls) {
                $window.FindName($control.Arguments[0].Value) | Should Not BeNullOrEmpty
            }
            $window.FindName('MemberCheckBox').IsChecked | Should Be $true
            $window.FindName('OwnerCheckBox').IsChecked | Should Be $false
            $window.FindName('BypassCheckBox').IsChecked | Should Be $false
            $window.FindName('AddButton').IsEnabled | Should Be $false
            $window.FindName('RemoveButton').IsEnabled | Should Be $false
            $themeAssignment = $script:groupAst.Find({
                param($node)
                $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$lightThemeColors'
            }, $true)
            $lightThemeColors = & ([scriptblock]::Create($themeAssignment.Right.Extent.Text))
            $themeFunction = $script:groupAst.Find({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Set-WindowTheme'
            }, $true)
            $script:window = $window
            $script:themeHexColors = $colors
            $script:lightThemeColors = $lightThemeColors
            & $themeFunction.Body.GetScriptBlock() $true
            $window.Resources['ThemeColor131417'].Color.ToString() | Should Be '#FFF4F6FA'
            & $themeFunction.Body.GetScriptBlock() $false
            $window.Resources['ThemeColor131417'].Color.ToString() | Should Be '#FF131417'
        }
        finally { $window.Close() }
    }
}
