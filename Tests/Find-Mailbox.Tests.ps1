#requires -Version 7.6

Set-StrictMode -Version 2.0

$script:lookupPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Find-Mailbox.ps1'
$tokens = $null
$parseErrors = $null
$script:lookupAst = [System.Management.Automation.Language.Parser]::ParseFile($script:lookupPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors.Message -join "`n") }
$script:guiAst = $script:lookupAst.Find({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Show-MailboxLookupWindow'
}, $true)

function Get-ConnectionInformation {
    [CmdletBinding()]
    param()
}
function Connect-ExchangeOnline {
    [CmdletBinding()]
    param([bool]$ShowBanner, [switch]$DisableWAM)
}
function Get-EXOMailbox {
    [CmdletBinding()]
    param($Filter, $ResultSize, $Properties)
}

Describe 'Mailbox lookup' {
    BeforeEach {
        Mock Import-Module {} -ParameterFilter { $Name -eq 'ExchangeOnlineManagement' }
        Mock Get-ConnectionInformation { [pscustomobject]@{ State = 'Connected' } }
        Mock Connect-ExchangeOnline {}
        Mock Write-Host {}
        Mock Get-EXOMailbox {
            [pscustomobject]@{
                DisplayName = 'Help Desk'
                Alias = 'helpdesk'
                PrimarySmtpAddress = 'helpdesk@example.com'
                RecipientTypeDetails = 'SharedMailbox'
                EmailAddresses = @('SMTP:helpdesk@example.com', 'smtp:support@example.com', 'X500:legacy')
            }
        }
    }

    It 'checks both fields by default and returns SMTP addresses' {
        $result = & $script:lookupPath -Identity 'helpdesk'
        $result.Alias | Should Be 'helpdesk'
        $result.EmailAddresses.Count | Should Be 2
        Assert-MockCalled Get-EXOMailbox -Times 1 -Exactly -Scope It -ParameterFilter {
            $Filter -eq "(Alias -eq 'helpdesk') -or (EmailAddresses -eq 'smtp:helpdesk')" -and
            $ResultSize -eq 'Unlimited' -and $Properties -contains 'EmailAddresses'
        }
        Assert-MockCalled Connect-ExchangeOnline -Times 0 -Exactly -Scope It
    }

    It 'looks up a full secondary email address' {
        & $script:lookupPath -Identity 'support@example.com' -Lookup Email
        Assert-MockCalled Get-EXOMailbox -Times 1 -Exactly -Scope It -ParameterFilter {
            $Filter -eq "EmailAddresses -eq 'smtp:support@example.com'"
        }
    }

    It 'trims input and escapes apostrophes in exact alias filters' {
        & $script:lookupPath -Identity " o'brien " -Lookup Alias
        Assert-MockCalled Get-EXOMailbox -Times 1 -Exactly -Scope It -ParameterFilter {
            $Filter -eq "Alias -eq 'o''brien'"
        }
    }

    It 'connects when there is no active connection' {
        Mock Get-ConnectionInformation { @() }
        & $script:lookupPath -Identity 'helpdesk'
        Assert-MockCalled Connect-ExchangeOnline -Times 1 -Exactly -Scope It
    }

    It 'reports no matches without returning a mailbox' {
        Mock Get-EXOMailbox { @() }
        $result = @(& $script:lookupPath -Identity 'missing')
        $result.Count | Should Be 0
        Assert-MockCalled Write-Host -Times 1 -Exactly -Scope It -ParameterFilter {
            $Object -like 'No mailbox found*'
        }
    }

    It 'does not report a query failure as a missing mailbox' {
        Mock Get-EXOMailbox { throw 'Access denied' }
        { & $script:lookupPath -Identity 'helpdesk' } | Should Throw 'Mailbox lookup failed: Access denied'
        Assert-MockCalled Write-Host -Times 0 -Exactly -Scope It
    }

    It 'does not query after a connection failure' {
        Mock Get-ConnectionInformation { @() }
        Mock Connect-ExchangeOnline { throw 'Sign-in failed' }
        { & $script:lookupPath -Identity 'helpdesk' } | Should Throw 'Mailbox lookup failed: Sign-in failed'
        Assert-MockCalled Get-EXOMailbox -Times 0 -Exactly -Scope It
    }

    It 'rejects whitespace-only input' {
        { & $script:lookupPath -Identity '   ' } | Should Throw 'Enter a mailbox alias or a full email address.'
        Assert-MockCalled Get-EXOMailbox -Times 0 -Exactly -Scope It
    }

    It 'uses browser sign-in in the GUI worker and returns the shared lookup results' {
        $assignment = $script:guiAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$worker'
        }, $true)
        $worker = $assignment.Right.Expression.ScriptBlock.GetScriptBlock()
        $workerState = @{ Connected = $false }
        Mock Get-ConnectionInformation ({
            if ($workerState.Connected) { [pscustomobject]@{ State = 'Connected' } }
        }.GetNewClosure())
        Mock Connect-ExchangeOnline ({ $workerState.Connected = $true }.GetNewClosure())
        $result = & $worker $script:lookupPath 'support@example.com' 'Email'
        $result.PrimarySmtpAddress | Should Be 'helpdesk@example.com'
        Assert-MockCalled Connect-ExchangeOnline -Times 1 -Exactly -Scope It -ParameterFilter { $DisableWAM }
        Assert-MockCalled Get-EXOMailbox -Times 1 -Exactly -Scope It -ParameterFilter {
            $Filter -eq "EmailAddresses -eq 'smtp:support@example.com'"
        }
    }
}

Describe 'Mailbox lookup GUI' {
    It 'loads the WPF window with all lookup modes and result columns' {
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        $assignment = $script:guiAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$xaml'
        }, $true)
        $reader = [System.Xml.XmlNodeReader]::new([xml]$assignment.Right.Expression.Value)
        try { $window = [System.Windows.Markup.XamlReader]::Load($reader) }
        finally { $reader.Close() }
        try {
            $window.FindName('LookupBox').Items.Count | Should Be 3
            $window.FindName('LookupBox').SelectedItem.Tag | Should Be 'Both'
            $window.FindName('ResultsGrid').Columns.Count | Should Be 5
            $window.FindName('ResultsGrid').IsReadOnly | Should Be $true
            $window.FindName('SearchButton').IsEnabled | Should Be $false
            $window.FindName('IdentityBox') | Should Not BeNullOrEmpty
            $window.FindName('StatusText').Text | Should Be 'Not connected'
        }
        finally { $window.Close() }
    }

    It 'opens the real window without arguments and enables search after typing' {
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        $guiSmoke = @{ Seen = $false; Enabled = $false; InitialLookup = ''; Error = $null }
        $closeTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $closeTimer.Interval = [TimeSpan]::FromMilliseconds(100)
        $closeTimer.Add_Tick({
            $lookupWindow = @([System.Windows.PresentationSource]::CurrentSources |
                ForEach-Object { $_.RootVisual } |
                Where-Object { $_ -is [System.Windows.Window] -and $_.Title -eq 'Mailbox Lookup' })
            if ($lookupWindow.Count -eq 0) { return }
            try {
                $guiSmoke.Seen = $true
                $guiSmoke.InitialLookup = $lookupWindow[0].FindName('LookupBox').SelectedItem.Tag
                $lookupWindow[0].FindName('IdentityBox').Text = 'support@example.com'
                $guiSmoke.Enabled = $lookupWindow[0].FindName('SearchButton').IsEnabled
            }
            catch { $guiSmoke.Error = $_.Exception.Message }
            finally { $lookupWindow[0].Close() }
        }.GetNewClosure())
        $closeTimer.Start()
        try { & $script:lookupPath -Lookup Email }
        finally { $closeTimer.Stop() }
        $guiSmoke.Error | Should BeNullOrEmpty
        $guiSmoke.Seen | Should Be $true
        $guiSmoke.Enabled | Should Be $true
        $guiSmoke.InitialLookup | Should Be 'Email'
    }
}