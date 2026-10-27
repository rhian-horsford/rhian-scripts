#requires -Version 7.6

Set-StrictMode -Version 2.0

$script:lookupPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Find-Mailbox.ps1'

function Get-ConnectionInformation {
    [CmdletBinding()]
    param()
}
function Connect-ExchangeOnline {
    [CmdletBinding()]
    param([bool]$ShowBanner)
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
}