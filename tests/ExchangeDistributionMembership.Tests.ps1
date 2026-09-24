$modulePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'ExchangeDistributionMembership.psm1'
Import-Module $modulePath -Force

Describe 'Test-RecipientIdentifier' {
    It 'accepts an email address or UPN shape' {
        Test-RecipientIdentifier -Identifier 'alex.wilber@contoso.com' | Should Be $true
    }

    It 'trims surrounding whitespace' {
        Test-RecipientIdentifier -Identifier '  alex.wilber@contoso.com  ' | Should Be $true
    }

    It 'rejects empty, wildcard, and malformed values' {
        Test-RecipientIdentifier -Identifier '' | Should Be $false
        Test-RecipientIdentifier -Identifier '*@contoso.com' | Should Be $false
        Test-RecipientIdentifier -Identifier 'not-an-address' | Should Be $false
    }
}

Describe 'Resolve-ExactExchangeRecipient lookup modes' {
    InModuleScope ExchangeDistributionMembership {
        if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
            function Get-Mailbox { }
        }
    }

    It 'includes soft-deleted recipients when requested' {
        Mock Get-EXORecipient -ModuleName ExchangeDistributionMembership {
            param($Identity, $IncludeSoftDeletedRecipients)
            [pscustomobject]@{
                DisplayName              = 'Former User'
                PrimarySmtpAddress       = 'former.user@contoso.com'
                UserPrincipalName        = 'former.user@contoso.com'
                EmailAddresses           = @('smtp:former.user@contoso.com')
                ExternalDirectoryObjectId = 'soft-deleted-1'
                Guid                      = '11111111-1111-1111-1111-111111111111'
            }
        }

        $recipient = Resolve-ExactExchangeRecipient `
            -Identifier 'former.user@contoso.com' `
            -LookupMode SoftDeleted

        $recipient.DisplayName | Should Be 'Former User'
        Assert-MockCalled Get-EXORecipient -ModuleName ExchangeDistributionMembership -ParameterFilter {
            $IncludeSoftDeletedRecipients -eq $true
        }
    }

    It 'resolves inactive mailboxes with the inactive-only switch' {
        Mock Get-Mailbox -ModuleName ExchangeDistributionMembership {
            param($Identity, $InactiveMailboxOnly)
            [pscustomobject]@{
                DisplayName              = 'Retained User'
                PrimarySmtpAddress       = 'retained.user@contoso.com'
                UserPrincipalName        = 'retained.user@contoso.com'
                EmailAddresses           = @('smtp:retained.user@contoso.com')
                ExternalDirectoryObjectId = 'inactive-1'
                Guid                      = '22222222-2222-2222-2222-222222222222'
            }
        }

        $recipient = Resolve-ExactExchangeRecipient `
            -Identifier 'retained.user@contoso.com' `
            -LookupMode InactiveMailbox

        $recipient.DisplayName | Should Be 'Retained User'
    }
}

Describe 'Test-RecipientIdentityMatch' {
    It 'uses an exact case-insensitive immutable identity match' {
        $target = [pscustomobject]@{
            ExternalDirectoryObjectId = 'ABC-123'
            Guid                      = '11111111-1111-1111-1111-111111111111'
            PrimarySmtpAddress        = 'alex.wilber@contoso.com'
        }
        $candidate = [pscustomobject]@{
            ExternalDirectoryObjectId = 'abc-123'
            Guid                      = '22222222-2222-2222-2222-222222222222'
            PrimarySmtpAddress        = 'other@contoso.com'
        }

        Test-RecipientIdentityMatch -Target $target -Candidate $candidate | Should Be $true
    }

    It 'does not fall back after a mutually populated stronger identity differs' {
        $target = [pscustomobject]@{
            ExternalDirectoryObjectId = 'ABC-123'
            Guid                      = '11111111-1111-1111-1111-111111111111'
            PrimarySmtpAddress        = 'alex.wilber@contoso.com'
        }
        $candidate = [pscustomobject]@{
            ExternalDirectoryObjectId = 'DIFFERENT'
            Guid                      = '11111111-1111-1111-1111-111111111111'
            PrimarySmtpAddress        = 'alex.wilber@contoso.com'
        }

        Test-RecipientIdentityMatch -Target $target -Candidate $candidate | Should Be $false
    }
}

Describe 'New-DynamicRecipientPreviewFilter' {
    It 'combines the original filter with an exact ExternalDirectoryObjectId predicate' {
        $recipient = [pscustomobject]@{
            ExternalDirectoryObjectId = 'ABC-123'
            Guid                      = '11111111-1111-1111-1111-111111111111'
        }

        $actual = New-DynamicRecipientPreviewFilter `
            -RecipientFilter "(RecipientTypeDetails -eq 'UserMailbox')" `
            -Recipient $recipient

        $actual | Should Be "((RecipientTypeDetails -eq 'UserMailbox')) -and (ExternalDirectoryObjectId -eq 'ABC-123')"
    }

    It 'escapes single quotes in the exact identifier value' {
        $recipient = [pscustomobject]@{
            ExternalDirectoryObjectId = "ABC'123"
            Guid                      = $null
        }

        $actual = New-DynamicRecipientPreviewFilter -RecipientFilter '(Alias -like ''A*'')' -Recipient $recipient
        $actual | Should Match "ABC''123"
    }

    It 'falls back to Guid and rejects a recipient without an immutable identity' {
        $guidRecipient = [pscustomobject]@{
            ExternalDirectoryObjectId = $null
            Guid                      = '11111111-1111-1111-1111-111111111111'
        }
        New-DynamicRecipientPreviewFilter -RecipientFilter '(Alias -like ''A*'')' -Recipient $guidRecipient |
            Should Match '\(Guid -eq'

        $invalidRecipient = [pscustomobject]@{
            ExternalDirectoryObjectId = $null
            Guid                      = $null
        }
        $didThrow = $false
        try {
            $null = New-DynamicRecipientPreviewFilter -RecipientFilter '(Alias -like ''A*'')' -Recipient $invalidRecipient
        }
        catch {
            $didThrow = $true
        }
        $didThrow | Should Be $true
    }
}

Describe 'Export-UserDistributionGroupMembership' {
    It 'writes sorted membership rows and returns an accurate summary' {
        Mock Resolve-ExactExchangeRecipient -ModuleName ExchangeDistributionMembership {
            [pscustomobject]@{
                DisplayName        = 'Alex Wilber'
                PrimarySmtpAddress = 'alex.wilber@contoso.com'
            }
        }
        Mock Get-StaticDistributionMembershipRows -ModuleName ExchangeDistributionMembership {
            [pscustomobject]@{
                Rows = @(
                    [pscustomobject][ordered]@{
                        DisplayName        = 'Zeta Group'
                        PrimarySmtpAddress = 'zeta@contoso.com'
                        Identity           = 'Zeta Group'
                        GroupType          = 'Distribution group'
                        MembershipType     = 'Direct'
                        RecipientFilter    = ''
                        RecipientContainer = ''
                    }
                )
                Errors = @()
                Count  = 5
            }
        }
        Mock Get-DynamicDistributionMembershipRows -ModuleName ExchangeDistributionMembership {
            [pscustomobject]@{
                Rows = @(
                    [pscustomobject][ordered]@{
                        DisplayName        = 'All Sales'
                        PrimarySmtpAddress = 'sales@contoso.com'
                        Identity           = 'All Sales'
                        GroupType          = 'Dynamic distribution group'
                        MembershipType     = 'Current recipient filter match'
                        RecipientFilter    = "(Department -eq 'Sales')"
                        RecipientContainer = 'contoso.com/Users'
                    }
                )
                Errors = @()
                Count  = 3
            }
        }

        $path = Join-Path $TestDrive 'membership.csv'
        $summary = Export-UserDistributionGroupMembership `
            -Identifier 'alex.wilber@contoso.com' `
            -Path $path
        $rows = @(Import-Csv -LiteralPath $path)

        $summary.TotalGroupCount | Should Be 2
        $summary.StaticGroupsChecked | Should Be 5
        $summary.DynamicGroupsChecked | Should Be 3
        $rows.Count | Should Be 2
        $rows[0].DisplayName | Should Be 'Zeta Group'
        $rows[1].DisplayName | Should Be 'All Sales'
    }

    It 'writes only column headers when there are no memberships' {
        Mock Resolve-ExactExchangeRecipient -ModuleName ExchangeDistributionMembership {
            [pscustomobject]@{
                DisplayName        = 'Alex Wilber'
                PrimarySmtpAddress = 'alex.wilber@contoso.com'
            }
        }
        Mock Get-StaticDistributionMembershipRows -ModuleName ExchangeDistributionMembership {
            [pscustomobject]@{ Rows = @(); Errors = @(); Count = 1 }
        }
        Mock Get-DynamicDistributionMembershipRows -ModuleName ExchangeDistributionMembership {
            [pscustomobject]@{ Rows = @(); Errors = @(); Count = 1 }
        }

        $path = Join-Path $TestDrive 'empty.csv'
        $summary = Export-UserDistributionGroupMembership `
            -Identifier 'alex.wilber@contoso.com' `
            -Path $path

        $summary.TotalGroupCount | Should Be 0
        @(Get-Content -LiteralPath $path).Count | Should Be 1
        (Get-Content -LiteralPath $path -First 1) | Should Match '"DisplayName"'
    }

    It 'does not create a CSV when any group evaluation fails' {
        Mock Resolve-ExactExchangeRecipient -ModuleName ExchangeDistributionMembership {
            [pscustomobject]@{
                DisplayName        = 'Alex Wilber'
                PrimarySmtpAddress = 'alex.wilber@contoso.com'
            }
        }
        Mock Get-StaticDistributionMembershipRows -ModuleName ExchangeDistributionMembership {
            [pscustomobject]@{ Rows = @(); Errors = @('Denied'); Count = 1 }
        }
        Mock Get-DynamicDistributionMembershipRows -ModuleName ExchangeDistributionMembership {
            [pscustomobject]@{ Rows = @(); Errors = @(); Count = 1 }
        }

        $path = Join-Path $TestDrive 'incomplete.csv'
        $didThrow = $false
        try {
            Export-UserDistributionGroupMembership `
                -Identifier 'alex.wilber@contoso.com' `
                -Path $path
        }
        catch {
            $didThrow = $true
        }
        $didThrow | Should Be $true
        Test-Path -LiteralPath $path | Should Be $false
    }
}
