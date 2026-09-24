Set-StrictMode -Version 2.0

function Test-RecipientIdentifier {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Identifier
    )

    $value = $Identifier.Trim()
    if ([string]::IsNullOrWhiteSpace($value) -or $value.IndexOfAny([char[]]'*?') -ge 0) {
        return $false
    }

    try {
        $address = New-Object System.Net.Mail.MailAddress($value)
        return $address.Address.Equals($value, [System.StringComparison]::OrdinalIgnoreCase)
    }
    catch {
        return $false
    }
}

function Test-RecipientIdentityMatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Target,

        [Parameter(Mandatory = $true)]
        [object]$Candidate
    )

    foreach ($propertyName in @('ExternalDirectoryObjectId', 'Guid', 'PrimarySmtpAddress')) {
        $targetValue = $Target.$propertyName
        $candidateValue = $Candidate.$propertyName
        if (-not [string]::IsNullOrWhiteSpace([string]$targetValue) -and
            -not [string]::IsNullOrWhiteSpace([string]$candidateValue)) {
            return ([string]$targetValue).Equals(
                [string]$candidateValue,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        }
    }

    return $false
}

function New-DynamicRecipientPreviewFilter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RecipientFilter,

        [Parameter(Mandatory = $true)]
        [object]$Recipient
    )

    if (-not [string]::IsNullOrWhiteSpace([string]$Recipient.ExternalDirectoryObjectId)) {
        $propertyName = 'ExternalDirectoryObjectId'
        $identifier = [string]$Recipient.ExternalDirectoryObjectId
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$Recipient.Guid)) {
        $propertyName = 'Guid'
        $identifier = [string]$Recipient.Guid
    }
    else {
        throw 'The recipient has no immutable identifier that can be used to evaluate a dynamic distribution group filter.'
    }

    $escapedIdentifier = $identifier.Replace("'", "''")
    return "($RecipientFilter) -and ($propertyName -eq '$escapedIdentifier')"
}

function Resolve-ExactExchangeRecipient {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identifier
    )

    $value = $Identifier.Trim()
    if (-not (Test-RecipientIdentifier -Identifier $value)) {
        throw 'Enter a valid email address or user principal name. Wildcards are not allowed.'
    }

    try {
        $recipients = @(
            Get-EXORecipient -Identity $value `
                -Properties EmailAddresses, UserPrincipalName, PrimarySmtpAddress, Guid, ExternalDirectoryObjectId, RecipientTypeDetails, DisplayName `
                -ResultSize 2 `
                -ErrorAction Stop
        )
    }
    catch {
        throw "Exchange Online could not resolve '$value'. Verify the address and your recipient-read permissions. $($_.Exception.Message)"
    }

    if ($recipients.Count -ne 1) {
        throw "Exchange Online returned $($recipients.Count) recipients for '$value'; exactly one is required."
    }

    $recipient = $recipients[0]
    $matchesUpn = -not [string]::IsNullOrWhiteSpace([string]$recipient.UserPrincipalName) -and
        ([string]$recipient.UserPrincipalName).Equals($value, [System.StringComparison]::OrdinalIgnoreCase)
    $matchesProxyAddress = $false

    foreach ($proxyAddress in @($recipient.EmailAddresses)) {
        $proxyValue = [string]$proxyAddress
        if ($proxyValue.StartsWith('smtp:', [System.StringComparison]::OrdinalIgnoreCase) -and
            $proxyValue.Substring(5).Equals($value, [System.StringComparison]::OrdinalIgnoreCase)) {
            $matchesProxyAddress = $true
            break
        }
    }

    if (-not $matchesUpn -and -not $matchesProxyAddress) {
        throw "Exchange Online resolved '$value' to a different recipient. Use an exact UPN or SMTP proxy address."
    }

    return $recipient
}

function Get-StaticDistributionMembershipRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Recipient,

        [Parameter(Mandatory = $false)]
        [scriptblock]$StatusAction
    )

    $groups = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop)
    $rows = New-Object System.Collections.Generic.List[object]
    $errors = New-Object System.Collections.Generic.List[string]

    for ($index = 0; $index -lt $groups.Count; $index++) {
        $group = $groups[$index]
        if ($StatusAction) {
            $null = & $StatusAction ($index + 1) $groups.Count "Checking static group: $($group.DisplayName)"
        }

        $groupIdentity = if (-not [string]::IsNullOrWhiteSpace([string]$group.Guid)) {
            [string]$group.Guid
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$group.PrimarySmtpAddress)) {
            [string]$group.PrimarySmtpAddress
        }
        else {
            $null
        }

        if ([string]::IsNullOrWhiteSpace($groupIdentity)) {
            $errors.Add("Static group '$($group.DisplayName)' has no usable identity.")
            continue
        }

        try {
            $isMember = $false
            foreach ($member in @(Get-DistributionGroupMember -Identity $groupIdentity -ResultSize Unlimited -ErrorAction Stop)) {
                if (Test-RecipientIdentityMatch -Target $Recipient -Candidate $member) {
                    $isMember = $true
                    break
                }
            }

            if ($isMember) {
                $groupType = if ([string]$group.RecipientTypeDetails -like '*SecurityGroup*' -or
                    [string]$group.GroupType -like '*SecurityEnabled*') {
                    'Mail-enabled security group'
                }
                else {
                    'Distribution group'
                }

                $rows.Add([pscustomobject][ordered]@{
                    DisplayName        = [string]$group.DisplayName
                    PrimarySmtpAddress = [string]$group.PrimarySmtpAddress
                    Identity           = [string]$group.Identity
                    GroupType          = $groupType
                    MembershipType     = 'Direct'
                    RecipientFilter    = ''
                    RecipientContainer = ''
                })
            }
        }
        catch {
            $errors.Add("Static group '$($group.DisplayName)': $($_.Exception.Message)")
        }
    }

    return [pscustomobject]@{
        Rows   = $rows.ToArray()
        Errors = $errors.ToArray()
        Count  = $groups.Count
    }
}

function Get-DynamicDistributionMembershipRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Recipient,

        [Parameter(Mandatory = $false)]
        [scriptblock]$StatusAction
    )

    $groups = @(Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction Stop)
    $rows = New-Object System.Collections.Generic.List[object]
    $errors = New-Object System.Collections.Generic.List[string]

    for ($index = 0; $index -lt $groups.Count; $index++) {
        $group = $groups[$index]
        if ($StatusAction) {
            $null = & $StatusAction ($index + 1) $groups.Count "Evaluating dynamic group: $($group.DisplayName)"
        }

        try {
            if ([string]::IsNullOrWhiteSpace([string]$group.RecipientFilter)) {
                throw 'The group has no recipient filter.'
            }

            $parameters = @{
                RecipientPreviewFilter = New-DynamicRecipientPreviewFilter -RecipientFilter ([string]$group.RecipientFilter) -Recipient $Recipient
                ResultSize             = 2
                ErrorAction            = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$group.RecipientContainer)) {
                $parameters.OrganizationalUnit = [string]$group.RecipientContainer
            }

            $matches = @(Get-Recipient @parameters)
            if ($matches.Count -gt 1) {
                throw 'The exact recipient predicate unexpectedly returned more than one object.'
            }

            if ($matches.Count -eq 1 -and (Test-RecipientIdentityMatch -Target $Recipient -Candidate $matches[0])) {
                $rows.Add([pscustomobject][ordered]@{
                    DisplayName        = [string]$group.DisplayName
                    PrimarySmtpAddress = [string]$group.PrimarySmtpAddress
                    Identity           = [string]$group.Identity
                    GroupType          = 'Dynamic distribution group'
                    MembershipType     = 'Current recipient filter match'
                    RecipientFilter    = [string]$group.RecipientFilter
                    RecipientContainer = [string]$group.RecipientContainer
                })
            }
        }
        catch {
            $errors.Add("Dynamic group '$($group.DisplayName)': $($_.Exception.Message)")
        }
    }

    return [pscustomobject]@{
        Rows   = $rows.ToArray()
        Errors = $errors.ToArray()
        Count  = $groups.Count
    }
}

function Export-UserDistributionGroupMembership {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identifier,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [scriptblock]$StatusAction
    )

    $destination = [System.IO.Path]::GetFullPath($Path)
    $destinationDirectory = [System.IO.Path]::GetDirectoryName($destination)
    if (-not [System.IO.Directory]::Exists($destinationDirectory)) {
        throw "The destination folder does not exist: $destinationDirectory"
    }

    if ($StatusAction) {
        $null = & $StatusAction 0 1 'Resolving recipient...'
    }
    $recipient = Resolve-ExactExchangeRecipient -Identifier $Identifier

    $staticProgress = {
        param($Current, $Total, $Message)
        if ($StatusAction) {
            $percent = if ($Total -gt 0) { [int](5 + (($Current / $Total) * 45)) } else { 50 }
            $null = & $StatusAction $percent 100 $Message
        }
    }
    $staticResult = Get-StaticDistributionMembershipRows -Recipient $recipient -StatusAction $staticProgress

    $dynamicProgress = {
        param($Current, $Total, $Message)
        if ($StatusAction) {
            $percent = if ($Total -gt 0) { [int](50 + (($Current / $Total) * 45)) } else { 95 }
            $null = & $StatusAction $percent 100 $Message
        }
    }
    $dynamicResult = Get-DynamicDistributionMembershipRows -Recipient $recipient -StatusAction $dynamicProgress

    $evaluationErrors = @($staticResult.Errors) + @($dynamicResult.Errors)
    if ($evaluationErrors.Count -gt 0) {
        $shownErrors = @($evaluationErrors | Select-Object -First 10)
        $moreText = if ($evaluationErrors.Count -gt 10) {
            "`n...and $($evaluationErrors.Count - 10) more error(s)."
        }
        else {
            ''
        }
        throw "The export was not created because membership could not be evaluated completely. Fix the permissions, connectivity, or group filter errors below and try again.`n`n$($shownErrors -join "`n")$moreText"
    }

    $rows = @($staticResult.Rows) + @($dynamicResult.Rows)
    if ($rows.Count -eq 0) {
        [pscustomobject][ordered]@{
            DisplayName        = ''
            PrimarySmtpAddress = ''
            Identity           = ''
            GroupType          = ''
            MembershipType     = ''
            RecipientFilter    = ''
            RecipientContainer = ''
        } |
            ConvertTo-Csv -NoTypeInformation |
            Select-Object -First 1 |
            Set-Content -LiteralPath $destination -Encoding UTF8
    }
    else {
        $rows |
            Sort-Object GroupType, DisplayName, PrimarySmtpAddress |
            Export-Csv -LiteralPath $destination -NoTypeInformation -Encoding UTF8
    }

    if ($StatusAction) {
        $null = & $StatusAction 100 100 'Export complete.'
    }

    return [pscustomobject]@{
        RecipientDisplayName = [string]$recipient.DisplayName
        RecipientAddress     = [string]$recipient.PrimarySmtpAddress
        StaticGroupCount     = @($staticResult.Rows).Count
        DynamicGroupCount    = @($dynamicResult.Rows).Count
        TotalGroupCount      = $rows.Count
        StaticGroupsChecked  = $staticResult.Count
        DynamicGroupsChecked = $dynamicResult.Count
        Path                 = $destination
    }
}

Export-ModuleMember -Function @(
    'Test-RecipientIdentifier',
    'Test-RecipientIdentityMatch',
    'New-DynamicRecipientPreviewFilter',
    'Resolve-ExactExchangeRecipient',
    'Get-StaticDistributionMembershipRows',
    'Get-DynamicDistributionMembershipRows',
    'Export-UserDistributionGroupMembership'
)
