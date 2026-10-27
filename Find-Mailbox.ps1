#requires -Version 7.6

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Identity,

    [ValidateSet('Both', 'Alias', 'Email')]
    [string]$Lookup = 'Both'
)

Set-StrictMode -Version 2.0

$searchIdentity = $Identity.Trim()
if ([string]::IsNullOrWhiteSpace($searchIdentity)) {
    throw 'Enter a mailbox alias or a full email address.'
}

try {
    Import-Module ExchangeOnlineManagement -MinimumVersion 3.7.0 -ErrorAction Stop
    $connections = @(Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' })
    if ($connections.Count -eq 0) {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    }

    $escapedIdentity = $searchIdentity.Replace("'", "''")
    $aliasFilter = "Alias -eq '$escapedIdentity'"
    $emailFilter = "EmailAddresses -eq 'smtp:$escapedIdentity'"
    $filter = switch ($Lookup) {
        'Alias' { $aliasFilter }
        'Email' { $emailFilter }
        'Both' { "($aliasFilter) -or ($emailFilter)" }
    }

    $mailboxes = @(Get-EXOMailbox -Filter $filter -ResultSize Unlimited -Properties EmailAddresses -ErrorAction Stop)
    if ($mailboxes.Count -eq 0) {
        Write-Host "No mailbox found for '$searchIdentity' ($Lookup lookup)."
        return
    }

    $mailboxes | Select-Object DisplayName, Alias, PrimarySmtpAddress, RecipientTypeDetails,
        @{ Name = 'EmailAddresses'; Expression = { @($_.EmailAddresses | Where-Object { $_ -like 'smtp:*' }) } }
}
catch {
    throw "Mailbox lookup failed: $($_.Exception.Message)"
}