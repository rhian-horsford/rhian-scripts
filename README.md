# Exchange Distribution Membership Exporter

A native Windows PowerShell/WPF tool that exports every Exchange Online static
distribution group and dynamic distribution group that currently matches one
recipient.

The export treats the two group models differently:

- **Static distribution groups and mail-enabled security groups:** checks the
  group's direct members with `Get-DistributionGroupMember`.
- **Dynamic distribution groups:** asks Exchange Online to evaluate each
  group's `RecipientFilter` for the resolved recipient, including the group's
  `RecipientContainer` scope when one is configured.

## Prerequisites

- Windows 10 or later with Windows PowerShell 5.1
- PowerShell script execution permitted by your organization's policy
- The free
  [ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement)
  module
- An Exchange Online account allowed to read recipients, distribution groups,
  group members, and dynamic distribution group configuration

Install the module for the current user from a trusted PowerShell Gallery
session:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

Use least privilege. The signed-in account must be able to run
`Get-EXORecipient`, `Get-DistributionGroup`,
`Get-DistributionGroupMember`, `Get-DynamicDistributionGroup`, and
`Get-Recipient`. The built-in **View-Only Recipients** role typically supplies
the required recipient read access; organizational RBAC customizations can
change this. The tool stops rather than producing a success-shaped, incomplete
export when a group cannot be evaluated.

## Run the tool

1. Download or clone this repository.
2. Review the scripts before running them. Do not weaken the machine execution
   policy or use an execution-policy bypass. If permitted by organizational
   policy, a typical current-user setting is `RemoteSigned`.
3. From Windows PowerShell 5.1, run:

   ```powershell
   .\Export-ExchangeDistributionMembership.ps1
   ```

4. Select **Connect** and complete the Microsoft Exchange Online sign-in flow.
5. Enter the user's exact email address or user principal name. Wildcards,
   aliases, and display names are intentionally rejected.
6. Select **Browse**, choose a CSV destination, and select **Export CSV**.

Authentication and membership evaluation run on a reusable background runspace,
so the WPF window remains responsive. The status area shows the current group
and progress. Permission, module, connection, recipient-resolution, and
per-group filter failures appear as actionable error dialogs.

The tool only reads Exchange Online and writes the selected local CSV. It does
not modify recipients or groups. The CSV is UTF-8 encoded and opens directly in
Excel.

## Output columns

| Column | Meaning |
| --- | --- |
| `DisplayName` | Group display name |
| `PrimarySmtpAddress` | Group primary SMTP address |
| `Identity` | Exchange group identity |
| `GroupType` | Distribution group, mail-enabled security group, or dynamic distribution group |
| `MembershipType` | `Direct` for static groups or `Current recipient filter match` for dynamic groups |
| `RecipientFilter` | OPATH recipient filter for a dynamic group; blank for a static group |
| `RecipientContainer` | Dynamic group recipient container/scope when configured |

An empty result still creates a CSV containing the column headers.

## Membership semantics and limitations

- **Static membership is direct only.** If a user belongs to a group only
  through a nested group, that parent group is not included. Exchange Online
  group member enumeration is deliberately not recursively expanded.
- **Dynamic results are current filter evaluations.** The tool combines the
  group's existing OPATH filter with the recipient's immutable Exchange
  identity and evaluates it server-side using `Get-Recipient
  -RecipientPreviewFilter`. When present, `RecipientContainer` is supplied as
  `-OrganizationalUnit`; it is not approximated with string matching.
- Exchange Online stores a calculated dynamic distribution group membership
  list for message delivery and refreshes it periodically (normally every 24
  hours). A current filter match can therefore temporarily differ from the
  stored delivery list.
- Enumeration time grows with the number and size of groups. Exchange Online
  throttling and transient service issues can lengthen or stop an export.
- Hidden membership or restricted recipient scopes can require additional
  RBAC permissions. Any group that cannot be checked makes the export fail so
  a missing permission cannot be mistaken for non-membership.
- The tool requires network access to Exchange Online and supports the
  authentication methods provided by the installed
  `ExchangeOnlineManagement` module.

## Tests

Focused Pester tests cover identifier validation, exact identity matching, and
safe dynamic preview-filter construction:

```powershell
Invoke-Pester .\tests\ExchangeDistributionMembership.Tests.ps1
```

Pester 5 is recommended, but the tests use syntax compatible with commonly
deployed Pester versions.
