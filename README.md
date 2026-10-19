# Exchange Administration Scripts

A collection of Windows PowerShell/WPF utilities for Exchange Online recipient
and group administration.

## Tools

| Script | Purpose | Exchange changes |
| --- | --- | --- |
| [`Export-ExpandedDistributionGroupMembers.ps1`](Export-ExpandedDistributionGroupMembers.ps1) | Expand nested groups, evaluate dynamic groups, and export Microsoft 365 group members and owners. | Read-only |
| [`Add-SharedMailboxPermission.ps1`](Add-SharedMailboxPermission.ps1) | Resolve users by name or email, then grant Full Access, Send As, or both on a shared mailbox. | Adds permissions |
| [`Add-GroupOwnerOrMember.ps1`](Add-GroupOwnerOrMember.ps1) | Resolve users by name or email, then add members, owners, or both to a distribution list or Microsoft 365 group. | Adds group roles |

## Requirements

All tools use a GUI and require PowerShell 7.6 or later with
ExchangeOnlineManagement 3.7 or later. ExchangeOnlineManagement 3.10.x also
requires PowerShell 7.6 or later. The GUIs follow Windows light/dark appearance
and update their colors while open.

Install the free
[ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement)
module for the current user from a trusted PowerShell Gallery session:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

From the VS Code PowerShell terminal, launch a tool with:

```powershell
& .\Export-ExpandedDistributionGroupMembers.ps1
& .\Add-SharedMailboxPermission.ps1
& .\Add-GroupOwnerOrMember.ps1
```

Review scripts before running them. Do not weaken the machine execution policy
or use an execution-policy bypass.

## Export-ExpandedDistributionGroupMembers

Enter a parent group identity and choose an export folder. The tool expands
nested distribution and mail-enabled security groups, evaluates dynamic group
filters, and retrieves direct Microsoft 365 group members and owners. Duplicate
recipients are combined by primary SMTP address, with source groups, group
types, and membership roles preserved in the CSV. Microsoft 365 owners who are
also members appear once with both roles.

The default export folder is `OneDrive - Slalom\Desktop`. The CSV filename
ends in `-Expanded-Members-and-Owners.csv`. If any group cannot be evaluated,
the Warnings tab reports the issue and the tool does not write a partial CSV.

The CSV contains one row per unique primary SMTP address:

| Column | Meaning |
| --- | --- |
| `DisplayName` | Recipient display name |
| `Alias` | Recipient alias |
| `PrimarySmtpAddress` | Recipient's primary SMTP address |
| `RecipientTypeDetails` | Exchange recipient type |
| `SourceGroupTypes` | Group types that contributed the recipient |
| `SourceGroups` | Group display names that contributed the recipient |
| `MembershipTypes` | `Member`, `Owner`, `Dynamic`, or `Direct / Exception`; combined when applicable |

The signed-in account needs permission to run `Get-DistributionGroup`,
`Get-DistributionGroupMember`, `Get-DynamicDistributionGroup`, `Get-Recipient`,
`Get-UnifiedGroup`, and `Get-UnifiedGroupLinks`. M365 owners and members are
exported; subscribers are not. Microsoft 365 group members are direct; nested
distribution groups are recursively expanded.

## Add-SharedMailboxPermission

Enter a shared mailbox email address and one or more user names or email
addresses. Select **Check users** to resolve entries; choose a specific match
when a name is ambiguous. Then select **Full Access**, **Send As**, or both.
Auto-mapping applies to Full Access. The tool asks for confirmation before
granting access and shows the result for each user and permission.

The signed-in account needs Exchange RBAC permission to run
`Add-MailboxPermission` and `Add-RecipientPermission`, commonly provided by
**Recipient Management** or an equivalent custom role. Permission changes are
applied immediately; use a test mailbox when validating the workflow.

## Add-GroupOwnerOrMember

This is a standalone tool with the same interface and connection workflow as the
shared mailbox permission manager; it does not depend on either existing script.
Enter one group email address and one or more user names or email addresses.
Select **Check users**, choose a specific match for ambiguous names, then select
**Member**, **Owner**, or both and **Add roles**. Only user mailboxes are offered,
matching the mailbox permission manager's user lookup; guests, contacts, and
nested groups are not offered as users to add.

The tool detects static distribution lists, mail-enabled security groups, and
Microsoft 365 groups automatically. Dynamic groups and other recipient types
are rejected. Groups synchronized from on-premises must be changed in their
source directory.

For Microsoft 365 groups, adding an owner also adds missing membership first,
even if only **Owner** was selected. Both results appear in the table. If
membership fails, ownership is not attempted. For distribution lists, ownership
does not imply membership; select both roles if both are needed. Existing owners
and members are preserved, and existing assignments are reported as **Already
present** rather than added again.

Every write requires confirmation. Results and Exchange warnings are shown per
user and role. Successful changes are not rolled back if another change fails;
for example, membership remains if a subsequent ownership change fails.

The signed-in account needs permission to run `Get-EXORecipient`, `Get-Recipient`,
`Get-DistributionGroup`, `Get-DistributionGroupMember`, `Get-UnifiedGroup`,
`Get-UnifiedGroupLinks`, `Add-DistributionGroupMember`, `Set-DistributionGroup`,
and `Add-UnifiedGroupLinks`. For distribution lists you do not own, select
**Admin override (distribution lists only)** only when your account has
**Organization Management** or the **Security Group Creation and Membership**
role. This explicitly uses `BypassSecurityGroupManagerCheck`; it does not grant
additional permissions or affect Microsoft 365 groups.

Changes are applied immediately. Use a test group to validate the authenticated
workflow before making production changes.

### Offline validation

The group tool's tests use mocked Exchange commands and load the WPF markup
without signing in or making tenant changes. On Windows, run them in PowerShell
7.6 or later with the Windows-bundled Pester 3.4.0:

```powershell
Import-Module Pester -RequiredVersion 3.4.0
Invoke-Pester -Script .\Tests\Add-GroupOwnerOrMember.Tests.ps1
```

These tests do not replace validation of sign-in and Exchange RBAC with a test
group.
