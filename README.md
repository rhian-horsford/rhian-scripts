# Exchange Administration Scripts

A collection of Windows PowerShell/WPF utilities for Exchange Online recipient
and group administration.

## Tools

| Script | Purpose | Exchange changes |
| --- | --- | --- |
| [`Export-ExpandedDistributionGroupMembers.ps1`](Export-ExpandedDistributionGroupMembers.ps1) | Expand nested groups, evaluate dynamic groups, and export Microsoft 365 group members and owners. | Read-only |
| [`Add-SharedMailboxPermission.ps1`](Add-SharedMailboxPermission.ps1) | Resolve users by name or email, then grant Full Access, Send As, or both on a shared mailbox. | Adds permissions |
| [`New-SharedMailbox.ps1`](New-SharedMailbox.ps1) | Create a shared mailbox and grant Full Access, Send As, or both to resolved users. | Creates mailbox and adds permissions |
| [`Add-GroupOwnerOrMember.ps1`](Add-GroupOwnerOrMember.ps1) | Resolve users by name or email, then add or remove members, owners, or both on a distribution list or Microsoft 365 group. | Adds or removes group roles |
| [`Find-Mailbox.ps1`](Find-Mailbox.ps1) | Find a mailbox by its alias or any assigned SMTP email address. | Read-only |

## Requirements

All tools use a GUI and require PowerShell 7.6 or later with
ExchangeOnlineManagement 3.7 or later. ExchangeOnlineManagement 3.10.x also
requires PowerShell 7.6 or later. The mailbox lookup GUI requires
ExchangeOnlineManagement 3.7.2 or later for browser sign-in.
The GUIs follow Windows light/dark appearance
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
& .\New-SharedMailbox.ps1
& .\Add-GroupOwnerOrMember.ps1
& .\Find-Mailbox.ps1
```

Review scripts before running them. Do not weaken the machine execution policy
or use an execution-policy bypass.

## Find-Mailbox

Run without arguments to open the mailbox lookup window. Enter an alias or full
email address, select **Alias and email**, **Alias only**, or **Email only**, and
select **Search**. Sign-in opens in the browser when needed. Searches run in a
background session so the window stays responsive. Results appear in a read-only
table; no match and lookup failures are reported separately below it.

Searches match the exact mailbox Alias field or a full primary or secondary SMTP
email address. By default, both fields are checked. It makes no Exchange changes.

Supplying `-Identity` retains terminal output instead of opening the GUI. Terminal
lookups reuse an active Exchange Online connection or prompt for sign-in:

```powershell
.\Find-Mailbox.ps1 -Identity 'helpdesk'
.\Find-Mailbox.ps1 -Identity 'support@example.com'
.\Find-Mailbox.ps1 -Identity 'helpdesk' -Lookup Alias
.\Find-Mailbox.ps1 -Identity 'support@example.com' -Lookup Email
```

Results include display name, alias, primary SMTP address, mailbox type, and
assigned SMTP addresses. A successful query with no results reports no mailbox
found; connection, permission, and query failures throw an error instead.
Only active mailboxes visible to the signed-in account are searched, including
user and shared mailboxes. This does not search contacts, groups, or deleted
mailboxes. The account needs Exchange RBAC permission to run `Get-EXOMailbox`.

Offline tests use mocked Exchange commands and briefly open and close the WPF
window without signing in:

```powershell
Import-Module Pester -RequiredVersion 3.4.0
Invoke-Pester -Script .\Tests\Find-Mailbox.Tests.ps1
```

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

## New-SharedMailbox

Run `New-SharedMailbox.ps1`, connect to Exchange Online, and enter a display
name, alias, and primary email address in an accepted domain for your tenant.
Names and aliases are limited to 64 characters. This tool accepts letters,
numbers, dots, underscores, and hyphens in aliases. Enter the users, select
**Check users**, and choose a match for any ambiguous names. Select **Full
Access**, **Send As**, or both; auto-mapping applies only to Full Access.
Select **Create mailbox** and review the confirmation before proceeding.

The launcher reuses `Add-SharedMailboxPermission.ps1`; keep both scripts in
the same folder. Launching the permissions script directly retains its normal
add/remove workflow. The new tool creates a shared mailbox using
`New-Mailbox -Shared`, then applies permissions using the same grant routine
as the existing tool. If creation fails, no permissions are granted and an
existing mailbox is not reused or modified.

Results show mailbox creation and each user/permission separately. Successful
changes are not rolled back. If provisioning delays or other errors prevent
permission grants, the mailbox remains created. Wait for provisioning to
complete, then use **Add-SharedMailboxPermission** to retry only failed grants.
Repeat creation is blocked in the current window after successful creation;
close and reopen the tool to create another mailbox. Exchange also rejects
conflicting names, aliases, or addresses.

The signed-in account needs Exchange RBAC permission to run `New-Mailbox`,
`Get-EXOMailbox`, `Get-EXORecipient`, `Get-Recipient`,
`Add-MailboxPermission`, and `Add-RecipientPermission`, commonly provided by
**Recipient Management** or an equivalent custom role.

Offline tests mock Exchange commands and check the WPF interface without
signing in or changing your tenant:

```powershell
Import-Module Pester -RequiredVersion 3.4.0
Invoke-Pester -Script .\Tests\New-SharedMailbox.Tests.ps1
```

Validate live sign-in, RBAC, and provisioning with a test mailbox before using
the tool for production changes.

## Add-GroupOwnerOrMember

This is a standalone tool with the same interface and connection workflow as the
shared mailbox permission manager; it does not depend on either existing script.
Enter one group email address and one or more user names or email addresses.
Use **Import CSV...** to load recipients from a CSV with a `PrimarySmtpAddress`,
`EmailAddress`, `Email`, `UserPrincipalName`, or `UPN` column. Imported addresses
are merged with the entries already typed and deduplicated. The export tool's
CSV can be imported directly; choose the roles to add or remove in this tool,
then select **Check users** and choose a specific match for ambiguous names.
Then select
**Member**, **Owner**, or both and **Add roles** or **Remove roles**. Only user mailboxes are offered,
matching the mailbox permission manager's user lookup; guests, contacts, and
nested groups are not offered as users to manage.

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

For removal, existing assignments are read before any changes. Absent roles are
reported as **Not present** without issuing a removal command. When removing
Microsoft 365 membership, the confirmation explicitly explains that existing
ownership will also be removed first, even if **Owner** is not selected. If
ownership removal fails, membership is not removed. Removing only **Owner**
leaves membership unchanged. Distribution-list roles are independent: only the
selected roles are removed, and other members and owners are preserved. The
last owner cannot be removed; add another owner first.

Every write requires confirmation. Results and Exchange warnings are shown per
user and role. Successful changes are not rolled back if another change fails;
for example, membership remains if a subsequent ownership change fails.
Likewise, ownership stays removed if a subsequent membership removal fails.

The signed-in account needs permission to run `Get-EXORecipient`, `Get-Recipient`,
`Get-DistributionGroup`, `Get-DistributionGroupMember`, `Get-UnifiedGroup`,
`Get-UnifiedGroupLinks`, `Add-DistributionGroupMember`, `Set-DistributionGroup`,
`Add-UnifiedGroupLinks`, `Remove-DistributionGroupMember`, and
`Remove-UnifiedGroupLinks`. For distribution lists you do not own, select
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
