#requires -Version 7.6

[CmdletBinding()]
param()

& (Join-Path $PSScriptRoot 'Add-SharedMailboxPermission.ps1') -CreateSharedMailbox