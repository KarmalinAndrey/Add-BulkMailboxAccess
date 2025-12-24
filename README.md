# Bulk Shared Mailbox Access Tool

PowerShell script for safely granting access to a shared mailbox
to a large number of users in Microsoft Exchange Online.

This tool is designed for bulk operations with maximum safety:
no existing permissions are removed or modified.

---

## Purpose

Managing access for many users via Azure Portal is slow and error-prone.

This script allows you to:
- process many users at once
- clearly see who already has access
- add only missing permissions
- avoid accidental changes

---

## Features

- One shared mailbox → many users
- Mandatory **dry-run** mode
- Adds **only missing permissions**
- Never removes existing permissions
- Interactive confirmation before applying changes
- Colorized console output
- Detailed audit logging to file

---

## Input format

Users are provided via a **TXT file**.

The script supports:
- one user per line
- space-separated users
- comma-separated users
- any combination of the above

Example users.txt:
---
```md
ivan.ivanov\@`company.com`  
petr.petrov\@`company.com`
alex.kim\@`company.com`
```

How it works
---
For each user, the script checks:

FullAccess permission

SendAs permission

Then:

If a permission already exists → it is not touched

If a permission is missing → it can be added (only in Apply mode)

The script is idempotent and safe to re-run.
---
Pre-run checklist

Before running the script, make sure:

You are connected to Exchange Online:

Connect-ExchangeOnline


You are running the script in the same PowerShell session
where the Exchange Online connection is active.

You have sufficient permissions to manage shared mailboxes
(same permissions as required in Azure Portal).
---
Usage
1. Dry-run (required)
.\Add-BulkMailboxAccess.ps1 `
  -Mailbox sales@company.com `
  -UsersFile users.txt `
  -DryRun


Dry-run will:

show current permissions

highlight missing permissions

make no changes
---
2. Apply changes
.\Add-BulkMailboxAccess.ps1 `
  -Mailbox sales@company.com `
  -UsersFile users.txt `
  -Apply


Before applying changes, the script will ask for explicit confirmation.