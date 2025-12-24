param (
    [Parameter(Mandatory)]
    [string]$Mailbox,

    [Parameter(Mandatory)]
    [string]$UsersFile,

    [switch]$DryRun,
    [switch]$Apply
)

# --- LOGGING SETUP ---

$LogDir  = ".\logs"
$LogFile = "$LogDir\bulk-mailbox-access.log"

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

function Write-Log {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp | $Message" | Out-File -FilePath $LogFile -Append -Encoding UTF8
}


function Fail($msg) {
    Write-Host "[ERROR] $msg" -ForegroundColor Red
    exit 1
}

# --- SAFETY CHECKS ---

if (-not $DryRun -and -not $Apply) {
    Fail "You must specify -DryRun or -Apply"
}

if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
    Fail "Not connected to Exchange Online. Run Connect-ExchangeOnline first."
}

if (-not (Test-Path $UsersFile)) {
    Fail "Users file not found: $UsersFile"
}

$mailboxObj = Get-Mailbox -Identity $Mailbox -ErrorAction SilentlyContinue
if (-not $mailboxObj) {
    Fail "Shared mailbox '$Mailbox' not found"
}

# --- LOAD & NORMALIZE USERS ---

$raw = Get-Content $UsersFile -Raw

$users = $raw `
    -split '[,\s]+' |
    Where-Object { $_ -and $_.Trim() -ne "" } |
    Sort-Object -Unique

if ($users.Count -eq 0) {
    Fail "Users list is empty"
}

Write-Host ""
Write-Host "Mailbox: $Mailbox" -ForegroundColor Cyan
Write-Host "Users found: $($users.Count)"
Write-Host ""
Write-Log "START | Mailbox=$Mailbox | Users=$($users.Count) | Mode=$(if ($DryRun) { 'DRY-RUN' } else { 'APPLY' })"

# --- CHECK PERMISSIONS ---

$results = @()

foreach ($user in $users) {

    $hasFull = Get-MailboxPermission `
        -Identity $Mailbox `
        -User $user `
        -ErrorAction SilentlyContinue |
        Where-Object { $_.AccessRights -contains "FullAccess" }

    $hasSendAs = Get-RecipientPermission `
        -Identity $Mailbox `
        -Trustee $user `
        -ErrorAction SilentlyContinue |
        Where-Object { $_.AccessRights -contains "SendAs" }

    $results += [PSCustomObject]@{
        User        = $user
        FullAccess = [bool]$hasFull
        SendAs     = [bool]$hasSendAs
    }
}

# --- OUTPUT (DRY-RUN STYLE) ---

foreach ($r in $results) {

    Write-Host $r.User -ForegroundColor White

    if ($r.FullAccess) {
        Write-Host "  FullAccess: HAS" -ForegroundColor Green
    } else {
        Write-Host "  FullAccess: MISSING (will be added)" -ForegroundColor Yellow
    }

    if ($r.SendAs) {
        Write-Host "  SendAs:     HAS" -ForegroundColor Green
    } else {
        Write-Host "  SendAs:     MISSING (will be added)" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Log "CHECK | User=$($r.User) | FullAccess(before)=$($r.FullAccess) | SendAs(before)=$($r.SendAs)"

}

if ($DryRun) {
    Write-Host "[DRY-RUN] No changes were made." -ForegroundColor Cyan
    exit 0
}

# --- CONFIRMATION BEFORE APPLY ---

$toChange = $results | Where-Object {
    -not $_.FullAccess -or -not $_.SendAs
}

if ($toChange.Count -eq 0) {
    Write-Host "No changes required." -ForegroundColor Green
    exit 0
}

Write-Host "Users to be modified: $($toChange.Count)" -ForegroundColor Yellow
$confirm = Read-Host "Type YES to apply changes"

if ($confirm -ne "YES") {
    Write-Host "Operation cancelled." -ForegroundColor Gray
    Write-Log "CANCELLED BY USER"
    exit 0
}
Write-Log "APPLY CONFIRMED | UsersToChange=$($toChange.Count)"

# --- APPLY CHANGES ---

foreach ($r in $toChange) {

    if (-not $r.FullAccess) {
        Write-Host "Granting FullAccess to $($r.User)" -ForegroundColor Cyan
        Write-Log "CHANGE | User=$($r.User) | FullAccess: false -> true"
        Add-MailboxPermission `
            -Identity $Mailbox `
            -User $r.User `
            -AccessRights FullAccess `
            -InheritanceType All `
            -AutoMapping $false `
            -Confirm:$false
    }

    if (-not $r.SendAs) {
        Write-Host "Granting SendAs to $($r.User)" -ForegroundColor Cyan
        Write-Log "CHANGE | User=$($r.User) | SendAs: false -> true"
        Add-RecipientPermission `
            -Identity $Mailbox `
            -Trustee $r.User `
            -AccessRights SendAs `
            -Confirm:$false
    }
}

Write-Host "Completed successfully." -ForegroundColor Green
Write-Log "END | Completed successfully"