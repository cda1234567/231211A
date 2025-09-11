param(
    [string]$Base = "�j������s",
    [string]$Prefix = "daily",
    [switch]$AutoStash
)

$ErrorActionPreference = "Stop"

# Determine repo root robustly (avoid encoding issues with non-ASCII paths)
$repoRoot = Split-Path -Path $PSScriptRoot -Parent
if (-not (Test-Path $repoRoot)) {
    throw "�䤣���x�s�w�ڥؿ��G$repoRoot"
}
Set-Location $repoRoot

# Guard: in-progress operations (merge/rebase/cherry-pick) prevent branch switch
$gitDir = (& git rev-parse --git-dir 2>$null)
if (-not $gitDir) { throw "���ؿ����O Git �x�s�w" }
$mergeHead = Join-Path $gitDir "MERGE_HEAD"
$rebaseApply = Join-Path $gitDir "rebase-apply"
$rebaseMerge = Join-Path $gitDir "rebase-merge"
$cherryPick = Join-Path $gitDir "CHERRY_PICK_HEAD"
if (Test-Path $mergeHead -or Test-Path $rebaseApply -or Test-Path $rebaseMerge -or Test-Path $cherryPick) {
    throw "������i�椤�� Git �@�~�]merge/rebase/cherry-pick�^�C�Х������Τ���Ggit merge --continue/--abort�Bgit rebase --continue/--abort�C"
}

# Optionally stash local changes
$didStash = $false
$dirty = $false
& git diff --quiet --ignore-submodules -- . 2>$null; if ($LASTEXITCODE -ne 0) { $dirty = $true }
& git diff --cached --quiet --ignore-submodules -- . 2>$null; if ($LASTEXITCODE -ne 0) { $dirty = $true }
if ($dirty) {
    if ($AutoStash) {
        Write-Host "�u�@�ؿ����ܧ�A����۰� stash" -ForegroundColor Yellow
        & git stash push -u -m "auto-stash before new daily branch" | Out-Null
        $didStash = $true
    } else {
        throw "�u�@�ؿ��μȦs�Ϧ��������ܧ�A�Х�����/���A�Υ[�W -AutoStash ��A�աC"
    }
}

# Fetch latest
& git fetch origin --prune

# Ensure base branch locally
$hasLocalBase = $false
& git rev-parse --verify $Base 2>$null | Out-Null
if ($LASTEXITCODE -eq 0) { $hasLocalBase = $true }

if (-not $hasLocalBase) {
    Write-Host "���a�䤣�� $Base�A��ΰl�� origin/$Base �إ�" -ForegroundColor Yellow
    & git switch -c $Base --track "origin/$Base"
} else {
    & git switch $Base
}

# Update base
& git pull --ff-only origin $Base

# Build daily branch name
$today = Get-Date -Format "yyyyMMdd"
$branch = "$Prefix/$today"

# Check remote/local existence
$remoteRef = (& git ls-remote --heads origin $branch 2>$null)
$remoteExists = ($remoteRef -and ($remoteRef.Trim().Length -gt 0))

& git rev-parse --verify $branch 2>$null | Out-Null
$localExists = $LASTEXITCODE -eq 0

if ($localExists -and $remoteExists) {
    Write-Host "����w�s�b�]���a�P���ݡ^�G$branch�A��������" -ForegroundColor Cyan
    & git switch $branch
}
elseif ($localExists -and -not $remoteExists) {
    Write-Host "���a�w�s�b�G$branch�A�]�w�l�ܨñ��e" -ForegroundColor Cyan
    & git switch $branch
    & git push -u origin $branch
}
elseif (-not $localExists -and $remoteExists) {
    Write-Host "���ݤw�s�b�G$branch�A�إ߰l�ܤ���ä���" -ForegroundColor Cyan
    & git switch -c $branch --track "origin/$branch"
}
else {
    Write-Host "�إߨñ��e�G$branch" -ForegroundColor Green
    & git switch -c $branch
    & git push -u origin $branch
}

# Try to restore stashed changes if any
if ($didStash) {
    Write-Host "�٭���e�� stash" -ForegroundColor Yellow
    & git stash pop 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "stash pop �o�ͽĬ�A�Ф�ʳB�z" -ForegroundColor Red
    }
}

Write-Host "�����A�ثe����G$(git branch --show-current)" -ForegroundColor Green
