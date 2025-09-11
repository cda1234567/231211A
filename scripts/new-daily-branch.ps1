param(
    [string]$Base = "大版本更新",
    [string]$Prefix = "daily",
    [switch]$AutoStash
)

$ErrorActionPreference = "Stop"

# Determine repo root robustly (avoid encoding issues with non-ASCII paths)
$repoRoot = Split-Path -Path $PSScriptRoot -Parent
if (-not (Test-Path $repoRoot)) {
    throw "找不到儲存庫根目錄：$repoRoot"
}
Set-Location $repoRoot

# Guard: in-progress operations (merge/rebase/cherry-pick) prevent branch switch
$gitDir = (& git rev-parse --git-dir 2>$null)
if (-not $gitDir) { throw "此目錄不是 Git 儲存庫" }
$mergeHead = Join-Path $gitDir "MERGE_HEAD"
$rebaseApply = Join-Path $gitDir "rebase-apply"
$rebaseMerge = Join-Path $gitDir "rebase-merge"
$cherryPick = Join-Path $gitDir "CHERRY_PICK_HEAD"
if (Test-Path $mergeHead -or Test-Path $rebaseApply -or Test-Path $rebaseMerge -or Test-Path $cherryPick) {
    throw "偵測到進行中的 Git 作業（merge/rebase/cherry-pick）。請先完成或中止：git merge --continue/--abort、git rebase --continue/--abort。"
}

# Optionally stash local changes
$didStash = $false
$dirty = $false
& git diff --quiet --ignore-submodules -- . 2>$null; if ($LASTEXITCODE -ne 0) { $dirty = $true }
& git diff --cached --quiet --ignore-submodules -- . 2>$null; if ($LASTEXITCODE -ne 0) { $dirty = $true }
if ($dirty) {
    if ($AutoStash) {
        Write-Host "工作目錄有變更，執行自動 stash" -ForegroundColor Yellow
        & git stash push -u -m "auto-stash before new daily branch" | Out-Null
        $didStash = $true
    } else {
        throw "工作目錄或暫存區有未提交變更，請先提交/丟棄，或加上 -AutoStash 後再試。"
    }
}

# Fetch latest
& git fetch origin --prune

# Ensure base branch locally
$hasLocalBase = $false
& git rev-parse --verify $Base 2>$null | Out-Null
if ($LASTEXITCODE -eq 0) { $hasLocalBase = $true }

if (-not $hasLocalBase) {
    Write-Host "本地找不到 $Base，改用追蹤 origin/$Base 建立" -ForegroundColor Yellow
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
    Write-Host "分支已存在（本地與遠端）：$branch，直接切換" -ForegroundColor Cyan
    & git switch $branch
}
elseif ($localExists -and -not $remoteExists) {
    Write-Host "本地已存在：$branch，設定追蹤並推送" -ForegroundColor Cyan
    & git switch $branch
    & git push -u origin $branch
}
elseif (-not $localExists -and $remoteExists) {
    Write-Host "遠端已存在：$branch，建立追蹤分支並切換" -ForegroundColor Cyan
    & git switch -c $branch --track "origin/$branch"
}
else {
    Write-Host "建立並推送：$branch" -ForegroundColor Green
    & git switch -c $branch
    & git push -u origin $branch
}

# Try to restore stashed changes if any
if ($didStash) {
    Write-Host "還原先前的 stash" -ForegroundColor Yellow
    & git stash pop 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "stash pop 發生衝突，請手動處理" -ForegroundColor Red
    }
}

Write-Host "完成，目前分支：$(git branch --show-current)" -ForegroundColor Green
