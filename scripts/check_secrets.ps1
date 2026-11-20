# check_secrets.ps1 — 发布/分发前自检：确认包里没混进凭证、个人路径、文献正文
# 用法：
#   pwsh -NoProfile -File scripts\check_secrets.ps1                      # 扫技能目录
#   pwsh -NoProfile -File scripts\check_secrets.ps1 -Path .\dist -Strict # 扫分发包，命中即 exit 1
param(
    [string]$Path = (Split-Path -Parent $PSScriptRoot),
    [switch]$Strict,
    [switch]$Quiet
)
$ErrorActionPreference = 'Stop'

$root = (Resolve-Path $Path).Path
if (-not $Quiet) { Write-Host "扫描：$root" -ForegroundColor Cyan }

$textExt = '.ps1', '.psm1', '.py', '.md', '.json', '.txt', '.yml', '.yaml', '.sh', '.cfg', '.example'

$patterns = [ordered]@{
    '登录态 cookie'   = '(?i)(web_session|accessId|JSESSIONID)\s*=\s*[A-Za-z0-9%_\-\.]{12,}'
    'Cookie 头'       = '(?i)Cookie\s*:\s*[A-Za-z0-9_\-]+='
    'GitHub token'    = 'gh[pousr]_[A-Za-z0-9]{20,}'
    'JWT'             = 'ey[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}'
    '私钥'            = '-----BEGIN [A-Z ]*PRIVATE KEY-----'
    '真实用户名路径'  = '(?i)[A-Z]:\\\\?Users\\\\?(?!<|\$|%|xxx|user)[A-Za-z0-9._-]{2,}'
    '手机号'          = '(?<!\d)1[3-9]\d{9}(?!\d)'
    '邮箱'            = '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}'
}

$hits = New-Object System.Collections.Generic.List[object]
$files = Get-ChildItem -Path $root -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $textExt -contains $_.Extension.ToLower() -and $_.FullName -notmatch '\\\.git\\' }

foreach ($f in $files) {
    $rel = $f.FullName.Substring($root.Length).TrimStart('\')
    $line = 0
    foreach ($l in (Get-Content $f.FullName -Encoding utf8 -ErrorAction SilentlyContinue)) {
        $line++
        foreach ($k in $patterns.Keys) {
            if ($l -match $patterns[$k]) {
                # 文档里作为「反面示例」出现的占位串放过
                if ($l -match '\*\*\*|<你的|示例|placeholder|xxx|示例值|替换') { continue }
                $hits.Add([pscustomobject]@{ 类型 = $k; 文件 = $rel; 行 = $line; 片段 = ($l.Trim() -replace '\s+', ' ').Substring(0, [Math]::Min(100, $l.Trim().Length)) })
            }
        }
    }
}

# 正文类文件：仓库/包内不应有 PDF/DOCX 全文
$docs = Get-ChildItem -Path $root -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -in '.pdf', '.docx', '.caj', '.nh' -and $_.Length -gt 50000 }

if (-not $Quiet) {
    if ($hits.Count -eq 0) { Write-Host "凭证/个人信息：未发现" -ForegroundColor Green }
    else {
        Write-Host "凭证/个人信息：$($hits.Count) 处" -ForegroundColor Yellow
        $hits | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
    }
    if ($docs.Count -eq 0) { Write-Host "文献正文文件：无（正确）" -ForegroundColor Green }
    else {
        Write-Host "文献正文文件：$($docs.Count) 个 —— 分发前必须删除" -ForegroundColor Red
        $docs | ForEach-Object { Write-Host ("  {0}  {1:N0} B" -f $_.Name, $_.Length) -ForegroundColor Red }
    }
}

if ($hits.Count -gt 0 -or $docs.Count -gt 0) {
    if ($Strict) { exit 1 }
    exit 0
}
exit 0
