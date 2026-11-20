# check_env.ps1 — yinzheng-suyuan 依赖体检
#
# 用法：pwsh -NoProfile -File check_env.ps1 [-Fix]
#   -Fix  缺少 python 包时自动 pip install

[CmdletBinding()]
param(
    [switch]$Fix
)

$ErrorActionPreference = "Continue"
$rows = @()
function Add-Row($name, $ok, $detail) {
    $script:rows += [pscustomobject]@{ 项目 = $name; 状态 = $(if ($ok) { "OK" } else { "缺" }); 说明 = $detail }
}

# PowerShell
$psv = $PSVersionTable.PSVersion
Add-Row "PowerShell" ($psv.Major -ge 7) "$psv"

# Python
$py = (Get-Command python -ErrorAction SilentlyContinue).Source
if ($py) {
    $pyv = (& python -V 2>&1) -join ""
    Add-Row "Python" $true "$pyv ($py)"
} else {
    Add-Row "Python" $false "未找到 python；装 3.10+ 并加入 PATH"
}

# Python 包
$mods = @{
    "pypdf"      = "pypdf"
    "docx"       = "python-docx"
    "pypdfium2"  = "pypdfium2（可选，仅预览用）"
}
foreach ($m in $mods.Keys) {
    $probe = & python -c "import $m" 2>&1
    $ok = ($LASTEXITCODE -eq 0)
    if (-not $ok -and $Fix -and $m -ne "pypdfium2") {
        Write-Host "  安装 $($mods[$m]) ..." -ForegroundColor DarkGray
        & python -m pip install --quiet $m | Out-Null
        & python -c "import $m" 2>&1 | Out-Null
        $ok = ($LASTEXITCODE -eq 0)
    }
    Add-Row $mods[$m] $ok $(if ($ok) { "已安装" } else { "pip install $m" })
}

# sctl
$sctl = (Get-Command sctl -ErrorAction SilentlyContinue).Source
if (-not $sctl) {
    $guess = Join-Path $env:LOCALAPPDATA "sctl\bin\sctl.exe"
    if (Test-Path $guess) { $sctl = $guess }
}
if ($sctl) {
    $st = (& $sctl status -o json 2>&1) -join "`n"
    $daemonVer = "未知"
    $online = $false
    try {
        $o = $st | ConvertFrom-Json -AsHashtable
        $daemonVer = $o.daemonVersion
        $online = @($o.browsers | Where-Object { $_.online }).Count -gt 0
    } catch {
        $online = ($st -match "online")
    }
    Add-Row "sctl" $true "daemon $daemonVer ($sctl)"
    Add-Row "浏览器 daemon" $online $(if ($online) { "已连上浏览器" } else { "启动 ScriptCat sctl 服务并保持浏览器打开，sctl status 查看" })
} else {
    Add-Row "sctl" $false "未找到 sctl；装 ScriptCat 配套 CLI"
}

# docx -> pdf 导出器
$word = Test-Path "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
$soffice = @(
    "C:\Program Files\LibreOffice\program\soffice.exe",
    "C:\Program Files (x86)\LibreOffice\program\soffice.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $soffice) { $soffice = (Get-Command soffice -ErrorAction SilentlyContinue).Source }
Add-Row "docx→pdf" ($word -or [bool]$soffice) $(if ($word) { "Microsoft Word" } elseif ($soffice) { $soffice } else { "装 Word 或 LibreOffice，否则只能手动导出 docx 后单跑 merge" })

# 字体（审核页用微软雅黑）
$font = Test-Path "$env:WINDIR\Fonts\msyh.ttc"
Add-Row "字体 msyh.ttc" $font $(if ($font) { "微软雅黑" } else { "缺微软雅黑：审核页会退回默认字体，排版可能溢出" })

# 条款文件（必须随包分发）
$policy = Join-Path (Split-Path -Parent $PSScriptRoot) "USAGE-POLICY.md"
$disc = Join-Path (Split-Path -Parent $PSScriptRoot) "DISCLAIMER.md"
Add-Row "条款文件" ((Test-Path $policy) -and (Test-Path $disc)) "USAGE-POLICY.md / DISCLAIMER.md（缺则视为非法副本）"

$rows | Format-Table -AutoSize

$bad = @($rows | Where-Object { $_.状态 -eq "缺" })
if ($bad.Count -eq 0) {
    Write-Host "全部就绪，可以跑 gradpaper_harvest.ps1 -DryRun 验证选择器。" -ForegroundColor Green
} else {
    Write-Host "有 $($bad.Count) 项待处理：$((($bad | ForEach-Object { $_.项目 }) -join '、'))" -ForegroundColor Yellow
}
