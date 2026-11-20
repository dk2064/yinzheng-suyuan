# 引证溯源 (yinzheng-suyuan) · 安装脚本（仓库扁平结构版）
#
# 用法（在 clone 下来的仓库根目录里跑）：
#   pwsh -NoProfile -File install.ps1                     # 装到 %USERPROFILE%\.agents\skills
#   pwsh -NoProfile -File install.ps1 -Target D:\ws\.agents\skills
#   pwsh -NoProfile -File install.ps1 -Force -WithDeps    # 覆盖已装版本，并顺手装缺的 python 包

[CmdletBinding()]
param(
    [string]$Target = (Join-Path $env:USERPROFILE ".agents\skills"),
    [switch]$Force,
    [switch]$WithDeps
)

$ErrorActionPreference = "Stop"
$src = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not (Test-Path (Join-Path $src "SKILL.md"))) {
    throw "找不到 SKILL.md：请在仓库根目录运行 install.ps1"
}

$dst = Join-Path $Target "yinzheng-suyuan"
Write-Host "源： $src"
Write-Host "目标：$dst"

if (Test-Path $dst) {
    if (-not $Force) {
        $v = (Get-Content (Join-Path $dst "VERSION") -ErrorAction SilentlyContinue | Select-Object -First 1)
        Write-Host "目标已存在（版本 $v）。加 -Force 覆盖。" -ForegroundColor Yellow
        exit 1
    }
    Remove-Item $dst -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $Target | Out-Null
# 只复制技能本体，跳过仓库元文件
$skip = @('.git', '.gitignore', '.github', 'install.ps1')
Get-ChildItem $src -Force | Where-Object { $skip -notcontains $_.Name } | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $dst $_.Name) -Recurse -Force
}

$ver = (Get-Content (Join-Path $dst "VERSION") -ErrorAction SilentlyContinue | Select-Object -First 1)
Write-Host "已安装 引证溯源 (yinzheng-suyuan) $ver" -ForegroundColor Green

# 依赖体检
$check = Join-Path $dst "scripts\check_env.ps1"
if (Test-Path $check) {
    Write-Host ""
    Write-Host "== 依赖体检 ==" -ForegroundColor Cyan
    if ($WithDeps) { & $check -Fix } else { & $check }
}

Write-Host ""
Write-Host "使用前请阅读 USAGE-POLICY.md（允许/禁止用途、内置上限）：$dst\USAGE-POLICY.md" -ForegroundColor Yellow
Write-Host "下一步：" -ForegroundColor Cyan
Write-Host "  1) 重开会话，技能列表应出现 引证溯源 (yinzheng-suyuan)"
Write-Host "  2) 浏览器里登录 gradpaper.com（A 入口）、kns.cnki.net（按被引筛时）、www.ncpssd.cn（取全文）"
Write-Host "  3) 选择器自检（不消耗生成次数）："
Write-Host "     pwsh -NoProfile -File `"$dst\scripts\gradpaper_harvest.ps1`" -Title 测试 -Degree 本科 -Discipline 交叉学科 -SubDiscipline AI自动识别 -Major 计算机科学与技术 -School `"某某大学`" -OutDir .\out -DryRun"
Write-Host "  详见 $dst\INSTALL.md"
