# _notice.ps1 — 运行横幅与硬上限（由各入口脚本 dot-source）
# 引证溯源 / yinzheng-suyuan  v3
# 这些限制写死在代码里，不提供关闭参数。

$script:ChengName = '引证溯源 (yinzheng-suyuan)'
$script:MaxCnkiLimit = 200      # 单次抓取条数上限
$script:MaxCnkiTop = 50         # 单次保留条数上限
$script:MaxCnkiPages = 8        # 翻页深度上限
$script:MaxFetchItems = 60      # 单次批量下载上限
$script:MinDelayMs = 600        # 请求最小间隔

function Show-Notice {
    param([string]$Step = '')
    $line = '─' * 62
    Write-Host $line -ForegroundColor DarkGray
    Write-Host (" {0}  {1}" -f $script:ChengName, $Step) -ForegroundColor Cyan
    Write-Host " 仅限个人学习研究 · 不提供文献正文 · 与 gradpaper/CNKI/NCPSSD 无隶属关系" -ForegroundColor DarkGray
    Write-Host " 禁止：绕过付费墙/验证码、并发或高频抓取、囤积转售、二次分发正文（详见 USAGE-POLICY.md）" -ForegroundColor DarkGray
    Write-Host $line -ForegroundColor DarkGray
}

function Assert-Cap {
    param(
        [Parameter(Mandatory = $true)][int]$Value,
        [Parameter(Mandatory = $true)][int]$Max,
        [Parameter(Mandatory = $true)][string]$Name
    )
    if ($Value -gt $Max) {
        throw "$Name = $Value 超出内置上限 $Max。这是防滥用硬限制（见 USAGE-POLICY.md 第四节），不接受通过改参数放大。"
    }
}

function Assert-Terms {
    param([switch]$Accepted, [string]$What = '批量取全文')
    if (-not $Accepted) {
        Write-Host ""
        Write-Host "本步骤为「$What」，需你先确认用途合规。" -ForegroundColor Yellow
        Write-Host "  允许：个人论文/综述自用；禁止：转售、代下、建库、二次分发正文、绕过付费墙。" -ForegroundColor Yellow
        Write-Host "  确认后加参数 -AcceptTerms 重新运行；条款见 USAGE-POLICY.md 第一、二节。" -ForegroundColor Yellow
        exit 2
    }
}
