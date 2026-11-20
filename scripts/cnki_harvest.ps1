# cnki_harvest.ps1 — 在知网（kns.cnki.net）按题名/主题自检索，抓结果并筛选，产出与 gradpaper 同格式的 refs.json
#
# 与 gradpaper_harvest.ps1 的输出结构一致（i/type/title/abs/authors/journal/year/keywords），
# 因此后面的 ncpssd_fetch.ps1、add_audit_page.ps1 不用改就能直接吃这份清单。
#
# 用法：
#   pwsh -NoProfile -File cnki_harvest.ps1 -Query "大学生心理健康 模型" -Limit 60 -Top 20 -OutDir .\out
#   pwsh -NoProfile -File cnki_harvest.ps1 -Query "SAM-MH 心理健康" -Field 篇名 -YearFrom 2018 -Keep 期刊 -MinCited 5
#   pwsh -NoProfile -File cnki_harvest.ps1 -Query "..." -Merge .\out\refs.json     # 与 gradpaper 结果合并去重
#
# 依赖：sctl（浏览器已登录知网；未登录也能看到题录，但看不到被引/下载数）

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Query,
    [ValidateSet('主题', '篇名', '关键词', '摘要', '全文', '作者')][string]$Field = '主题',
    [int]$Limit = 40,
    [int]$Top = 20,
    [int]$YearFrom = 0,
    [int]$YearTo = 0,
    [int]$MinCited = 0,
    [double]$MinScore = 0.5,
    [ValidateSet('全部', '期刊', '学位')][string]$Keep = '全部',
    [ValidateSet('相关度', '被引', '时间', '下载')][string]$Sort = '相关度',
    [string[]]$Drop = @(),
    [string]$Merge = '',
    [string]$OutDir = '.',
    [string]$SctlExe = '',
    [switch]$WithDetail,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
. "$PSScriptRoot\_notice.ps1"
Show-Notice '入口 B · 知网检索与筛选'
Assert-Cap -Value $Limit -Max $script:MaxCnkiLimit -Name 'Limit'
Assert-Cap -Value $Top -Max $script:MaxCnkiTop -Name 'Top'
if (-not $SctlExe) {
    $SctlExe = @(
        (Join-Path $env:LOCALAPPDATA 'sctl\bin\sctl.exe'),
        (Join-Path $env:USERPROFILE 'AppData\Local\sctl\bin\sctl.exe')
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $SctlExe) {
        $cmd = Get-Command sctl -ErrorAction SilentlyContinue
        if ($cmd) { $SctlExe = $cmd.Source }
    }
    if (-not $SctlExe) { throw "找不到 sctl（sctl.exe）。请安装 ScriptCat 的 sctl CLI，或用 -SctlExe 指定路径。" }
}
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
$OutDir = (Resolve-Path $OutDir).Path

$korder = @{ '主题' = 'SU'; '篇名' = 'TI'; '关键词' = 'KY'; '摘要' = 'AB'; '全文' = 'FT'; '作者' = 'AU' }[$Field]
$sortId = @{ '相关度' = 'FFD'; '被引' = 'CF'; '时间' = 'PT'; '下载' = 'DFR' }[$Sort]

function Eval-Js {
    param([int]$Tab, [string]$Js)
    $r = & $SctlExe page eval $Js --tab $Tab --timeout 40s 2>&1
    $raw = ($r | Where-Object { $_ -match '^\s*tab \d+:' } | Select-Object -Last 1)
    if (-not $raw) { $raw = ($r | Select-Object -Last 1) }
    $json = $raw -replace '^\s*tab \d+:\s*', ''
    if ($json -match '^"') { try { $json = ($json | ConvertFrom-Json) } catch {} }
    try { return ($json | ConvertFrom-Json -AsHashtable) } catch { return @{ ok = $false; raw = "$json" } }
}

function Wait-Until {
    param([int]$Tab, [string]$Js, [int]$Seconds = 40, [string]$What = 'condition')
    $deadline = (Get-Date).AddSeconds($Seconds)
    $last = $null
    while ((Get-Date) -lt $deadline) {
        $last = Eval-Js -Tab $Tab -Js $Js
        if ($last.ok -eq $true) { return $last }
        Start-Sleep -Milliseconds 800
    }
    return @{ ok = $false; err = "timeout: $What" }
}

$EXTRACT = @'
(() => {
  const rows = [...document.querySelectorAll('table.result-table-list tbody tr')];
  const cell = (tr, cls) => { const c = tr.querySelector('td.' + cls); return c ? c.innerText.trim() : ''; };
  const out = rows.map(tr => {
    const nameTd = tr.querySelector('td.name');
    const a = nameTd ? nameTd.querySelector('a') : null;
    return {
      title: a ? a.innerText.trim() : (nameTd ? nameTd.innerText.trim() : ''),
      url: a ? (a.getAttribute('href') || '') : '',
      authors: cell(tr, 'author').replace(/\s+/g, ' '),
      journal: cell(tr, 'source'),
      date: cell(tr, 'date'),
      db: cell(tr, 'data'),
      cited: cell(tr, 'quote'),
      dl: cell(tr, 'download')
    };
  }).filter(r => r.title);
  return { ok: out.length > 0, n: out.length, rows: out, totalText: (document.querySelector('.total-count') || {innerText:''}).innerText.trim().slice(0,40) };
})()
'@

# 1) 开标签并检索
$q = [uri]::EscapeDataString($Query)
$url = "https://kns.cnki.net/kns8s/defaultresult/index?kw=$q&korder=$korder"
$tabLine = & $SctlExe tabs open $url --background 2>&1 | Select-Object -Last 1
$tab = [int]([regex]::Match($tabLine, '(\d{6,})').Groups[1].Value)
if (-not $tab) { throw "开标签失败：$tabLine" }
Write-Host "[tab] $tab  查询：$Query（$Field）"
Start-Sleep -Seconds 4

$ready = Wait-Until -Tab $tab -Seconds 40 -What '结果表' -Js @'
(() => {
  const n = document.querySelectorAll('table.result-table-list tbody tr').length;
  const blocked = /验证码|安全验证|异常访问/.test(document.body.innerText);
  return { ok: n > 0, n: n, blocked: blocked, title: document.title };
})()
'@
if (-not $ready.ok) { throw "知网结果未加载（blocked=$($ready.blocked)），可能在要验证码或未登录够权限" }
Write-Host "[1/5] $($ready.title) — $(if ($ready.totalText) { $ready.totalText } else { "$($ready.n) 条/页" })"

# 2) 排序（知网默认可能是「发表时间」，所以显式点一次，保证可复现）
$js = @"
(() => {
  const li = document.getElementById('$sortId');
  if (!li) return { ok: false, err: 'no-sort' };
  li.click();
  return { ok: true, cur: li.className };
})()
"@
$r = Eval-Js -Tab $tab -Js $js
Write-Host "[2/5] 排序：$Sort（$($r.ok)）"
Start-Sleep -Seconds 3
[void](Wait-Until -Tab $tab -Seconds 25 -What '排序后结果' -Js '(() => ({ ok: document.querySelectorAll("table.result-table-list tbody tr").length > 0 }))()')

# 3) 翻页收集
$all = @()
$seen = @{}
$page = 1
while ($all.Count -lt $Limit -and $page -le 8) {
    $d = Eval-Js -Tab $tab -Js $EXTRACT
    if (-not $d.ok) { break }
    foreach ($row in $d.rows) {
        $k = ($row.title -replace '[\s\p{P}]', '')
        if ($seen.ContainsKey($k)) { continue }
        $seen[$k] = 1
        $all += $row
    }
    Write-Host "      第 $page 页：累计 $($all.Count) 条"
    if ($all.Count -ge $Limit) { break }
    $next = Eval-Js -Tab $tab -Js @'
(() => {
  const b = document.getElementById('PageNext');
  if (!b || b.className.match(/disabled/)) return { ok: false, err: 'no-next' };
  b.click();
  return { ok: true };
})()
'@
    if (-not $next.ok) { break }
    $page++
    Start-Sleep -Seconds 3
    [void](Wait-Until -Tab $tab -Seconds 25 -What '下一页' -Js '(() => ({ ok: document.querySelectorAll("table.result-table-list tbody tr").length > 0 }))()')
}
Write-Host "[3/5] 抓取 $($all.Count) 条"

# 4) 筛选（规则全部落到 screening.tsv，人工可复核）
$tokens = @($Query -split '[\s,，、;；/|]+' | Where-Object { $_.Trim().Length -gt 0 })
# 查询本身是一个长串时（如整条题名），先去掉无区分度的功能词，再用字符 bigram 覆盖率算相似度
$stopPhrases = @('基于', '研究', '分析', '探讨', '初探', '应用', '及其', '以及', '的', '对', '与', '和', '中', '以')
$bigrams = @()
if ($tokens.Count -eq 1 -and $tokens[0].Length -gt 6) {
    $s = $tokens[0] -replace '[\s\p{P}]', ''
    foreach ($w in $stopPhrases) { $s = $s.Replace($w, ' ') }
    $s = $s -replace '\s+', ''
    for ($i = 0; $i -lt $s.Length - 1; $i++) { $bigrams += $s.Substring($i, 2) }
    $bigrams = @($bigrams | Select-Object -Unique)
    Write-Host "      长查询：用 bigram 相似度（去功能词后 $($bigrams.Count) 个二元组：$($s)）"
}
$existing = @{}
if ($Merge -and (Test-Path $Merge)) {
    $m = Get-Content $Merge -Raw -Encoding utf8 | ConvertFrom-Json
    foreach ($it in $m) { $existing[($it.title -replace '[\s\p{P}]', '')] = $it }
    Write-Host "      合并基准：$Merge（$($existing.Count) 条）"
}

$screened = foreach ($row in $all) {
    $hits = @($tokens | Where-Object { $row.title -like "*$_*" }).Count
    if ($bigrams.Count -gt 0) {
        $t = $row.title -replace '[\s\p{P}]', ''
        $hit = @($bigrams | Where-Object { $t.Contains($_) }).Count
        $score = [math]::Round($hit / $bigrams.Count, 2)
        $exact = $false
    } else {
        $score = if ($tokens.Count) { [math]::Round($hits / $tokens.Count, 2) } else { 1 }
        $exact = ($tokens.Count -gt 1 -and $hits -eq $tokens.Count)
    }
    $year = 0; [void][int]::TryParse(($row.date -split '-')[0], [ref]$year)
    $cited = 0; [void][int]::TryParse(($row.cited -replace '[^\d]', ''), [ref]$cited)
    $dl = 0; [void][int]::TryParse(($row.dl -replace '[^\d]', ''), [ref]$dl)
    $type = if ($row.db -match '硕士|博士|学位') { '学位论文' } else { '学术期刊' }

    $why = @()
    if ($score -lt $MinScore) { $why += "相关性<$MinScore" }
    if ($YearFrom -gt 0 -and $year -gt 0 -and $year -lt $YearFrom) { $why += "早于$YearFrom" }
    if ($YearTo -gt 0 -and $year -gt $YearTo) { $why += "晚于$YearTo" }
    if ($MinCited -gt 0 -and $cited -lt $MinCited) { $why += "被引<$MinCited" }
    if ($Keep -eq '期刊' -and $type -ne '学术期刊') { $why += '非期刊' }
    if ($Keep -eq '学位' -and $type -ne '学位论文') { $why += '非学位' }
    foreach ($w in $Drop) { if ($w -and $row.title -like "*$w*") { $why += "含排除词:$w" } }
    $np = ($row.title -replace '[\s\p{P}]', '')
    if ($existing.ContainsKey($np)) { $why += '与已有清单重复' }

    [pscustomobject]@{
        title   = $row.title
        authors = $row.authors
        journal = $row.journal
        year    = $(if ($year) { "$year" } else { ($row.date -split '-')[0] })
        type    = $type
        db      = $row.db
        cited   = $cited
        dl      = $dl
        score   = $score
        url     = $row.url
        keep    = ($why.Count -eq 0)
        why     = ($why -join ';')
    }
}

$sorted = $screened | Sort-Object @{ Expression = 'score'; Descending = $true }, @{ Expression = 'cited'; Descending = $true }
$kept = @($sorted | Where-Object { $_.keep } | Select-Object -First $Top)
Write-Host "[4/5] 通过筛选 $(@($sorted | Where-Object { $_.keep }).Count) 条，取前 $($kept.Count) 条"

# 4.5) 可选：逐个打开详情页补摘要/关键词/DOI（后续核验与审核页要用）
$detail = @{}
if ($WithDetail -and $kept.Count -gt 0) {
    Write-Host "[4.5] 补详情：$($kept.Count) 条（每条约 2-3 秒）"
    for ($i = 0; $i -lt $kept.Count; $i++) {
        $u = $kept[$i].url
        if (-not $u) { continue }
        & $SctlExe page goto $u --tab $tab 2>&1 | Out-Null
        Start-Sleep -Milliseconds 1200
        $d = Wait-Until -Tab $tab -Seconds 15 -What '详情页' -Js @'
(() => {
  const a = document.querySelector('#ChDivSummary') || document.querySelector('.abstract-text');
  const kw = [...document.querySelectorAll('.keywords a')].map(e => e.innerText.trim().replace(/[;；]\s*$/, ''));
  if (!a && kw.length === 0) return { ok: false };
  const doi = (document.body.innerText.match(/DOI[:：]\s*(10\.[^\s，;]+)/) || [])[1] || '';
  return { ok: true, abs: a ? a.innerText.trim().replace(/\s+/g, ' ') : '', kw: kw.slice(0, 8), doi: doi };
})()
'@
        if ($d.ok) {
            $detail[$i] = $d
            Write-Host "      $($i+1)/$($kept.Count) 摘要 $($d.abs.Length) 字，关键词 $($d.kw.Count) 个"
        } else {
            Write-Host "      $($i+1)/$($kept.Count) 详情页没取到（跳过）" -ForegroundColor DarkYellow
        }
        Start-Sleep -Milliseconds 600
    }
}

# 5) 落盘
$rows = for ($i = 0; $i -lt $kept.Count; $i++) {
    $k = $kept[$i]
    $dt = $detail[$i]
    [pscustomobject]@{
        i        = $i + 1
        type     = $k.type
        title    = $k.title
        abs      = $(if ($dt) { $dt.abs } else { '' })
        authors  = $k.authors
        journal  = $k.journal
        year     = $k.year
        keywords = $(if ($dt -and $dt.kw) { @($dt.kw) } else { @() })
        source   = 'CNKI'
        cited    = $k.cited
        doi      = $(if ($dt) { $dt.doi } else { '' })
        url      = $k.url
    }
}
if ($Merge -and $existing.Count -gt 0) {
    $seq = $rows.Count
    foreach ($kv in $existing.GetEnumerator()) {
        $it = $kv.Value
        $seq++
        $rows += [pscustomobject]@{
            i = $seq; type = $it.type; title = $it.title; abs = $it.abs; authors = $it.authors
            journal = $it.journal; year = $it.year; keywords = $it.keywords; source = 'GradPaper'
            cited = 0; url = ''
        }
    }
}

if ($DryRun) {
    $sorted | Select-Object -First 15 keep, score, cited, year, title, journal | Format-Table -AutoSize
    Write-Host "[DryRun] 未落盘。标签 $tab 保留。"
    return
}

$rows | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $OutDir 'refs.json') -Encoding utf8
$lines = foreach ($r in $rows) { "$($r.i)|$($r.type)|$($r.title)|$($r.authors)|$($r.journal)|$($r.year)" }
$lines | Set-Content (Join-Path $OutDir 'refs_list.txt') -Encoding utf8
$sorted | ForEach-Object -Begin { $n = 0 } -Process {
    $n++
    [pscustomobject]@{ 序号 = $n; 保留 = $(if ($_.keep) { 'Y' } else { 'N' }); 评分 = $_.score; 被引 = $_.cited; 下载 = $_.dl; 类型 = $_.type; 库 = $_.db; 年份 = $_.year; 题名 = $_.title; 作者 = $_.authors; 来源 = $_.journal; 未通过原因 = $_.why; 链接 = $_.url }
} | Export-Csv (Join-Path $OutDir 'screening.tsv') -Delimiter "`t" -Encoding utf8

Write-Host "[5/5] 写出 $($rows.Count) 条 → refs.json / refs_list.txt / screening.tsv"
Write-Host "[tab] 检索用标签 $tab，用完可关：sctl tabs close $tab"
