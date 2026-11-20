# GradPaper 推荐参考文献抓取（sctl 驱动用户已登录的浏览器）
# 用法见 SKILL.md 第一步。所有 Arco 交互都走 page eval，不走 page click。
param(
    [Parameter(Mandatory = $true)][string]$Title,
    [string]$Degree = '本科',
    [string]$Discipline = '交叉学科',
    [string]$SubDiscipline = 'AI自动识别',
    [string]$Major = '计算机科学与技术',
    [string]$School = '',
    [string]$OutDir = '.',
    [string]$SctlExe = '',
    [switch]$DryRun,
    [int]$WaitSeconds = 120
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
. "$PSScriptRoot\_notice.ps1"
Show-Notice '入口 A · GradPaper 推荐文献'

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
$URL = 'https://www.gradpaper.com/dissertation'

function Eval-Js {
    param([int]$Tab, [string]$Js)
    $r = & $SctlExe page eval $Js --tab $Tab 2>&1
    $raw = ($r | Where-Object { $_ -match '^\s*tab \d+:' } | Select-Object -Last 1)
    if (-not $raw) { $raw = ($r | Select-Object -Last 1) }
    $json = $raw -replace '^\s*tab \d+:\s*', ''
    try { return ($json | ConvertFrom-Json -AsHashtable) } catch { return @{ ok = $false; raw = $json } }
}

function Invoke-Select {
    param([int]$Tab, [string]$Placeholder, [string]$Value, [int]$FallbackIndex = -1, [int]$Tries = 12)
    $open = @'
(() => {
  const q = '__PH__';
  const idx = __IDX__;
  const sels = [...document.querySelectorAll('.arco-select')].filter(e => e.offsetParent);
  let t = sels.find(e => { const i = e.querySelector('input'); return i && (i.placeholder || '').includes(q); });
  if (!t && idx >= 0) t = sels[idx];
  if (!t) return { ok: false, err: 'select-not-found', placeholders: sels.map(e => (e.querySelector('input') || {}).placeholder) };
  t.click();
  return { ok: true };
})()
'@.Replace('__PH__', $Placeholder).Replace('__IDX__', "$FallbackIndex")
    $o = Eval-Js -Tab $Tab -Js $open
    if (-not $o.ok) { return $o }
    for ($i = 0; $i -lt $Tries; $i++) {
        Start-Sleep -Milliseconds 350
        $pick = @'
(() => {
  const want = '__VAL__';
  const opts = [...document.querySelectorAll('.arco-select-option')].filter(e => e.offsetParent);
  let hit = opts.find(e => e.innerText.trim() === want) || opts.find(e => e.innerText.trim().includes(want));
  if (!hit) return { ok: false, list: opts.map(e => e.innerText.trim()).slice(0, 40) };
  hit.click();
  return { ok: true, picked: hit.innerText.trim() };
})()
'@.Replace('__VAL__', $Value)
        $p = Eval-Js -Tab $Tab -Js $pick
        if ($p.ok) { return $p }
    }
    return @{ ok = $false; err = 'option-not-found'; want = $Value; list = $p.list }
}

function Wait-Until {
    param([int]$Tab, [string]$Js, [int]$Seconds = 60, [string]$What = 'condition')
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
        $r = Eval-Js -Tab $Tab -Js $Js
        if ($r.ok -eq $true) { return $r }
        Start-Sleep -Milliseconds 900
    }
    return @{ ok = $false; err = "timeout waiting for $What" }
}

# 1) 开标签
$tabLine = & $SctlExe tabs open $URL --background 2>&1 | Select-Object -Last 1
$tab = [int]([regex]::Match($tabLine, '(\d{6,})').Groups[1].Value)
if (-not $tab) { throw "open tab failed: $tabLine" }
Write-Host "[tab] $tab"
Start-Sleep -Seconds 5

$ready = Wait-Until -Tab $tab -Seconds 30 -What 'form' -Js @'
(() => {
  const i = document.querySelector('input[placeholder="请输入标题"]');
  return { ok: !!i, loggedIn: document.body.innerText.includes('订单结果查询') };
})()
'@
if (-not $ready.ok) { throw "表单未就绪，可能未登录 gradpaper" }
if ($ready.loggedIn -ne $true) { Write-Warning "页面没有「订单结果查询」，可能未登录，继续尝试" }

# 2) 标题
$setTitle = @'
(() => {
  const v = __TITLE__;
  const el = document.querySelector('input[placeholder="请输入标题"]');
  if (!el) return { ok: false, err: 'no-title-input' };
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
  el.focus();
  setter.call(el, v);
  el.dispatchEvent(new Event('input', { bubbles: true }));
  el.dispatchEvent(new Event('change', { bubbles: true }));
  return { ok: true, val: el.value };
})()
'@.Replace('__TITLE__', (ConvertTo-Json $Title -Compress))
$r = Eval-Js -Tab $tab -Js $setTitle
if (-not $r.ok) { throw "填标题失败: $($r | ConvertTo-Json -Compress)" }
Write-Host "[1/6] 标题 = $($r.val)"

# 3) 学历
$r = Invoke-Select -Tab $tab -Placeholder '学历' -Value $Degree -FallbackIndex 1
if (-not $r.ok) { throw "选学历失败: $($r | ConvertTo-Json -Compress -Depth 3)" }
Write-Host "[2/6] 学历 = $($r.picked)"
Start-Sleep -Milliseconds 600

# 4) 学科门类（第一级）
$r = Invoke-Select -Tab $tab -Placeholder '学科门类' -Value $Discipline -FallbackIndex 3
if (-not $r.ok) {
    $r = Invoke-Select -Tab $tab -Placeholder '专业大类' -Value $Discipline -FallbackIndex 3
}
if (-not $r.ok) { throw "选学科门类失败: $($r | ConvertTo-Json -Compress -Depth 3)" }
Write-Host "[3/6] 学科门类 = $($r.picked)"
Start-Sleep -Milliseconds 800

# 5) 一级学科 / AI自动识别（第二级）
$r = Invoke-Select -Tab $tab -Placeholder 'AI' -Value $SubDiscipline -FallbackIndex 4
if (-not $r.ok) {
    $r = Invoke-Select -Tab $tab -Placeholder '一级学科' -Value $SubDiscipline -FallbackIndex 4
}
if (-not $r.ok) { throw "选一级学科失败: $($r | ConvertTo-Json -Compress -Depth 3)" }
Write-Host "[4/6] 一级学科 = $($r.picked)"
Start-Sleep -Milliseconds 600

# 6) 专业名称
$setMajor = @'
(() => {
  const v = __VAL__;
  const el = document.querySelector('input[placeholder="请输入专业名称"]');
  if (!el) return { ok: false, err: 'no-major-input' };
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
  el.focus();
  setter.call(el, v);
  el.dispatchEvent(new Event('input', { bubbles: true }));
  el.dispatchEvent(new Event('change', { bubbles: true }));
  return { ok: true, val: el.value };
})()
'@.Replace('__VAL__', (ConvertTo-Json $Major -Compress))
$r = Eval-Js -Tab $tab -Js $setMajor
if (-not $r.ok) { $r = @{ ok = $false } }
Write-Host "[5/6] 专业名称 = $Major （$($r.ok)）"
Start-Sleep -Milliseconds 400

# 7) 学校（搜索型下拉）
$openSchool = @'
(() => {
  const el = [...document.querySelectorAll('input')].find(i => (i.placeholder || '').includes('学校'));
  if (!el) return { ok: false, err: 'no-school-input' };
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
  el.focus();
  setter.call(el, __VAL__);
  el.dispatchEvent(new Event('input', { bubbles: true }));
  el.dispatchEvent(new Event('change', { bubbles: true }));
  el.click();
  return { ok: true };
})()
'@.Replace('__VAL__', (ConvertTo-Json $School -Compress))
$r = Eval-Js -Tab $tab -Js $openSchool
if (-not $r.ok) { throw "学校下拉打不开: $($r | ConvertTo-Json -Compress)" }
Start-Sleep -Seconds 2
$r = Invoke-Select -Tab $tab -Placeholder '学校' -Value "$School-本科" -FallbackIndex -1 -Tries 15
if (-not $r.ok) { $r = Invoke-Select -Tab $tab -Placeholder '学校' -Value $School -FallbackIndex -1 -Tries 15 }
if (-not $r.ok) { throw "选学校失败: $($r | ConvertTo-Json -Compress -Depth 3)" }
Write-Host "[6/6] 学校 = $($r.picked)"

if ($DryRun) {
    Write-Host "[DryRun] 表单已填好，未点下一步。标签 $tab 保留给你检查。"
    return
}

# 8) 下一步
$clickNext = @'
(() => {
  const btns = [...document.querySelectorAll('button')].filter(e => e.offsetParent);
  const b = btns.find(e => e.innerText.trim() === '下一步');
  if (!b) return { ok: false, list: btns.map(e => e.innerText.trim()) };
  b.click();
  return { ok: true };
})()
'@
$r = Eval-Js -Tab $tab -Js $clickNext
if (-not $r.ok) { throw "点下一步失败: $($r | ConvertTo-Json -Compress -Depth 3)" }
Write-Host "[..] 已点下一步，等待生成推荐文献"

$waitItems = @'
(() => {
  const n = document.querySelectorAll('.arco-list-item').length;
  return { ok: n >= 5, n: n };
})()
'@
$r = Wait-Until -Tab $tab -Seconds $WaitSeconds -What 'recommend list' -Js $waitItems
if (-not $r.ok) { throw "等待推荐文献超时（当前 $($r.n) 条）" }
Write-Host "[ok] 推荐文献 $($r.n) 条"

# 9) 提取
$extract = @'
(() => {
  const items = [...document.querySelectorAll('.arco-list-item')];
  const rows = items.map((it, i) => {
    const divs = [...it.querySelectorAll('div')];
    const tags = [...it.querySelectorAll('.arco-tag')].map(e => e.innerText.trim());
    const titleEl = it.querySelector('.font-bold');
    const absEl = divs.find(d => (d.innerText || '').trim().startsWith('摘要'));
    let authors = '', journal = '', year = '';
    const meta = divs.find(d => d.children.length === 2 && (d.innerText || '').includes('|'));
    if (meta) {
      authors = (meta.children[0].innerText || '').replace(/\s+/g, ' ').trim();
      const parts = (meta.children[1].innerText || '').split('|').map(x => x.trim());
      journal = parts[0] || '';
      year = parts[1] || '';
    }
    return {
      i: i + 1,
      type: tags[0] || '',
      title: titleEl ? titleEl.innerText.trim() : '',
      abs: absEl ? absEl.innerText.trim().replace(/^摘要[:：]\s*/, '') : '',
      authors: authors,
      journal: journal,
      year: year,
      keywords: tags.slice(1)
    };
  });
  return { ok: rows.length > 0, n: rows.length, rows: rows };
})()
'@
$data = Eval-Js -Tab $tab -Js $extract
if (-not $data.ok) { throw "提取失败" }

$refsPath = Join-Path $OutDir 'refs.json'
$data.rows | ConvertTo-Json -Depth 5 | Set-Content $refsPath -Encoding utf8
$lines = foreach ($row in $data.rows) { "$($row.i)|$($row.type)|$($row.title)|$($row.authors)|$($row.journal)|$($row.year)" }
$lines | Set-Content (Join-Path $OutDir 'refs_list.txt') -Encoding utf8
Write-Host "[done] $($data.rows.Count) 条 → $refsPath"
Write-Host "[tab] 抓取用标签 $tab，用完可关：sctl tabs close $tab"
