# NCPSSD（国家哲学社会科学文献中心）纯 HTTP 取全文。
# 依赖：用户在浏览器里已登录 www.ncpssd.cn，本脚本用 sctl 偷 cookie 复用登录态。
# 不需要逐篇点页面，也不需要 CTJCJS 参数。
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Resolve-Sctl {
  $c = @(
    (Join-Path $env:LOCALAPPDATA 'sctl\bin\sctl.exe'),
    (Join-Path $env:USERPROFILE 'AppData\Local\sctl\bin\sctl.exe')
  ) | Where-Object { Test-Path $_ } | Select-Object -First 1
  if (-not $c) {
    $cmd = Get-Command sctl -ErrorAction SilentlyContinue
    if ($cmd) { $c = $cmd.Source }
  }
  return $c
}

$script:SctlExe = Resolve-Sctl
if (-not $script:SctlExe) {
  throw "找不到 sctl（sctl.exe）。请安装 ScriptCat 的 sctl CLI，或用 -SctlExe 指定路径。"
}
$script:Base = 'https://www.ncpssd.cn'
$script:UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0'
$script:MaxPdfBytes = 200MB   # 单份 PDF 体积上限，挡住被塞进来的大文件/非目标资源

# 只用正常 TLS：不降低校验等级，不跳过证书验证
try {
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
} catch {
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
}

# 统一发请求：异常信息先脱敏再抛，避免 cookie/token 进日志
function Invoke-NcpssdSafe {
  param(
    [Parameter(Mandatory = $true)][string]$Uri,
    [string]$Method = 'GET',
    [hashtable]$Headers,
    $Body,
    [string]$OutFile,
    [int]$TimeoutSec = 60
  )
  $arg = @{ Uri = $Uri; Method = $Method; Headers = $Headers; TimeoutSec = $TimeoutSec; UseBasicParsing = $true }
  if ($null -ne $Body) { $arg['Body'] = $Body }
  if ($OutFile) { $arg['OutFile'] = $OutFile }
  try {
    return Invoke-WebRequest @arg
  } catch {
    $msg = $_.Exception.Message
    if (Get-Command Protect-Secret -ErrorAction SilentlyContinue) { $msg = Protect-Secret $msg }
    throw "请求失败（$Method $([Uri]$Uri | ForEach-Object { $_.GetLeftPart('Authority') + $_.AbsolutePath })）：$msg"
  }
}

function Get-NcpssdCookieHeader {
    if ($script:CookieHeader) { return $script:CookieHeader }
    $json = & $script:SctlExe cookies list --url "$($script:Base)/" -o json 2>$null | ConvertFrom-Json
    if (-not $json.items -or $json.items.Count -eq 0) { throw 'NCPSSD cookie 为空：请先在浏览器登录 www.ncpssd.cn' }
    $skip = 'qimo_xstKeywords_', 'qimo_seosource_0', 'qimo_seokeywords_0', 'qimo_seosource_', 'qimo_seokeywords_', 'Hm_lvt_', 'Hm_lpvt_', 'HMACCOUNT', 'sensorsdata2015jssdkcross'
    $pairs = $json.items | Where-Object { $n = $_.name; -not ($skip | Where-Object { $n -like "$_*" }) } |
        ForEach-Object { "$($_.name)=$($_.value)" }
    $script:CookieHeader = ($pairs -join '; ')
    return $script:CookieHeader
}

# 会话清理：cookie 只活在本进程内存里，用完调一次
function Clear-NcpssdSession {
    $script:CookieHeader = $null
    [GC]::Collect()
}

function New-NcpssdHeaders {
    param([switch]$Json)
    $h = @{}
    $h['Cookie'] = Get-NcpssdCookieHeader
    $h['User-Agent'] = $script:UA
    $h['Accept'] = '*/*'
    $h['Accept-Language'] = 'zh-CN,zh;q=0.9'
    $h['Referer'] = "$($script:Base)/"
    $h['Origin'] = 'https://www.ncpssd.cn'
    if ($Json) { $h['Content-Type'] = 'application/json; charset=utf-8' }
    return $h
}

# 检索：返回 rows（含 id/gch/type/title/creator/cbw_name/years/num/beginpage/endpage/yearspdf/volumn/ssnum/volurl/numurl/fileName/class）
function Search-Ncpssd {
    param([string]$Title, [int]$PageSize = 10)
    $dsl = '(IKTE="{0}" OR IKPYTE="{0}"  OR IKST="{0}" OR IKET="{0}" OR IKSE="{0}")' -f $Title
    $h = New-NcpssdHeaders
    $h['Referer'] = "$($script:Base)/Literature/articlelist?sType=0"
    $h['X-Requested-With'] = 'XMLHttpRequest'
    $form = @{ search = $dsl; pageNum = 1; pageSize = $PageSize; sort = '' }
    $r = Invoke-NcpssdSafe -Uri "$($script:Base)/searchHandler/search" -Method POST -Headers $h -Body $form
    $j = $r.Content | ConvertFrom-Json -AsHashtable    # 必须 -AsHashtable：返回里有 HtmlUrl/htmlUrl 大小写冲突
    if ($j.data -and $j.data.rows) { return $j.data.rows }
    return @()
}

# 从 readDownloadUrl 拿到受签名保护的真实 PDF 地址
function Get-NcpssdPdfUrl {
    param($Row)
    $h = New-NcpssdHeaders -Json
    $obj = [ordered]@{
        type       = $Row['type']
        articleid  = $Row['id']
        id         = $Row['id']
        periodname = ''
        downcount  = 1
        readcount  = -1
        gch        = $Row['gch']
        class      = $(if ($Row['class']) { $Row['class'] } else { '' })
        titleC     = $Row['title']
        showWriter = $Row['creator']
        pageType   = 1
    }
    $body = $obj | ConvertTo-Json -Compress
    $r = Invoke-NcpssdSafe -Uri "$($script:Base)/Literature/readDownloadUrl" -Method POST -Headers $h -Body ([Text.Encoding]::UTF8.GetBytes($body))
    $j = $r.Content | ConvertFrom-Json -AsHashtable
    if ($j['url']) { return $j['url'] }
    return $null
}

function Get-NcpssdMinioSign {
    $h = New-NcpssdHeaders -Json
    $r = Invoke-NcpssdSafe -Uri "$($script:Base)/common/getMinioSign" -Method POST -Headers $h -Body ([Text.Encoding]::UTF8.GetBytes('{}'))
    $j = $r.Content | ConvertFrom-Json -AsHashtable
    return $j['data']
}

# 下载并校验 %PDF 头；失败删除半成品
function Save-NcpssdPdf {
    param([string]$Url, [string]$OutFile)
    $h = @{}
    $h['sign'] = Get-NcpssdMinioSign
    $h['site'] = 'npssd'
    $h['dotype'] = 'down'
    $h['User-Agent'] = $script:UA
    $h['Accept'] = '*/*'
    $h['Referer'] = "$($script:Base)/"
    $dir = Split-Path -Parent $OutFile
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    try { Invoke-NcpssdSafe -Uri $Url -Headers $h -OutFile $OutFile -TimeoutSec 180 | Out-Null }
    catch { if (Get-Command Write-SafeHost -ErrorAction SilentlyContinue) { Write-SafeHost "  下载异常：$($_.Exception.Message)" 'DarkYellow' }; return $false }
    # 体积上限：挡住被塞进来的大文件/非目标资源
    $len = (Get-Item $OutFile -ErrorAction SilentlyContinue).Length
    if ($len -gt $script:MaxPdfBytes) {
        Remove-Item $OutFile -Force -ErrorAction SilentlyContinue
        return $false
    }
    $b = [IO.File]::ReadAllBytes($OutFile)
    if ($b.Length -lt 10000 -or -not [Text.Encoding]::ASCII.GetString($b[0..4]).StartsWith('%PDF')) {
        Remove-Item $OutFile -Force -ErrorAction SilentlyContinue
        return $false
    }
    return $true
}

# 题名相似度选最优命中（去标点 + 前缀比 + 包含关系）
function Get-NcpssdNormTitle {
    param([string]$s)
    if (-not $s) { return '' }
    return ($s -replace '[\s\u3000]', '' -replace '[—\-–－_:：,，.。;；!！?？""''''()（）《》\[\]【】…·]', '')
}

function Select-NcpssdRow {
    param($Rows, [string]$Title)
    if (-not $Rows -or $Rows.Count -eq 0) { return $null }
    $t = Get-NcpssdNormTitle $Title
    $best = $null; $bestScore = -1
    foreach ($row in $Rows) {
        $rt = Get-NcpssdNormTitle $row['title']
        if (-not $rt) { continue }
        $score = 0
        if ($rt -eq $t) { $score = 1000 }
        elseif ($rt.Contains($t) -or $t.Contains($rt)) { $score = 500 + [Math]::Min($rt.Length, $t.Length) }
        else {
            $n = [Math]::Min($rt.Length, $t.Length); $p = 0
            while ($p -lt $n -and $rt[$p] -eq $t[$p]) { $p++ }
            $score = $p
        }
        if ($score -gt $bestScore) { $bestScore = $score; $best = $row }
    }
    return $best
}
