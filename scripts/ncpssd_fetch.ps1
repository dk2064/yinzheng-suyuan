# 批量把推荐文献清单里能在 NCPSSD 找到的全文抓到本地。
# 输入 refs_list.txt：每行 `序号|类型|标题|作者|期刊|年份`（gradpaper_harvest.ps1 的产物）
param(
    [Parameter(Mandatory = $true)][string]$ListFile,
    [string]$OutDir = '.\fulltext',
    [int]$MaxItems = 60,
    [switch]$AcceptTerms,
    [switch]$Force
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
. "$PSScriptRoot\_notice.ps1"
Show-Notice '第三步 · 批量取全文'
Assert-Cap -Value $MaxItems -Max $script:MaxFetchItems -Name 'MaxItems'
Assert-Terms -Accepted:$AcceptTerms -What '批量取全文'
. "$PSScriptRoot\ncpssd_api.ps1"

$OutDir = Assert-SafeOutDir -Path $OutDir
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
$OutDir = (Resolve-Path $OutDir).Path
$ListFile = (Resolve-Path $ListFile).Path

$log = New-Object System.Collections.Generic.List[string]
$log.Add("序号`t状态`t字节`t类型`t实际标题`t期刊`t年`t页`tNCPSSD-ID`t备注")

$done = 0
foreach ($line in (Get-Content $ListFile -Encoding utf8)) {
    if ($line.Trim() -eq '') { continue }
    if ($done -ge $MaxItems) {
        Write-Host "已达单次上限 $MaxItems 篇，停止（见 USAGE-POLICY.md 第四节）。" -ForegroundColor Yellow
        break
    }
    $done++
    $p = $line.Split('|')
    if ($p.Count -lt 6) { continue }
    $no = Protect-Field ($p[0].Trim()) 20; $type = Protect-Field ($p[1].Trim()) 20; $title = ($p[2].Trim() -replace "[\t\r\n]", " ")
    $author = Protect-Field ($p[3].Trim()) 80; $journal = Protect-Field ($p[4].Trim()) 80; $year = Protect-Field ($p[5].Trim()) 10
    if ($no -notmatch "^\d{1,4}$") { continue }
    $outFile = Join-Path $OutDir ('{0:d2}.pdf' -f [int]$no)

    if ((Test-Path $outFile) -and -not $Force) {
        $log.Add("$no`t已有`t$((Get-Item $outFile).Length)`t$type`t$title`t$journal`t$year`t`t`t本地已存在")
        Write-Host "[$no] 已有，跳过"
        continue
    }

    try {
        $rows = Search-Ncpssd -Title $title
        $row = Select-NcpssdRow -Rows $rows -Title $title
        if (-not $row -and $title.Length -gt 24) {
            # 长题名常常只有前半段能命中
            $short = $title.Substring(0, [int]($title.Length * 0.6))
            $rows = Search-Ncpssd -Title $short
            $row = Select-NcpssdRow -Rows $rows -Title $short
        }
        if (-not $row) {
            $log.Add("$no`t未收录`t0`t$type`t`t$journal`t$year`t`t`tNCPSSD 无此刊/此文，需走期刊官网或知网/万方")
            Write-Host "[$no] NCPSSD 未收录：$title"
            continue
        }
        $url = Get-NcpssdPdfUrl -Row $row
        if (-not $url) {
            $log.Add("$no`t无资源`t0`t$($row['type'])`t$($row['title'])`t$($row['cbw_name'])`t$($row['years'])`t`t$($row['id'])`t接口拒发下载地址")
            Write-Host "[$no] 无资源：$($row['title'])"
            continue
        }
        $ok = Save-NcpssdPdf -Url $url -OutFile $outFile
        $bytes = if (Test-Path $outFile) { (Get-Item $outFile).Length } else { 0 }
        $pages = "$($row['beginpage'])-$($row['endpage'])"
        if ($ok) {
            $log.Add("$no`t成功`t$bytes`t$($row['type'])`t$($row['title'])`t$($row['cbw_name'])`t$($row['years'])`t$pages`t$($row['id'])`t")
            Write-Host "[$no] OK $bytes B  $($row['title'])"
        } else {
            $log.Add("$no`t失败`t0`t$($row['type'])`t$($row['title'])`t$($row['cbw_name'])`t$($row['years'])`t$pages`t$($row['id'])`tPDF 校验未通过")
            Write-Host "[$no] 下载失败：$($row['title'])"
        }
    } catch {
        $msg = Protect-Secret $_.Exception.Message
        $log.Add("$no`t异常`t0`t$type`t`t$journal`t$year`t`t`t$msg")
        Write-SafeHost "[$no] 异常：$msg" "DarkYellow"
    }
    Start-Sleep -Milliseconds 800
}

$logPath = Join-Path (Split-Path -Parent $ListFile) 'download_log.tsv'
$log | Set-Content $logPath -Encoding utf8
Clear-NcpssdSession
$okCount = ($log | Where-Object { $_ -match '成功|已有' }).Count
Write-Host "---- $okCount / $($log.Count - 1) 篇落盘；清单：$logPath ----"
