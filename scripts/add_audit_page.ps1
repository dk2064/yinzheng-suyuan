# add_audit_page.ps1 — 给已下载的文献 PDF 逐篇加一页「审核页」（题名/出处/核验结论/引用串/文章链接/文件指纹）
#
# 依赖：python（python-docx、pypdf）、Microsoft Word（docx→pdf，COM 导出）
#
# 示例：
#   pwsh -NoProfile -File add_audit_page.ps1 -Meta .\audit\audit_meta.json -SrcDir .\fulltext -Out .\fulltext_audited

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Meta,
    [string]$SrcDir = ".",
    [string]$Out = "fulltext_audited",
    [string]$DocxDir = "",
    [string]$Python = "python",
    [switch]$RemoveCovers
)

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here "_notice.ps1")
Show-Notice '第四步 · 加审核页'
$py = Join-Path $here "audit_page.py"

$Meta = (Resolve-Path $Meta).Path
$SrcDir = (Resolve-Path $SrcDir).Path
if ([string]::IsNullOrWhiteSpace($DocxDir)) {
    $DocxDir = Join-Path (Split-Path -Parent $Meta) "_covers"
}
New-Item -ItemType Directory -Force -Path $DocxDir | Out-Null

Write-Host "== 1/3 生成审核页 docx ==" -ForegroundColor Cyan
& $Python $py build --meta $Meta --src-dir $SrcDir --docx-dir $DocxDir
if ($LASTEXITCODE -ne 0) { throw "audit_page.py build 失败（exit $LASTEXITCODE）" }

$docx = @(Get-ChildItem $DocxDir -Filter *.docx)
if ($docx.Count -eq 0) { throw "没有生成任何 docx：检查 audit_meta.json 里 local.file 是否存在" }

Write-Host "== 2/3 docx -> pdf（$($docx.Count) 份）==" -ForegroundColor Cyan
$done = 0

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try {
        foreach ($d in $docx) {
            $pdfPath = [System.IO.Path]::ChangeExtension($d.FullName, ".pdf")
            $doc = $word.Documents.Open($d.FullName, $false, $true)
            try { $doc.ExportAsFixedFormat($pdfPath, 17) } finally { $doc.Close(0) }
            Write-Host "  $($d.Name) -> $(Split-Path $pdfPath -Leaf)"
            $done++
        }
    } finally {
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        [System.GC]::Collect()
    }
    Write-Host "  导出器：Microsoft Word" -ForegroundColor DarkGray
} catch {
    Write-Host "  Word 不可用（$($_.Exception.Message.Trim())），改用 LibreOffice" -ForegroundColor Yellow
}

if ($done -eq 0) {
    $soffice = @(
        "C:\Program Files\LibreOffice\program\soffice.exe",
        "C:\Program Files (x86)\LibreOffice\program\soffice.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $soffice) { $soffice = (Get-Command soffice -ErrorAction SilentlyContinue).Source }
    if (-not $soffice) {
        throw "docx 转 PDF 失败：本机既没有 Microsoft Word，也没有 LibreOffice。装其一后重跑，或把 $DocxDir 里的 docx 手动导成同名 PDF 再单跑 merge。"
    }
    Write-Host "  改用 LibreOffice：$soffice" -ForegroundColor DarkGray
    try {
        & $soffice --headless --norestore --convert-to pdf --outdir $DocxDir @($docx.FullName)
        $done = @(Get-ChildItem $DocxDir -Filter *.pdf).Count
    } catch {
        Write-Host "  LibreOffice 也失败：检查 $DocxDir 是否已有同名 PDF，或手动导出后单跑 merge。" -ForegroundColor Red
        throw
    }
}

Write-Host "== 3/3 合并审核页 + 正文 ==" -ForegroundColor Cyan
& $Python $py merge --meta $Meta --src-dir $SrcDir --cover-dir $DocxDir --out $Out
if ($LASTEXITCODE -ne 0) { throw "audit_page.py merge 失败（exit $LASTEXITCODE）" }

if ($RemoveCovers) {
    Remove-Item $DocxDir -Recurse -Force -ErrorAction SilentlyContinue
} else {
    Write-Host "审核页 docx/pdf 保留在：$DocxDir（可单独改字后再跑 merge）" -ForegroundColor DarkGray
}
Write-Host "完成：$((Resolve-Path $Out).Path)" -ForegroundColor Green
