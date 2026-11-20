# 引证溯源 (yinzheng-suyuan) · 安装与依赖

给另一台电脑装这个技能时看这份。技能本身是几个 PowerShell 脚本 + 一个 Python 脚本，没有编译产物，
依赖只有「浏览器里已登录的账号」和「本机几个常规工具」。

## 1. 装到哪

DSH / 兼容 agent 的技能目录有两种放法，二选一：

| 位置 | 作用范围 |
|---|---|
| `%USERPROFILE%\.agents\skills\yinzheng-suyuan\` | 当前用户所有会话（本机就是 `C:\Users\<你>\.agents\skills\`） |
| `<你的工作区>\.agents\skills\yinzheng-suyuan\` | 只在那个工作区生效 |

解压后得到 `yinzheng-suyuan\` 文件夹，把**整个文件夹**放进上面任一 `skills\` 目录即可，
`SKILL.md` 必须在文件夹最外层（即 `...\skills\yinzheng-suyuan\SKILL.md`）。
根目录的 `install.ps1` 会自动做这件事并顺手体检：

```powershell
pwsh -NoProfile -File install.ps1                  # 装到 %USERPROFILE%\.agents\skills
pwsh -NoProfile -File install.ps1 -Target D:\ws\.agents\skills -Force
pwsh -NoProfile -File install.ps1 -WithDeps        # 缺 python 包时顺手 pip 装
```

装完重开会话，技能列表里应出现 `yinzheng-suyuan`。

## 2. 依赖清单

必需：

| 依赖 | 用途 | 检查 |
|---|---|---|
| PowerShell 7（`pwsh`） | 跑全部脚本 | `pwsh -v` |
| Python 3.10+ | 抓取结果解析、审核页生成、PDF 合并 | `python -V` |
| Python 包 `pypdf`、`python-docx` | 读页数/抽页/合并；生成审核页 | `python -c "import pypdf,docx"` |
| `sctl` CLI（ScriptCat 配套） | 驱动**你自己已登录**的浏览器填表、取 cookie | `sctl --version` |
| 一个运行中的浏览器（Edge/Chrome）+ ScriptCat 扩展 | sctl 的 daemon 要连上它 | `sctl status` |

docx→pdf 转换（第四步要用）满足其一即可：

| 方案 | 说明 |
|---|---|
| Microsoft Word（含 Office 16） | 默认走 COM `ExportAsFixedFormat`，最快最稳 |
| LibreOffice | 找不到 Word 时脚本自动改用 `soffice --headless --convert-to pdf` |
| 都不要 | 脚本会把 docx 留在 `_covers\`，你手动导出同名 PDF 后单跑 `audit_page.py merge` |

可选：`pypdfium2`（`pip install pypdfium2`）—— 只想把审核页渲染成 PNG 抽查时才需要。

## 2.5 使用条款（装完立刻读）

`USAGE-POLICY.md` 是强制条款：用途白名单 6 条、黑名单 10 条、合规对照、内置硬上限、责任划分。
不同意其中任一条，请删除本技能。

## 3. 账号前提

- **GradPaper**（A 入口）：浏览器里已登录 gradpaper.com（免费账号即可，推荐文献是免费返回的）。
- **知网 CNKI**（B 入口）：不登录也能看题录，但**被引数、下载数只有登录后才稳定显示**；
  想按被引筛选就登录一下（`kns.cnki.net`）。
- **NCPSSD**：浏览器里已登录 `www.ncpssd.cn`（第三步下载全文要偷它的 `web_session` cookie）。
  没登录也能跑，只是拿不到全文，只剩元数据与直达链接。

## 4. 装完自检

```powershell
# ① 依赖体检（逐项 OK/缺）
pwsh -NoProfile -File scripts\check_env.ps1

# ② A 入口干跑：只填表不点下一步，不消耗生成次数，验证选择器没变
pwsh -NoProfile -File scripts\gradpaper_harvest.ps1 `
  -Title "测试题名" -Degree 本科 -Discipline 交叉学科 `
  -SubDiscipline "AI自动识别" -Major 计算机科学与技术 `
  -School "某某大学" -OutDir .\out -DryRun

# ③ B 入口干跑：只检索+筛选，不落盘
pwsh -NoProfile -File scripts\cnki_harvest.ps1 `
  -Query "大学生心理健康 双因素模型" -Limit 20 -Top 5 -DryRun
```

A 入口六步全绿（标题/学历/学科门类/一级学科/专业名称/学校）说明 gradpaper 这一侧通了；
B 入口能打印出筛选表（含「保留/评分/被引/未通过原因」）说明知网这一侧通了。

## 5. 常见故障

| 现象 | 原因 / 处理 |
|---|---|
| `sctl: daemon not running` | 装并启动 ScriptCat 的 sctl 服务，浏览器保持开着；`sctl status` 应显示 online |
| `TIMEOUT: the element is not visible` | 点到了 Arco 的 `input`；只能点 `.arco-select` 容器和 `.arco-select-option`，且先过滤 `offsetParent` |
| 抓回 0 条（A 入口） | 站点改版或未登录；先 `-DryRun` 看卡在第几步，再对照 `references/gradpaper-dom.md` |
| 知网报「验证码 / 安全验证 / 异常访问」 | 抓太快被限流；隔几分钟再跑，或把 `-Limit` 调小；脚本遇到这类页会直接报错停下 |
| 知网筛选后 0 条 | 看 `screening.tsv` 的「未通过原因」列逐条对照；常见是 `-MinScore` 太严或没开 `-Sort 被引` 又设了 `-MinCited` |
| 知网被引/下载列全空 | 未登录知网，或没按被引排序；登录后加 `-Sort 被引` |
| NCPSSD 下载返回 HTML | cookie 过期（`web_session` 失效），浏览器重新登录一次 |
| `monitor - 500` | 不要手拼 `/Literature/articleinfo?id=` 链接，缺 `CTJCJS`；用 `ncpssd_api.ps1` 的接口 |
| 审核页第 2 页才是正文却没插进去 | `audit_meta.json` 里该条没有 `local.file`，或文件不在 `-SrcDir` 下 |
| 审核页文字被截断 | `notes` 太长溢到第 2 页，而 merge 只取第 1 页；缩短 `notes` 重跑 |
| `ConvertFrom-Json` 报「具有不同大小写的键」 | NCPSSD 返回 `HtmlUrl`/`htmlUrl` 并存，加 `-AsHashtable` |

## 6. 目录

```
yinzheng-suyuan\
├─ SKILL.md                     技能说明（四步流程 + 两个入口 + 筛选参数 + 坑）
├─ USAGE-POLICY.md              强制条款：用途白名单/黑名单、合规对照、内置上限、责任划分
├─ DISCLAIMER.md                免责声明：第三方站点、账号风险、准确性、再分发
├─ LICENSE                      MIT（仅覆盖脚本与文档）
├─ INSTALL.md                   本文件
├─ VERSION
├─ requirements.txt
├─ references\
│  ├─ gradpaper-dom.md          gradpaper 表单/结果页选择器实录
│  └─ ncpssd-api.md             NCPSSD 三个接口的参数与返回
├─ examples\
│  └─ audit_meta.example.json   审核页数据源样例（2 条，去掉一条改字段即可复制）
└─ scripts\
   ├─ _notice.ps1               运行横幅 + 内置上限 + -AcceptTerms 门槛（各脚本 dot-source）
   ├─ check_env.ps1             依赖体检
   ├─ gradpaper_harvest.ps1     入口 A：填表抓 GradPaper 推荐文献
   ├─ cnki_harvest.ps1          入口 B：知网自检索 + 筛选（翻页/排序/被引/年份/类型/排除词/合并去重）
   ├─ ncpssd_api.ps1            NCPSSD 接口函数库（被 fetch 调用）
   ├─ ncpssd_fetch.ps1          第三步：批量下全文（需 -AcceptTerms，≤60 篇）
   ├─ audit_page.py             第四步：build 审核页 / merge 到正文前
   └─ add_audit_page.ps1        第四步一键封装（build → 转 pdf → merge）
```
