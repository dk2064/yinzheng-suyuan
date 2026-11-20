---
name: yinzheng-suyuan
display_name: 引证溯源
description: "引证溯源（yinzheng-suyuan）：本地整理参考文献并逐条追溯到源头、核验真实性、归档全文，仅供个人学习研究使用。两个取候选入口——①从 GradPaper（gradpaper.com/dissertation 智能写作）按题名+学历+学科专业+学校取它推荐的 18 条；②直接在知网（kns.cnki.net）按题名/主题自检索，按相关性、年份、被引、文献类型、排除词筛选。两条来源产出同格式 refs.json，之后做真实性核验、从开放渠道取全文、并在每份全文正文前插入一页「审核页」（题名/出处/核验结论/GB-T 7714 与 APA 引用串/文章链接/文件指纹）后存到本地。当用户说 gradpaper 推荐文献、抓推荐引文、按题目取参考文献、从 gradpaper 导文献、去知网搜参考文献、知网检索筛选文献、审核参考文献是否真实、核对引用出处、下载这批文献全文、给文献加审核页/引用页/封面页、参考文献溯源 时使用。也用于只跑其中一段：仅抓取（-DryRun）、仅知网筛选、仅核验、仅下载全文、仅加审核页。使用前必读包内 DISCLAIMER.md 与 USAGE-POLICY.md；脚本不含也不接受绕过鉴权、破解、验证码绕过、并发爬取等用途。"
license: MIT
version: 3.0.0
notice: "仅限个人学习研究；不提供文献正文；与 gradpaper.com / kns.cnki.net / ncpssd.cn 无隶属关系；使用前读 USAGE-POLICY.md"
---

# 参考文献：取候选 → 核验 → 全文 → 审核页

两条取候选的路子，产出**同一个 schema**，后面的核验/下载/审核页完全共用：

- **A. GradPaper 推荐**：`gradpaper.com/dissertation` 智能写作在填完题目/学历/学科专业/学校并点「下一步」后，免费返回一批参考文献（实测 18 条）。优点是体裁贴合"本科毕业论文"，缺点是只有 18 条、不能按条件筛。
- **B. 知网自检索**：直接在 `kns.cnki.net` 按题名/主题/关键词搜，能控制年限、被引、文献类型，条数自己定。优点是量大可控，缺点是要自己筛掉不相关的。

本技能做四件事：取候选并结构化；逐条核验是否真实存在、元数据是否串行；从国家哲学社会科学文献中心
（NCPSSD）和其它开放渠道把 PDF 抓下来；最后给每份 PDF 正文前插一页**审核页**，形成可直接归档的「带证据的全文」。

> **许可与免责**：脚本以 MIT 许可发布；仅限个人学习研究使用，使用前请读 `DISCLAIMER.md`
> （第三方站点协议、账号风险、结果准确性、再分发限制）。

## 环境与安装

装到新机器时先看 `INSTALL.md`，然后跑一次体检：

```powershell
pwsh -NoProfile -File scripts\check_env.ps1        # 加 -Fix 会自动 pip 装缺的包
```

必需：PowerShell 7、Python 3.10+（`pypdf`、`python-docx`）、`sctl` CLI（ScriptCat 配套，daemon 连上浏览器）、
浏览器里已登录 gradpaper.com（走 A 入口）与 www.ncpssd.cn（要下全文时）。
知网（走 B 入口）不登录也能看题录，但**被引/下载数只有登录后才稳定显示**。
docx→pdf 用 Microsoft Word（默认）或 LibreOffice（Word 缺失时脚本自动改用它；两者都没有就把 `_covers\` 里的
docx 手动导成同名 PDF，再单跑 `audit_page.py merge`）。

## 交付物

| 文件 | 内容 |
|---|---|
| `refs.json` | 结构化记录：i/type/title/abs/authors/journal/year/keywords（+source/cited/doi/url） |
| `refs_list.txt` | 制表用竖线清单 `序号\|类型\|标题\|作者\|期刊\|年份` |
| `screening.tsv` | **仅 B 入口**：逐条保留/评分/被引/下载/类型/年份/未通过原因/链接，人工复核用 |
| `verify.tsv` | 逐条真实性核验：exists/结论/修正标题/卷期页/DOI/知网号/直达链接 |
| `download_log.tsv` | 逐条下载结果：状态/字节/实际匹配标题/期刊/年/页/NCPSSD-ID |
| `fulltext\NN.pdf` | 全文原件，NN = 序号补零 |
| `audit\audit_meta.json` | 审核页数据源：逐条题名修正/作者/刊年卷期页/核验结论/引用串/链接/文件指纹 |
| `audit\fulltext_audited\NN_<题名>_审核版.pdf` | **最终交付**：审核页 1 页 + 正文，题名与链接都在首页，可直接发给导师 |
| `audit\fulltext_audited\audit_index.tsv` | 审核版清单：序号/题名/核验/页数/文件名/源文件 SHA-256 前 16 位 |

## 第一步：取候选（两个入口，产出同一 schema）

### A. GradPaper 推荐（浏览器，需已登录 gradpaper）

```powershell
pwsh -NoProfile -File scripts\gradpaper_harvest.ps1 `
  -Title "基于SAM-MH模型对大学生心理研究" -Degree 本科 `
  -Discipline 交叉学科 -SubDiscipline "AI自动识别" `
  -Major 计算机科学与技术 -School "某某大学" `
  -OutDir .\out
```

要点：

- 走 `sctl` 驱动用户自己已登录的浏览器，复用其 cookie；脚本自己开 background 标签，不动用户已开的页面。
- 表单是 Arco Design：三层联动（学科门类 → 一级学科/AI自动识别 → 专业名称自由输入），学校是搜索型下拉。
  **只能点 `.arco-select` 容器和 `.arco-select-option`，不能点 input** —— `sctl page click` 对 Arco 的 input 会报
  TIMEOUT「the element is not visible」。
- 下拉里的选项在 `body` 下有隐藏副本，选之前必须过滤 `e.offsetParent`。
- 点「下一步」后要等 `.arco-list-item` 出现；18 条一次性渲染，无分页。
- `-DryRun` 只填表不点下一步：不会消耗生成次数，用于确认选择器没变。
- 生成结果记录在站点订单里（保留 60 天），重复跑同一题名会得到同一批文献；第 3 条的标题在两个不同会话里都
  出现「正题 + 另一篇标题」粘连，属于站点数据错误，核验阶段必须修正（见下）。

细节选择器与实测取值见 `references/gradpaper-dom.md`。

### B. 知网自检索（可筛可控，条数自己定）

```powershell
# 按整条题名找同类研究（篇名检索 + 相关性筛选）
pwsh -NoProfile -File scripts\cnki_harvest.ps1 `
  -Query "基于SAM-MH模型对大学生心理研究" -Field 篇名 -Limit 40 -Top 12 `
  -YearFrom 2015 -MinScore 0.45 -Drop 高职,中职,护理 -OutDir .\out

# 按主题面找高被引奠基文献（被引排序 + 只要期刊 + 补摘要关键词）
pwsh -NoProfile -File scripts\cnki_harvest.ps1 `
  -Query "大学生心理健康 双因素模型" -Field 主题 -Limit 60 -Top 10 `
  -Sort 被引 -MinCited 5 -Keep 期刊 -YearFrom 2010 -WithDetail -OutDir .\out
```

参数就是筛选器，全部命中才 `keep=Y`，落 `screening.tsv` 供复核：

| 参数 | 作用 | 备注 |
|---|---|---|
| `-Field` | 主题 SU / 篇名 TI / 关键词 KY / 摘要 AB / 全文 FT / 作者 AU | 默认主题 |
| `-Limit` | 抓多少条（自动翻页，最多 8 页×20） | 抓得多才筛得准，一般 40–60 |
| `-Top` | 最终留几条 | 进入核验/下载的条数 |
| `-YearFrom/-YearTo` | 年份窗口 | 0 = 不限 |
| `-MinCited` | 被引下限 | 想用它必须配 `-Sort 被引`，否则列表里的被引数常为空 |
| `-Keep` | 全部/期刊/学位 | 学位论文在知网常拿不到全文，只要文献综述就选期刊 |
| `-Sort` | 相关度/被引/时间/下载 | 脚本会显式点一次排序，避免站点默认值漂移导致结果不可复现 |
| `-MinScore` | 相关性阈值 | 多词查询＝题名命中词比例；单条长题名＝去掉「基于/研究/分析」等功能词后的字符 bigram 覆盖率 |
| `-Drop` | 排除词，逗号分隔 | 例：`高职,中职,中专,护理,思政,中学生` |
| `-WithDetail` | 逐条开详情页补摘要/关键词/DOI | 每条 +2~3 秒，但后面核验与审核页都要用 |
| `-Merge` | 合并已有 `refs.json` 并去重 | 例：`-Merge .\gpd\refs.json` 把 GradPaper 的 18 条并进来 |

判定与产物：

- 相关性用可解释的规则算，`screening.tsv` 里写明每条为什么被砍（`相关性<0.45;被引<5;非期刊`），不是黑箱。
- 输出与 A 入口完全同构，**后续 `ncpssd_fetch.ps1`、`add_audit_page.ps1` 不用改**。
- 知网结果里的 `url` 是可回溯的详情页直链（`kcms2/article/abstract?v=...`），审核页「文章链接」块优先用它。
- `-DryRun` 只打印前 15 条筛选结果，不落盘。

## 第二步：核验（每篇都要过）

抓回来不等于真的。核验至少确认四件事：

1. **题名真实存在**，且与站点给的题名一致（注意上面的粘连错误）。
2. **元数据不串行**：作者、期刊、年、卷(期)、页码要对得上；GradPaper 偶尔把 A 文的年/刊挂到 B 文。
3. **补全标识**：CNKI 文件名式编号（如 `HZHX201712022`）、DOI、万方 id —— 后续取全文全靠它们。
4. **拿到直达链接**：知网/万方/维普/NCPSSD/期刊官网任一可定位页面。

做法：并行派 subagent，每个负责 3–5 条，返回固定字段（exists / confidence / 修正标题 / 卷期页 / DOI / 链接）。
超过 8 条时分批，一批不超过 8 个并发；要求返回结构化 JSON，schema 用 `type` + `enum` 明确写。

## 第三步：全文下载

首选 NCPSSD（国家哲学社会科学文献中心）：中文期刊覆盖面好、支持登录后整篇下载，且能用纯 HTTP 脚本跑，
不用逐篇点页面。**必须先让用户在浏览器里登录 ncpssd.cn**，脚本从浏览器偷 cookie。

```powershell
pwsh -NoProfile -File scripts\ncpssd_fetch.ps1 -ListFile .\out\refs_list.txt -OutDir .\out\fulltext
```

NCPSSD 没有的，按顺序试：

1. **期刊官网**：部分刊（如《现代计算机》`mpc.zsup.com.cn/downloads/<年>年第<卷>卷第<期>期总第<总期>期.pdf`）直接
   提供整期 PDF —— 抓到整期后按页码取当篇。
2. **DOI 落地页**：`https://<刊官网>/CN/<DOI>` 这类 Magtech 站点常有 `/CN/article/downloadArticleFile.do?attachType=PDF&id=<articleId>`，但多半要订阅，返回 HTML 登录页即视为失败。
3. **学位论文**：#16–#18 这类只有知网/万方有全文，无订阅账号时只能给直达链接，不要伪造。
4. 万方/维普/知网正文一律需要机构权限；`chinajournal.net.cn/kbDownload.aspx?fn=<CNKI号>` 返回 ASP.NET 登录表单，不是直链。

实测结论：同一批 18 条里，NCPSSD 能拿到 9 条，加期刊官网整期 PDF 共 10 条；其余 8 条只有元数据与直达链接。
知网入口抓来的清单同样吃这套：一次 6 条实测 4 条成功（《心理科学》《西南大学学报》《中国卫生事业管理》《高教论坛》），
未收录的是《济南职业学院学报》《循证护理》。
把这个覆盖率如实写进交付说明，不要含糊成「全部下载完成」。

## 第四步：加审核页（最终交付形态）

用户要的不是裸 PDF，而是「打开第一页就能看见这篇是什么、从哪来、可不可信、该怎么引」的归档件。

```powershell
pwsh -NoProfile -File scripts\add_audit_page.ps1 `
  -Meta .\audit\audit_meta.json -SrcDir .\fulltext -Out .\audit\fulltext_audited
```

步骤：`audit_page.py build` 生成审核页 docx → Word COM 导出 PDF（`ExportAsFixedFormat(...,17)`）→
`audit_page.py merge` 用 pypdf 把审核页放到正文第一页前 → 落盘 `NN_<题名>_审核版.pdf` + `audit_index.tsv`。

审核页上固定有这些块（缺字段自动显示「—」，不会报错）：

1. 顶部色带：`文献审核页 · LITERATURE VERIFICATION SHEET` + 右对齐 `No.NN / 总数`。
2. 题名（大字，用核验后的**修正题名**，不是站点原题）。
3. 信息表：作者 / 文献类型 / 来源载体 / 年卷期页 / 收录标识（DOI、知网号、万方号、文章编号、ISSN）。
4. 核验结论：`✔ 通过` / `✔ 通过 · 需修正` / `未能核实`，配色区分，后跟一句判定口径；下一行写具体修正说明。
5. 引用格式：`GB/T 7714-2015`、`APA（第 7 版）`、`文中引用`（作者 >3 自动「前 3 + 等」；2 人用「和」）。
6. 文章链接：可点击超链接，逐条列出（知网检索 / DOI / 万方 / 期刊官网 / 整期 PDF）。
7. 文件与来源：文件名 / 正文页数 / 下载渠道 / SHA-256 前 32 位；页脚写审核时间、批次来源、生成工具。

`audit_meta.json` 的结构（`batch` + `items[]`）：`no`、`kind`（学术期刊|学位论文）、`title`（修正后）、
`title_raw`（站点原题，可选）、`authors[]`、`venue`、`city`（学位论文）、`year`、`volume`、`issue`、`pages`、
`ids{doi,cnki,wanfang,cnki_cdmd,article_no,issn}`、`verdict`（confirmed|corrected|unverified）、`notes`、
`links[{label,url}]`、`local{file,extract,source,ncpssd_id,note}`。

`local.extract: [起页, 止页]` 用于从整期 PDF 里抽当篇（页码是**该 PDF 文件内的物理页序**，不是印刷页码；
先在整期 PDF 里搜题名定位）。没有 `local.file` 的条目不出审核版，只在 `audit_index.tsv` 里标「无正文」。

审核页宽度按 A4、`微软雅黑`、上下边距 1.6/1.4cm 调过，正文内容再多也是一页；若某条 `notes` 特别长导致溢出，
merge 只取审核页第 1 页，会截掉尾部 —— 先缩短 `notes` 再重跑。

## 坑

- **NCPSSD 详情页 URL 不能拼**：`/Literature/articleinfo?id=<id>` 直接访问返回 `monitor - 500`，因为还差站点内跳转
  生成的 `CTJCJS` 参数。要么走站内检索，要么用 `ncpssd_api.ps1` 的 API（它不需要 CTJCJS）。
- **NCPSSD 首页检索框 placeholder 是「请输入检索词」**，且首页有 200+ 个隐藏 input（高级检索面板），必须过滤可见性。
- **`sctl` 标签会整批消失**：长循环里每轮都要校验 `--tab` 还在，报 `no tab` 就重开。
- **PowerShell 的 `.NET` 调用不吃 `cd`**：`[IO.File]::ReadAllBytes('相对路径')` 按进程 CWD（会话工作区）解析，
  一律用绝对路径。
- `ConvertFrom-Json` 遇到大小写不同的同名键（NCPSSD 返回里 `HtmlUrl`/`htmlUrl`）会直接报错，加 `-AsHashtable`。
- **知网列表的「被引」列在非被引排序时常为空**：想按被引筛就先 `-Sort 被引`，否则 `-MinCited` 会把好文献全砍掉。
- **知网默认排序不是相关度**（实测会停在「发表时间 DESC」），所以脚本每次显式点一次排序，保证同 query 可复现。
- **知网长题名检索要放宽匹配**：整条题名当查询时，知网会做模糊匹配返回一堆同前缀文献；用 `-Field 篇名` + bigram
  相似度（已内置去「基于/研究/分析」等功能词），阈值 `-MinScore` 一般 0.45–0.55；单篇精确核对请改用 `-Field 篇名`
  且把 `-MinScore` 提到 0.7。
- **知网详情页标题 h1 不可靠**（有时是「自动登录」弹层的 h1）：题名一律取自列表的 `td.name a`，详情页只取
  `#ChDivSummary`（摘要）、`.keywords a`（关键词，尾部带 `;` 要清）和正文里的 `DOI:`。
- 知网 `td.quote`/`td.download` 里的数字可能带单位，取值前用 `[^\d]` 清洗再 `TryParse`。
