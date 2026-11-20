# 引证溯源 · yinzheng-suyuan

> 把论文里的参考文献**逐条追到源头**：取候选 → 核验真实性 → 取全文 → 给每份全文加一页「审核页」。
>
> 面向本科/硕士毕业论文的文献工作流。**仅供个人学习研究使用**，使用前请读 [USAGE-POLICY.md](USAGE-POLICY.md) 与 [DISCLAIMER.md](DISCLAIMER.md)。

---

## 它解决什么问题

写论文时最容易出事的三步，全靠手工：

1. **找不到够用的文献**——题目定了，不知道该引什么。
2. **不知道引的是真是假**——从各种「智能写作」站导出来的参考文献，题名、期次、页码张冠李戴，甚至整条是拼出来的。
3. **引了却没读过**——PDF 存了一堆，答辩一问就露。

这个技能把三步串成一条可复核的流水线，每步都留下可核对的痕迹。

## 四个阶段

| 阶段 | 做什么 | 产出 |
|---|---|---|
| **① 取候选** | A. 从 GradPaper 智能写作页按「题名+学历+学科专业+学校」取它推荐的文献（实测 18 条）<br>B. 直接在知网按题名/主题自检索，按相关性、年份、被引、类型、排除词筛选 | 同一 schema 的 `refs.json` + `refs_list.txt`（B 入口另出 `screening.tsv` 筛选记录） |
| **② 核验** | 逐条查是否真实存在、元数据是否串行（期次错、页码错、题名被拼接、作者数不符） | 核验结论：题名/作者/刊名/年卷期页/DOI/知网号；标注 `通过` / `需修正` / `存疑` |
| **③ 取全文** | 从国家哲学社会科学文献中心（NCPSSD）等开放渠道取 PDF，逐份校验是不是对的那篇文章 | `fulltext/NN.pdf` + `download_log.tsv` |
| **④ 加审核页** | 在每份 PDF 正文前插一页 A4 审核页 | `fulltext_audited/NN_<题名>_审核版.pdf` + `audit_index.tsv` |

**审核页长什么样**（7 块，可打印、可归档、超链接可点）：

1. 顶部色带 + 编号 `No.07 / 18`
2. 修正后的正式题名
3. 信息表：作者 / 文献类型 / 来源载体 / 年卷期页 / 收录标识
4. 核验结论：`✔ 通过` / `✔ 通过 · 需修正`（附具体修正说明）/ `⚠ 存疑`
5. 引用格式三连：**GB/T 7714-2015** / **APA 第 7 版** / 文中引用写法
6. 文章链接：DOI、知网检索、官网全文、NCPSSD
7. 文件与来源：文件名 / 正文页数 / 下载渠道 / SHA-256 指纹

## 快速开始

### 安装

```powershell
git clone https://github.com/dk2064/yinzheng-suyuan.git
cd yinzheng-suyuan
pwsh -NoProfile -File install.ps1            # 装到 %USERPROFILE%\.agents\skills
pwsh -NoProfile -File install.ps1 -WithDeps  # 同时补齐 python 依赖
```

也可以手动把整个目录复制到 `~/.agents/skills/yinzheng-suyuan`。

### 体检

```powershell
pwsh -NoProfile -File scripts\check_env.ps1      # 加 -Fix 自动 pip 装缺的包
```

十项全 OK 再往下走：PowerShell / Python / pypdf / python-docx / pypdfium2 / sctl daemon / 浏览器 daemon / docx→pdf（Word 或 LibreOffice）/ 中文字体 / 条款文件。

### 跑一遍

```powershell
# A. 取 GradPaper 推荐（先自检选择器，不消耗生成次数）
pwsh -File scripts\gradpaper_harvest.ps1 -Title "你的论文题目" -Degree 本科 `
  -Discipline 交叉学科 -SubDiscipline AI自动识别 -Major 计算机科学与技术 `
  -School "某某大学" -OutDir .\out -DryRun

# B. 知网自检索 + 筛选（按被引降序，只看期刊，2010 年后，被引≥2）
pwsh -File scripts\cnki_harvest.ps1 -Query "大学生心理健康 双因素模型" -Field 主题 `
  -Sort 被引 -MinCited 2 -Keep 期刊 -YearFrom 2010 -Top 20 -WithDetail -OutDir .\out

# 取全文
pwsh -File scripts\ncpssd_fetch.ps1 -List .\out\refs_list.txt -OutDir .\fulltext -AcceptTerms

# 加审核页（需要先按 examples\audit_meta.example.json 写好自己的 audit_meta.json）
pwsh -File scripts\add_audit_page.ps1 -Meta .\audit\audit_meta.json -SrcDir .\fulltext -Out .\audit\fulltext_audited
```

## 环境要求

| 项 | 要求 | 说明 |
|---|---|---|
| PowerShell | 7+ | 脚本用了 `-AsHashtable` 等 |
| Python | 3.10+ | 配 `pypdf`、`python-docx`；`pypdfium2` 可选（仅预览） |
| sctl | 0.3+ | 驱动你自己的浏览器；`install` 后 `sctl status` 能连上即可 |
| docx → PDF | Word 或 LibreOffice | 二选一，脚本会自动挑 |
| 字体 | 微软雅黑 `msyh.ttc` | 审核页排版用 |
| 账号 | 浏览器里登录对应站点 | 见下 |

**账号前提**：GradPaper（A 入口）、知网（要按被引筛时必须登录）、NCPSSD（取全文需要）。技能只使用**你自己浏览器里已存在的登录态**，不代管密码、不共享账号。

## 边界（写在代码里，不是口号）

`USAGE-POLICY.md` 是**强制条款**，脚本会执行其中的上限：

| 限制 | 值 |
|---|---|
| 知网单次抓取 | ≤ 200 条 |
| 知网单次保留 | ≤ 50 条 |
| 知网翻页 | ≤ 8 页 |
| 请求间隔 | ≥ 0.6 s，详情页 +1.2 s，单线程 |
| 批量取全文 | ≤ 60 篇，且必须显式 `-AcceptTerms` |
| 风控命中 | 检测到「验证码 / 安全验证 / 异常访问」立即中止，不做绕过 |

不做：绕付费墙、绕鉴权、伪造 cookie、绕验证码、并发爬取、代理池、囤积转售、二次分发正文。详见 [USAGE-POLICY.md](USAGE-POLICY.md)。

## 目录

```
yinzheng-suyuan/
├── SKILL.md                    # 技能定义（入口）
├── INSTALL.md                  # 换机器安装、故障排查
├── USAGE-POLICY.md             # 强制用途条款 + 内置上限
├── DISCLAIMER.md               # 免责声明
├── LICENSE                     # MIT（不覆盖文献元数据/正文）
├── requirements.txt
├── install.ps1
├── scripts/
│   ├── gradpaper_harvest.ps1   # A 入口：抓推荐文献
│   ├── cnki_harvest.ps1        # B 入口：知网检索 + 筛选
│   ├── ncpssd_api.ps1          # NCPSSD 三个接口的封装
│   ├── ncpssd_fetch.ps1        # 批量取全文
│   ├── audit_page.py           # 生成审核页 + 合并进 PDF
│   ├── add_audit_page.ps1      # 上面那步的编排（docx→pdf 双通道）
│   ├── check_env.ps1           # 依赖体检
│   └── _notice.ps1             # 条款横幅 + 上限断言
├── references/
│   ├── gradpaper-dom.md        # 表单控件与结果页 DOM 事实
│   └── ncpssd-api.md           # 检索/取址/签名三接口全参数
└── examples/
    └── audit_meta.example.json # audit_meta.json 的脱敏样例
```

## 已知限制

- **覆盖率不是 100%**：NCPSSD 未收录的刊（部分学报、行业刊、学位论文）取不到正文，只能给出知网/万方链接。
- **站点会变**：GradPaper 与知网都是前端频繁改版的站，选择器写在 `references/` 里，坏了按那里的事实改。
- **核验靠交叉比对**：脚本给的是「证据链 + 判定」，不是权威结论；关键引用仍要自己点开原文看一眼。
- **只支持中文文献场景**：GB/T 7714 与 APA 的拼接规则按中文期刊/学位论文写死。

## 许可

脚本与文档 MIT。**不覆盖**任何文献元数据与正文——那些版权属于原刊与原库。取回的 PDF 请自用，不要二次分发。

---

与 GradPaper、知网（CNKI）、国家哲学社会科学文献中心（NCPSSD）均无隶属关系，也未获其授权或背书。
