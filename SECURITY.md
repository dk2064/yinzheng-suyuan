# 安全说明 · Security Policy

本技能会读取你浏览器里的登录态、向第三方站点发请求、把 PDF 写到本地。下面是它的安全边界、已做的加固，以及漏洞怎么报。

---

## 一、它碰什么，不碰什么

| 项目 | 行为 |
|---|---|
| 浏览器数据 | 只通过 `sctl cookies list --url <site>` 读**目标站点**的 cookie；不读密码库、不读历史记录、不读其他站点 |
| 凭证存储 | cookie 只存在进程内存（`$script:CookieHeader`），**不落盘、不写日志、不写临时文件**；进程结束即消失 |
| 网络请求 | 只访问 `www.ncpssd.cn` / `ft.ncpssd.cn` / `kns.cnki.net` / `www.gradpaper.com`（以及你自己传进来的期刊官网 URL） |
| 上传 | 无。没有任何遥测、崩溃上报、版本统计 |
| 文件写入 | 只写你指定的输出目录；拒绝写入 `%SystemRoot%`、`%ProgramFiles%`、`System32` 等系统路径 |
| 子进程 | 只调用 `sctl`（浏览器桥）与 `WINWORD.EXE`（docx→pdf）；不执行 shell 字符串插值 |
| 内容信任 | 抓回来的题名/摘要/期刊名一律当**不可信数据**：写进表格前净化控制字符与路径分隔符，写文件名用序号不用户文本 |

---

## 二、已内置的加固

- **TLS**：强制 TLS 1.2/1.3，任何情况下不关闭证书校验（没有 `-SkipCertificateCheck`、没有 `ServerCertificateValidationCallback` 例外）。
- **异常脱敏**：所有异常先过 `Protect-Secret` 再打印/落盘，`web_session=…`、`Cookie: …`、`ghp_…`、JWT 会被替换成 `***`。
- **体积上限**：单份 PDF 200 MB；超过即删并标记失败，避免被重定向到超大/无关资源。
- **量级上限**：知网单次 ≤200 条 / 保留 ≤50 条 / 翻页 ≤8 页；批量下载 ≤60 篇；请求间隔 ≥600 ms。硬编码，无关闭参数。
- **路径护栏**：`audit_meta.json` 里 `local.file` 必须落在 `--src-dir` 内（`realpath` 比对，拒绝 `../`）；`extract` 必须是正序正整数页码。
- **输出目录护栏**：`Assert-SafeOutDir` 拒绝系统目录。
- **条款门槛**：批量取全文需显式 `-AcceptTerms`，否则 `exit 2`。
- **发布前自检**：`scripts/check_secrets.ps1` 扫描仓库/包内是否混进 cookie、token、私钥、个人路径。

---

## 三、使用者的责任（这几条不加固就出事）

1. **只用自己的账号**。cookie 代表你的登录身份；把技能目录连同 `sctl` 会话一起交给别人等于交出账号。
2. **别在共享机器上长时间挂着登录态**。用完执行 `Clear-NcpssdSession`（脚本结尾已自动调用）或直接关浏览器。
3. **不要把你的 `refs.json` / `download_log.tsv` / 审核版 PDF 提交进任何公开仓库**——它们含你真实的检索历史与文献正文。
4. **不要把 cookie 写进命令行参数**。本技能从不接受 `-Cookie` 这类入参，请保持这样。
5. **输出目录别指向网盘同步目录**，避免正文被自动同步到云端再被二次分发。

---

## 四、已知风险与不修的边界

| 风险 | 说明 |
|---|---|
| 第三方站点改版 | 接口/选择器失效属于**功能问题**，不是安全问题；脚本会明确报错而不是静默产出错数据 |
| 站点侧风控 | 触发验证码/异常访问时脚本**立即停止**，不做绕过——这是设计选择，不是缺陷 |
| Word COM 导出 | 依赖本机 Word；畸形 docx 理论上可触发 Word 解析器问题。审核页 docx 由本技能自己生成，不接受外部 docx 输入 |
| 依赖供应链 | `pypdf` / `python-docx` / `pypdfium2` 从 PyPI 安装，请自行核对来源；`check_env.ps1 -Fix` 只跑 `pip install`，不执行远程脚本 |

---

## 五、报告漏洞

发现以下任一类问题，请**不要开 public issue**，直接私信或开 GitHub Security Advisory：

- 能把 cookie / token / 本机路径带进日志、异常消息、输出文件
- 能通过 `refs.json` / `audit_meta.json` 让脚本读写 `src_dir` 之外的路径
- 能绕过内置上限（条数、页数、体积、间隔）
- 命令注入、路径穿越、任意文件写入/删除

报告时请给：版本号（`VERSION`）、复现命令、期望 vs 实际、`check_env.ps1` 输出。**请勿附带真实 cookie 或未脱敏的文献全文**。

响应：确认后 7 天内给结论；修复版在 `VERSION` 递增并写入 CHANGELOG 段落。

---

## 六、支持版本

只维护最新 `VERSION`。旧版本请升级后再报。
