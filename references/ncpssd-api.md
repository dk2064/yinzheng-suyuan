# NCPSSD 取全文的三个接口（2026-11 实测有效）

站点 `https://www.ncpssd.cn`。登录后 cookie（`web_session` + `accessId`）即权限凭证，脚本用
`sctl cookies list --url https://www.ncpssd.cn/ -o json` 取。

## 为什么不能直接拼详情页

`https://www.ncpssd.cn/Literature/articleinfo?id=<id>` 直接访问返回 `monitor - 500`。
详情页 URL 还需要 `CTJCJS` 参数，它由站内检索跳转时生成，每次不同。所以要么走站内检索，
要么跳过详情页，直接用下面三个接口（推荐后者；这是省掉页面渲染，不涉及任何鉴权规避）。

## 1. 检索：`POST /searchHandler/search`

`Content-Type: application/x-www-form-urlencoded; charset=UTF-8`，表单字段：

| 字段 | 值 |
|---|---|
| `search` | **裸 DSL 文本**（不是 base64）：`(IKTE="题名" OR IKPYTE="题名"  OR IKST="题名" OR IKET="题名" OR IKSE="题名")` |
| `pageNum` | `1` |
| `pageSize` | `10` |
| `sort` | 空 |

`IKTE` = 题名，`IKPYTE` = 拼音题名，`IKST` / `IKET` / `IKSE` = 其它题名字段。

返回 `{result:true, code:200, data:{total, rows:[...], pages, currentPage}}`。每行关键字段：

```
id, type("中文期刊文章"/"学位论文"), title, creator("王文娟[1]"), cbw_name(刊名),
years, num(期), beginpage, endpage, gch(馆藏/路径码), class,
yearspdf("2017pdf"), volumn("2017JY1253"), ssnum("L1"), volurl("039"), numurl("012"),
fileName("7000461185.pdf"), pdfurl(旧 nssd.org 地址，已失效), subject, remark(摘要), issn
```

**坑**：返回 JSON 里同时有 `HtmlUrl` 和 `htmlUrl`，`ConvertFrom-Json` 默认直接报
「包含具有不同大小写的键」，必须加 `-AsHashtable`。

题名匹配建议：整题不中时用前 60% 题名重试（长题名常被分词断开）；再用去标点后的
「全等 > 包含 > 公共前缀长度」打分挑最优行，避免拿到同主题的另一篇。

## 2. 取下载地址：`POST /Literature/readDownloadUrl`

`Content-Type: application/json; charset=utf-8`，body：

```json
{"type":"中文期刊文章","articleid":"7000461185","id":"7000461185","periodname":"",
 "downcount":1,"readcount":-1,"gch":"97335A","class":"G448",
 "titleC":"对分课堂提升大学生幸福感研究——基于积极心理学PERMA模型视角",
 "showWriter":"王文娟[1]","pageType":1}
```

返回 `{"url":"https://ft.ncpssd.cn/pdf/getn/2017pdf/2017JY1253/L1/97335A/039/012/7000461185.pdf"}`。
这个 url 就是 `https://ft.ncpssd.cn/pdf/getn/{yearspdf}/{volumn}/{ssnum}/{gch}/{volurl}/{numurl}/{id}.pdf`，
可以直接从检索行拼。`code` 非 200 表示该篇无资源。

## 3. 下载签名：`POST /common/getMinioSign`

`Content-Type: application/json`，body `{}` → `{"result":true,"code":200,"data":"<sign>"}`。

拿到后在**下载请求头**里带：

```
sign: <sign>
site: npssd
dotype: down
```

`GET <url>` 即得 PDF 二进制（约 200KB–6MB）。缺 `sign` 头会被拒。签名一次性/短时效，逐篇重新取。

## 覆盖率实测（同一批 18 条推荐文献）

NCPSSD 命中 9 条；缺的集中在《心理月刊》《中国健康心理学杂志》《山东师范大学学报(自然科学版)》
《现代计算机》《淮北师范大学学报(自然科学版)》《福建电脑》《智库时代》《健康教育与健康促进》
与三篇学位论文（南方医科大学 / 吉首大学 / 西南交通大学）。学位论文只在知网/万方有正文。

## 其它渠道的实测结论

- 《现代计算机》官网给整期 PDF：`http://mpc.zsup.com.cn/downloads/{年}年第{卷}卷第{期}期总第{总期}期.pdf`
  （例：`.../2023年第29卷第03期总第771期.pdf`，12.4MB，`%PDF` 头正常）。
- Magtech 系期刊（如《心理月刊》`www.xlykzz.com`）DOI 落地页 `https://www.xlykzz.com/CN/<DOI>` 有
  `articleId`，但 `/CN/article/downloadArticleFile.do?attachType=PDF&id=<articleId>` 返回 `text/html`
  登录页，不是 PDF。
- `chinajournal.net.cn/kbDownload.aspx?fn=<CNKI号>` 返回 ASP.NET 登录表单。
- 百度学术可用来按题名反查 ncpssd / 万方 / 维普 / 官网链接。
