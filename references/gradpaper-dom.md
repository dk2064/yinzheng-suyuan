# GradPaper 表单与结果页 DOM（实测 2025-11 / 站点为 Arco Design + Vue）

入口：`https://www.gradpaper.com/dissertation`（带 `?msclkid=...` 的是广告来源参数，可省略）。
需登录；未登录时看不到「订单结果查询」。

## 第 1 步：主题表单

| 控件 | 定位 | 说明 |
|---|---|---|
| 文章标题 | `input[placeholder="请输入标题"]` | 智能选题下拉的可输入 input；Vue 受控，必须用原生 setter + `dispatchEvent('input')` |
| 研究思路 | `textarea` | 0/1500 字，可空 |
| 语言 | radio，默认「中文」 | 英文项要求标题与资料全用英文 |
| 学历 | `input[placeholder="请选择学历"]` | 选项：专科 / 本科 / 硕士 / 博士（另混有历史学校项） |
| 字数 | `input[placeholder="请选择字数"]` | 默认；学历选完后可点开 |
| 学科门类 | `.arco-select` 可见列表里第 4 个（无 placeholder） | 选「交叉学科」等 |
| 一级学科 | 第 5 个，placeholder 随门类变化（交叉学科下为「AI自动识别」） | 交叉学科门类下选项：AI自动识别 / 交叉学科类 |
| 专业名称 | `input[placeholder="请输入专业名称"]` | 自由输入，如「计算机科学与技术」 |
| 学校 | `input` 的 placeholder 含「学校」 | 搜索型下拉，输入校名前几个字（如「华东」）得到「华东师范大学-本科」 |
| 下一步 | `button` 文本「下一步」 | |

### Arco 交互铁律

- `sctl page click --selector <arco input>` → `TIMEOUT: the element is not visible`。**点 `.arco-select` 容器**。
- 选项节点 `.arco-select-option` 在 `body` 下会同时存在可见与隐藏副本，**必须 `filter(e => e.offsetParent)`**。
- 点开下拉后选项有 ~200ms 渲染延迟，轮询重试比固定 sleep 稳。
- 下拉打开状态是全局的：同时只能开一个，选完等 400–800ms 再开下一个。

## 第 2 步：推荐文献（点「下一步」之后）

- 列表项 `.arco-list-item`，**一次性 18 条**，无分页、无「加载更多」，内部滚动容器
  `div.overflow-y-auto.h-[calc(100%-50px)]`。
- 单条结构：
  - `.arco-tag` 序列：第 1 个是文献类型（`学术期刊` / `学术论文`），其余是关键词
  - `.font-bold`：标题
  - 以「摘要：」开头的 `div`：摘要
  - 元信息行：`children.length === 2` 且 innerText 含 `|`；`children[0]` 是作者，`children[1]` 是「期刊 | 年份」
- 页面按钮（18 条时）：解析校验文献格式、追加英文、搜索、上一步、下一步。

## 已知数据质量问题

- 第 3 条标题长期粘连另一篇（`《大学生心理健康素养与专业心理求助态度的相关性分析》` +
  `新冠肺炎疫情期间成人自评健康状况及影响因素分析`），跨会话可复现，属站点数据错误。
- 学位论文条目给的是学校名（如「南方医科大学」）而非期刊名。
- 同一题名重复跑，返回的 18 条完全一致 → 结果可缓存，不必反复消耗生成次数。
