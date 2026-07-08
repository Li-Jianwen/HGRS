---
description: Use this skill when 用户提到人遗系统、HGR、HGRS、人类遗传资源、项目公示、审批结果爬取、抓取审批数据、更新HGR数据。全自动抓取公示PDF，提取审批表格，按批次独立保存+汇总到总表，输出运行简报。也支持从科技部政务服务平台(fuwu.most.gov.cn)补充2021~2023年历史审批结果。
name: HGRS
---

# HGRS — 人类遗传资源行政许可公示自动爬取

全自动从 [中国人类遗传资源服务系统](https://apply.hgrg.net/login) 抓取项目公示 PDF/DOCX，提取审批表格数据，按批次独立保存 + 汇总到总表，每次运行输出简报。  
另支持从 [科技部政务服务平台](https://fuwu.most.gov.cn/html/jgcx/index.html) 补充 2021~2023 年历史审批结果数据（HTML 表格直接解析）。

---

## 前置依赖

```bash
pip install requests pdfplumber openpyxl python-docx lxml
```

---

## 执行

本 skill 附带了 batch JSON 文件 `scripts/process.json`，两步完成全部流程。

**严格按照以下格式调用 `run_tool_batch`，使用 `file_path` 加载文件执行。不要自行构造 `actions` 列表内联传入。**

`run_tool_batch` 的 `file_path` 需要**绝对路径**。你在读取本 SKILL.md 时看到的目录路径即为本 skill 的绝对目录，请用它拼接出完整的 `file_path`。

```
run_tool_batch(
  file_path="<本skill目录>/scripts/process.json",
  args={
    "skill_dir": "<本skill目录>",
    "work_dir": "<工作目录>"
  }
)
```

### Batch 参数

| 参数 | 说明 | 示例值 |
|------|------|--------|
| `skill_dir` | 本 skill 的绝对路径（AI 运行时自动获取） | `<运行时动态获取>` |
| `work_dir` | 临时工作目录（items.json / cookies.json 输出位置） | `<运行时动态获取>` |

**必须传入这两个参数，不要传 `args={}`，否则 `${args.skill_dir}` 和 `${args.work_dir}` 不会展开。**

### Batch 步骤说明

| 步骤 | 工具 | 做什么 | 超时 |
|------|------|--------|------|
| Step 0 | `execute_shell_command` | 运行 `hgr_discover.py`：直接调用人遗系统 API → 获取全部批次文件信息 → 输出 `items.json` + `cookies.json`（无需浏览器） | 60s |
| Step 1 | `execute_shell_command` | 运行 `hgr_main.py`：读取 items.json + cookies.json → 增量判断 → 下载新批次 PDF/DOCX → 提取表格 → 分类写入 Excel → 输出简报 | 300s |

### Batch 失败处理

如果 `run_tool_batch` 执行失败：

1. 先检查「Batch 参数」中的 `skill_dir` 和 `work_dir` 是否已传入实际值，不要使用空的 `args={}`。
2. 如果 Step 0（hgr_discover.py）失败：
   - 可能是 Playwright 未安装：`pip install playwright && python -m playwright install chromium`
   - 可能是网站改版导致选择器失效：改为手动用 `browser_use` 完成发现，详见下方「分步参考」
3. 如果 Step 1（hgr_main.py）失败：
   - 检查 `items.json` 和 `cookies.json` 是否存在且格式正确
   - 检查 PDF 是否能正常下载（Cookie 可能过期）
4. 如果仍然失败，参照下方「分步参考」手动逐步执行。
5. 执行完毕后，提示用户：「本次 batch 执行遇到问题，已改为手动完成。是否需要我用 edit_file 调整和优化这个 skill 的 batch 脚本？」

---

## 分步参考

以下步骤仅在 batch 失败需要调试或手动执行时参考。

### 阶段一：浏览器发现（替代 hgr_discover.py）

如果 Playwright 脚本不可用，使用 `browser_use` 手动完成：

**1. 打开网站**
```
browser_use action=open url=https://apply.hgrg.net/login
```
等待页面完全加载（约 3-5 秒）。

**2. 点击「项目公示」**
```
browser_use action=evaluate code="for(const a of document.querySelectorAll('a')){if(a.textContent.includes('项目公示')){a.click();break;}}"
```
为什么用 `evaluate`：该链接是 `<a>` 标签但无 `href`，`target="_blank"`，普通 click 找不到元素。

**3. 获取 API 返回的批次列表（最佳方案）**
点击"项目公示"后会触发 API 调用，直接用 `network_requests` 捕获：
```
browser_use action=network_requests
```
在结果中找到 `POST /api/backend/projectPublicity/fileInfo` 的响应体，这是核心数据源：
- URL: `https://apply.hgrg.net/api/backend/projectPublicity/fileInfo`
- 方法: POST
- 请求体: `{}`（无需额外参数）
- 响应: 全部 86 条记录的数组，每条包含 `fileName`、`publicityFileUrl`、`publicityFileId`

**⚠️ 过滤规则（重要！）**
所有文件名用 `"批审批结果"` 过滤（**不是** `"审批结果公示"`）：
- 2023年~2024年第9批：命名格式为 `...XX批审批结果`（无"公示"）
- 2024年第10批起：命名格式为 `...XX批审批结果公示`
- 排除关键词：`"审查结果"`、`"补充公示"`、`"计划通知"`、`"受理截止"`、`"会议时间"`
- 2024年第18批特殊：`...审批结果与补充公示信息`（含"补充公示"，如需提取需手动处理）

**4. 逐个点击链接，捕获 PDF 源 URL（弃用方案）**
此方案已不推荐使用。优先使用 API 直接获取数据（步骤3），后者一次性返回全部86条记录，无需逐个点击。

**5. 获取 Cookie**
```
browser_use action=cookies_get domain=apply.hgrg.net
```
保存为 `cookies.json`，格式：`{"cookie_name": "cookie_value", ...}`

**6. 保存 items.json**

格式：
```json
[
  {
    "title": "中国人类遗传资源行政许可事项2026年第11批审批结果公示",
    "pdf_url": "https://apply.hgrg.net/geneticbucket/20260628/xxx.pdf"
  }
]
```

### 阶段二：Python 处理（hgr_main.py）

```bash
cd <本skill目录>/scripts
python hgr_main.py --items-json <工作目录>/items.json --cookies-json <工作目录>/cookies.json --config <本skill目录>/config.ini
```

脚本自动完成：
- 读取 `metadata.json` 判断增量 vs 全量
- 下载新批次 PDF（带 Cookie，3 次重试）
- pdfplumber 提取表格，按审批号前缀分类
- 写入独立批次 Excel + 更新汇总 Excel
- 输出 markdown 简报

---

## 历史数据补充：科技部政务服务平台

对于 HGR API 无法覆盖的 2021~2023 年早期数据（早于 2023年第19批），使用 `scripts/hgr_most_scraper.py` 从 [科技部政务服务平台](https://fuwu.most.gov.cn/html/jgcx/index.html) 补充。

### 数据特征

| 属性 | 值 |
|------|-----|
| 覆盖范围 | 2021 年第 10 批 ~ 2023 年第 18 批 |
| 数据格式 | HTML 表格（无需浏览器） |
| 翻页方式 | URL 编号：`index.html`(第1页) ~ `index_25.html`(第25页) |
| 表结构 | 与 PDF/DOCX 完全一致（5 列采集/保藏 + 8 列国际合作） |
| 过滤条件 | 标题含「审批结果」且不含「简化流程」「审查结果」「备案情况」 |

### 运行方式

```bash
cd <本skill目录>/scripts
python hgr_most_scraper.py
```

脚本自动完成：
1. 遍历 25 页列表页，提取「审批结果」链接（支持中文数字批号如「第二十五批」）
2. 逐页下载详情页 HTML，解析表格
3. 分类写入独立批次文件 → 更新汇总 Excel（与现有数据合并、去重、排序）
4. 输出简报

### 数据源覆盖对照

| 数据源 | URL | 覆盖批次 | 技术方式 |
|--------|-----|----------|----------|
| HGR API | `apply.hgrg.net/api/...` | 2023-19 ~ 2026-11 | POST API + PDF/DOCX 解析 |
| MOST 平台 | `fuwu.most.gov.cn/html/jgcx/` | 2021-10 ~ 2023-18 | Python requests + HTML 解析 |

> ⚠️ MOST 爬虫通常只需运行**一次**（回填历史数据），后续只更新 HGR API 即可。

---

## 输出说明

### 文件结构

```
skills/HGRS/data/
├── metadata.json                                    # 已处理批次记录（HGR API）
├── 汇总_中国人类遗传资源行政许可事项.xlsx             # 汇总（4个Sheet），包含HGR API + MOST 全量数据
└── batches/
    ├── 中国人类遗传资源行政许可事项{year}年第{batch}批审批结果公示.xlsx  # HGR API 批次
    └── 中国人类遗传资源行政许可事项{year}年第{batch}批审批结果.xlsx      # MOST 批次
```

### 汇总 Excel 的 4 个 Sheet

| Sheet | 列 | 审批号前缀 |
|-------|-----|-----------|
| 采集审批 | 序号、批次、审批号、项目名称、申请单位、批准时间 | CJ |
| 保藏审批 | 序号、批次、审批号、项目名称、申请单位、批准时间 | BC |
| 国际科学研究合作审批 | 序号、批次、审批号、项目名称、医疗机构(组长单位)、申办方、合同研究组织、检测/数据单位、批准时间 | GH |
| 材料出境证明 | 序号、批次、审批号、项目名称、申请单位、批准时间 | CC |

### 增量更新逻辑

1. 读取 `data/metadata.json` 中的 `processed_batches` 列表
2. 与 `items.json` 对比，筛选出新批次
3. 按批次号从小到大处理（旧→新），最新批次插入汇总表
4. 每个批次处理完更新 `metadata.json`
5. 如无新批次，输出「无更新」简报

### ⚠️ 处理完成后注意事项

1. **重新排序 Excel**：数据是按处理顺序插入的（非时间顺序），需对其重新排序+编号
2. **修复 metadata.latest**：`hgr_main.py` 的迭代逻辑会覆盖 `latest_year/batch`，最终指向最后处理的批次而非真正的 latest。处理完成后需手动修正为该正确的最新批次号

### metadata.json 结构

```json
{
  "latest_year": 2026,
  "latest_batch": 11,
  "latest_title": "中国人类遗传资源行政许可事项2026年第11批审批结果公示",
  "processed_batches": ["2026-11"],
  "updated_at": "2026-06-30T12:00:00"
}
```

---

## 注意事项

1. **首次运行**：metadata.json 不存在时会自动创建，并处理 items.json 中的全部批次（全量模式）
2. **Cookie 有效期**：人遗系统的 session cookie 有时效，每次运行都会重新获取
3. **分类规则**：按审批号前缀自动分类，未知前缀默认归入「国际科学研究合作审批」
4. **容错**：单个批次 PDF 下载/解析失败不影响其他批次，跳过并记录日志
5. **编码**：Windows 下脚本已处理 UTF-8 输出编码问题
6. **无更新**：如果所有批次都已处理，输出「所有批次已是最新，无需更新」简报
7. **MOST 回填**：只需运行一次，不要重复覆盖汇总表（脚本内有去重逻辑，但仍需谨慎）
