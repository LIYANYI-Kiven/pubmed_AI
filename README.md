# PubMed 文献检索与 AI 筛选工具

从 PubMed 批量下载文献，并通过 AI API 按照自定义条件进行智能筛选，输出格式化 Excel 报告。

适用于任何研究方向的文献调研，只需修改检索式和 AI 提示词即可。

---

## 工作流程

```
config.json          api_config.json
     │                      │
     ▼                      ▼
pubmed-improved.py   analyze_glycosylation.py
     │                      │
     ▼                      ▼
output/pubmed_results.xlsx  →  output/filtered_results.xlsx
                                output-1/batch_*.json (中间结果)
```

**Step 1** — `pubmed-improved.py`：根据检索式从 PubMed 下载文献，保存为 Excel  
**Step 2** — `analyze_glycosylation.py`：将文献分批提交给 AI，按条件筛选，输出最终 Excel

---

## 安装依赖

```bash
pip install requests pandas openpyxl
```

---

## Step 1：配置并下载文献

编辑 `config.json`：

```json
{
  "date_start": "2017",
  "date_end":   "2025",
  "api_key":    null,
  "batch_size": 100,
  "sleep_sec":  0.4,
  "out_file":   "output/pubmed_results.xlsx",
  "query":      "your PubMed query here"
}
```

| 字段 | 说明 |
|------|------|
| `date_start` / `date_end` | 年份范围，格式 `YYYY` 或 `YYYY/MM/DD` |
| `api_key` | PubMed API Key（可选，申请地址见下方）|
| `batch_size` | 每批请求数量，建议 100 |
| `sleep_sec` | 请求间隔（秒）。无 Key 时建议 ≥ 0.4，有 Key 时可设为 0.1 |
| `out_file` | 输出 Excel 路径 |
| `query` | PubMed 检索式，JSON 中双引号需转义为 `\"` |

### 如何构建检索式

推荐在 [PubMed 网页端](https://pubmed.ncbi.nlm.nih.gov/advanced/) 的 Advanced Search 中构建检索式，确认结果数量后，复制 URL 中的 `term=` 参数，URL 解码后填入 `query` 字段。

**注意事项：**
- 通配符 `*` 很重要：`plant*` 可匹配 plant/plants/planting，`protein*` 可匹配 protein/proteins/proteomics
- 网页端 URL 中的 `+` 对应空格，`%22` 对应 `"`，`%5B` 对应 `[`
- 日期范围建议使用年份（如 `2017`），而非精确日期，与网页端行为一致

**检索式示例：**

```
# 植物蛋白质糖基化
(((Plants[MeSH Terms] OR plant*[Title/Abstract]) AND protein*[Title/Abstract]) AND glycosylat*[Title/Abstract]) NOT "Review"[Publication Type]

# CRISPR 基因编辑
(CRISPR[Title/Abstract] OR "Cas9"[Title/Abstract]) AND ("gene editing"[Title/Abstract] OR "genome editing"[Title/Abstract])

# 癌症免疫治疗
("immune checkpoint"[Title/Abstract] OR "PD-1"[Title/Abstract] OR "PD-L1"[Title/Abstract]) AND (cancer[MeSH Terms] OR tumor[Title/Abstract])
```

### 获取 PubMed API Key（可选）

有 API Key 时速率限制从 3次/秒 提升到 10次/秒，下载大量文献时明显更快。

申请地址：https://www.ncbi.nlm.nih.gov/account/

申请后填入 `config.json` 的 `api_key` 字段。

### 运行下载

```bash
python pubmed-improved.py
```

---

## Step 2：配置 AI 筛选

编辑 `api_config.json`：

```json
{
  "api_url":   "https://api.deepseek.com/v1/chat/completions",
  "api_key":   "your_api_key_here",
  "model":     "deepseek-chat",
  "batch_size": 10,
  "sleep_sec":  1.0,
  "abstract_max_len": 500,

  "input_file":  "output/pubmed_results.xlsx",
  "output_dir":  "output-1",
  "output_file": "output/filtered_results.xlsx",

  "result_key": "is_relevant",

  "system_prompt": "你是一位专业的文献分析专家，严格按照给定标准判断每篇文献是否符合筛选条件。",

  "user_prompt_template": "请分析以下 {n} 篇文献，判断每篇是否符合筛选条件。\n\n筛选条件：\n1. 文献涉及植物蛋白质的糖基化研究\n2. 包含糖基化位点的鉴定或定位信息\n3. 使用了质谱、生信或实验方法鉴定位点\n\n请严格按照以下 JSON 格式返回，不要输出其他内容：\n{{\n  \"articles\": [\n    {{\"pmid\": \"PMID\", \"is_relevant\": true, \"confidence\": \"high\", \"reason\": \"简短理由\"}}\n  ]\n}}\n\n文献列表：\n{articles}"
}
```

### 字段说明

| 字段 | 说明 |
|------|------|
| `api_url` | AI API 端点，兼容 OpenAI 格式 |
| `api_key` | AI API 密钥 |
| `model` | 模型名称 |
| `batch_size` | 每批文献数量，建议 10（过多会超出 token 限制）|
| `sleep_sec` | API 请求间隔（秒），避免限流 |
| `abstract_max_len` | 摘要最大字符数，超出截断，建议 500 |
| `input_file` | Step 1 生成的 Excel 文件路径 |
| `output_dir` | 中间 JSON 文件保存目录 |
| `output_file` | 最终筛选结果 Excel 路径 |
| `result_key` | AI 返回 JSON 中的布尔判断字段名，需与提示词中一致 |
| `system_prompt` | 系统提示词，定义 AI 角色 |
| `user_prompt_template` | 用户提示词模板（见下方详细说明）|

### 支持的 AI API

任何兼容 OpenAI Chat Completions 格式的 API 均可使用：

| 服务 | api_url | 推荐模型 |
|------|---------|---------|
| [DeepSeek](https://platform.deepseek.com/) | `https://api.deepseek.com/v1/chat/completions` | `deepseek-chat` |
| [OpenAI](https://platform.openai.com/) | `https://api.openai.com/v1/chat/completions` | `gpt-4o-mini` |
| [Moonshot (Kimi)](https://platform.moonshot.cn/) | `https://api.moonshot.cn/v1/chat/completions` | `moonshot-v1-8k` |
| [智谱 GLM](https://open.bigmodel.cn/) | `https://open.bigmodel.cn/api/paas/v4/chat/completions` | `glm-4-flash` |

---

## 如何修改筛选条件（提示词）

这是使用本工具最关键的配置。`user_prompt_template` 中有两个占位符：

- `{n}` — 自动替换为当前批次的文献数量
- `{articles}` — 自动替换为文献列表（PMID + 标题 + 摘要）

**AI 返回的 JSON 必须包含：**
- `pmid`：文献 PMID
- `is_relevant`（或你在 `result_key` 中指定的字段名）：`true` / `false`
- `confidence`：置信度，`high` / `medium` / `low`
- `reason`：判断理由

### 提示词示例

**示例 1：筛选包含临床试验数据的文献**

```json
"system_prompt": "You are a clinical research expert. Evaluate articles strictly based on the given criteria.",

"user_prompt_template": "Analyze the following {n} articles. Determine whether each one reports results from a clinical trial.\n\nCriteria:\n1. The study involves human participants\n2. It reports clinical outcomes or efficacy data\n3. It is a randomized controlled trial or observational study\n\nReturn ONLY this JSON:\n{{\n  \"articles\": [\n    {{\"pmid\": \"PMID\", \"is_relevant\": true, \"confidence\": \"high\", \"reason\": \"brief reason\"}}\n  ]\n}}\n\nArticles:\n{articles}"
```

**示例 2：筛选包含特定物种的研究**

```json
"system_prompt": "你是一位微生物学专家，专注于分析细菌耐药性研究文献。",

"user_prompt_template": "分析以下 {n} 篇文献，判断是否涉及大肠杆菌的抗生素耐药机制研究。\n\n筛选标准：\n1. 研究对象为大肠杆菌（E. coli）\n2. 涉及抗生素耐药性机制\n3. 包含实验数据\n\n返回 JSON：\n{{\n  \"articles\": [\n    {{\"pmid\": \"PMID\", \"is_relevant\": true, \"confidence\": \"high\", \"reason\": \"理由\"}}\n  ]\n}}\n\n文献：\n{articles}"
```

**示例 3：筛选包含特定实验方法的文献**

```json
"system_prompt": "You are a proteomics expert specializing in mass spectrometry-based research.",

"user_prompt_template": "Analyze {n} articles. Determine if each uses mass spectrometry to identify post-translational modifications.\n\nCriteria:\n1. Uses LC-MS/MS or similar mass spectrometry\n2. Identifies specific PTM sites (phosphorylation, ubiquitination, etc.)\n3. Provides site-level evidence\n\nReturn ONLY:\n{{\n  \"articles\": [\n    {{\"pmid\": \"PMID\", \"is_relevant\": true, \"confidence\": \"high\", \"reason\": \"reason\"}}\n  ]\n}}\n\nArticles:\n{articles}"
```

### 提示词编写建议

1. **明确判断标准**：列出 2-4 条具体、可验证的标准，避免模糊描述
2. **强调格式**：在提示词中明确要求"只返回 JSON，不输出其他内容"
3. **result_key 保持一致**：`result_key` 字段名必须与提示词中 JSON 格式里的布尔字段名完全一致
4. **控制批次大小**：摘要较长时适当减小 `batch_size`（建议 5-10），避免超出模型 token 限制
5. **语言一致性**：系统提示词和用户提示词建议使用同一语言

---

## 运行 AI 筛选

```bash
python analyze_glycosylation.py
```

程序会显示每批处理进度，失败的批次会标记 ❌ 但不会中断整体流程。

---

## 输出文件说明

### output/pubmed_results.xlsx

PubMed 下载的全部文献：

| 字段 | 说明 |
|------|------|
| PMID | PubMed 文献 ID |
| Title | 文章标题 |
| Abstract | 摘要 |
| First Author | 第一作者 |
| MeSH Terms | MeSH 主题词 |
| Publication Date | 发表日期 |

### output-1/batch_XXX.json

每批 AI 分析的中间结果，格式如下：

```json
{
  "batch_num": 1,
  "success": true,
  "results": [
    {
      "pmid": "12345678",
      "is_relevant": true,
      "confidence": "high",
      "reason": "Uses LC-MS/MS to identify N-glycosylation sites on rice proteins"
    }
  ]
}
```

中间结果保留在 `output-1/` 目录，方便排查问题或重新汇总。

### output/filtered_results.xlsx

最终筛选结果，在原始字段基础上新增：

| 字段 | 说明 |
|------|------|
| AI_Confidence | AI 判断置信度（high / medium / low）|
| AI_Reason | AI 判断理由 |

---

## 文件结构

```
.
├── config.json                 # PubMed 检索配置
├── api_config.json             # AI API 配置（含提示词）
├── pubmed-improved.py          # PubMed 文献下载脚本（推荐）
├── pubmed-simple.py            # PubMed 文献下载脚本（简化版）
├── analyze_glycosylation.py    # AI 文献筛选脚本
├── output/                     # 输出目录（不上传 git）
│   ├── pubmed_results.xlsx     # 下载的全部文献
│   └── filtered_results.xlsx  # AI 筛选后的文献
└── output-1/                   # AI 分析中间结果（不上传 git）
    ├── batch_001.json
    └── ...
```

---

## 常见问题

**Q：API 返回结果数量与网页端不一致？**  
A：检查以下几点：
1. 检索式中是否使用了通配符（`plant*` 而非 `plant`）
2. 日期范围建议使用年份格式（`2017`）而非精确日期（`2017/01/01`）
3. 在网页端复制完整 URL，URL 解码后对比检索式是否一致

**Q：AI 分析某批次失败怎么办？**  
A：失败的批次会保存 `"success": false` 的 JSON 文件。可以减小 `batch_size` 后重新运行，程序会覆盖已有文件。

**Q：如何只重新运行失败的批次？**  
A：目前程序会重新处理所有批次。如需跳过已成功的批次，可以手动删除失败的 JSON 文件，然后修改代码中的起始批次编号。

**Q：提示词中的 `{{` 和 `}}` 是什么？**  
A：JSON 字符串中的 `{{` 和 `}}` 是对花括号的转义，实际传给 AI 的内容是 `{` 和 `}`。这是 Python `str.format()` 的语法要求。
