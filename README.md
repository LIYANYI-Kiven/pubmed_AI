# PubMed 植物蛋白质糖基化文献分析工具

自动从 PubMed 检索文献，并通过 AI API 筛选出包含植物蛋白质糖基化位点鉴定的研究。

## 工作流程

```
PubMed API → 下载文献 → Excel文件 → AI分析 → 筛选结果 → Excel报告
```

1. `pubmed-improved.py` — 从 PubMed 检索并下载文献到 Excel
2. `analyze_glycosylation.py` — 调用 AI API 分析文献，筛选出包含糖基化位点鉴定的文章

## 快速开始

### 1. 安装依赖

```bash
pip install requests pandas openpyxl
```

### 2. 配置检索参数

编辑 `config.json`：

```json
{
  "date_start": "2017",
  "date_end": "2025",
  "api_key": null,
  "batch_size": 100,
  "sleep_sec": 0.4,
  "out_file": "output/plant_glyco_simple.xlsx",
  "query": "(((Plants[MeSH Terms] OR \"Plant Sciences\"[MeSH Terms] OR plant*[Title/Abstract] OR crop*[Title/Abstract] OR \"Oryza sativa\"[Title/Abstract] OR Arabidopsis[Title/Abstract] OR \"Zea mays\"[Title/Abstract]) AND (protein*[Title/Abstract] OR proteome*[Title/Abstract])) AND (\"Glycosylation\"[MeSH Terms] OR glycosylat*[Title/Abstract] OR \"glycan binding\"[Title/Abstract] OR \"N-glycosylation\"[Title/Abstract] OR \"O-glycosylation\"[Title/Abstract])) NOT \"Review\"[Publication Type]"
}
```

| 字段 | 说明 |
|------|------|
| `date_start` | 检索起始年份，格式 `YYYY` 或 `YYYY/MM/DD` |
| `date_end` | 检索结束年份 |
| `api_key` | PubMed API Key（可选，有 key 时速率更高） |
| `batch_size` | 每批下载数量，建议 100 |
| `sleep_sec` | 请求间隔（秒），无 API Key 时建议 ≥ 0.4 |
| `out_file` | 输出 Excel 文件路径 |
| `query` | PubMed 检索式，双引号需转义为 `\"` |

> **提示**：检索式中的通配符 `*` 很重要，`plant*` 可以匹配 plants、planting 等，`protein*` 可以匹配 proteins、proteomics 等。

### 3. 下载文献

```bash
python pubmed-improved.py
```

输出文件保存在 `output/plant_glyco_simple.xlsx`。

### 4. 配置 AI 分析参数

编辑 `api_config.json`：

```json
{
  "api_url": "https://api.deepseek.com/v1/chat/completions",
  "api_key": "your_api_key_here",
  "model": "deepseek-chat",
  "batch_size": 10,
  "sleep_sec": 1.0,
  "input_file": "output/plant_glyco_simple.xlsx",
  "output_dir": "output-1",
  "output_file": "output/glycosylation_sites_identified.xlsx"
}
```

| 字段 | 说明 |
|------|------|
| `api_url` | AI API 端点，兼容 OpenAI 格式 |
| `api_key` | AI API 密钥 |
| `model` | 使用的模型名称 |
| `batch_size` | 每批分析的文献数量，建议 10 |
| `sleep_sec` | API 请求间隔（秒） |
| `input_file` | 输入的 Excel 文件（步骤3的输出） |
| `output_dir` | 中间 JSON 文件保存目录 |
| `output_file` | 最终筛选结果 Excel 文件路径 |

支持任何兼容 OpenAI 格式的 API，例如：
- [DeepSeek](https://platform.deepseek.com/)：`https://api.deepseek.com/v1/chat/completions`
- [OpenAI](https://platform.openai.com/)：`https://api.openai.com/v1/chat/completions`

### 5. 运行 AI 分析

```bash
python analyze_glycosylation.py
```

程序会：
- 将文献分批（每批 10 篇）提交给 AI
- 每批结果保存为 `output-1/batch_001.json` 等文件
- 最终将判定为"包含糖基化位点鉴定"的文献整理成 Excel

## 输出文件说明

### output/plant_glyco_simple.xlsx

PubMed 下载的全部文献，包含字段：

| 字段 | 说明 |
|------|------|
| PMID | PubMed 文献 ID |
| Title | 文章标题 |
| Abstract | 摘要 |
| First Author | 第一作者 |
| MeSH Terms | MeSH 主题词 |
| Publication Date | 发表日期 |

### output-1/batch_XXX.json

每批 AI 分析的中间结果：

```json
{
  "batch_num": 1,
  "success": true,
  "results": [
    {
      "pmid": "12345678",
      "contains_glycosylation_sites": true,
      "confidence": "high",
      "reason": "使用质谱技术鉴定了水稻蛋白的N-糖基化位点"
    }
  ]
}
```

### output/glycosylation_sites_identified.xlsx

最终筛选结果，在原始字段基础上新增：

| 字段 | 说明 |
|------|------|
| Confidence | AI 判断置信度（high / medium / low） |
| Analysis Reason | AI 判断理由 |

## AI 判断标准

AI 会根据以下标准判断文献是否包含植物蛋白质糖基化位点鉴定：

1. 文献是否涉及**植物蛋白质**的糖基化研究
2. 是否包含糖基化**位点的鉴定或定位**信息
3. 是否使用了**质谱、生物信息学或实验方法**鉴定糖基化位点

## 注意事项

- PubMed 无 API Key 时每秒最多 3 次请求，有 Key 时最多 10 次
- AI API 调用会产生费用，请注意控制 `batch_size` 和文献总量
- 如果某批次分析失败，程序会继续处理下一批，不会中断
- 重新运行 `analyze_glycosylation.py` 会覆盖已有的 JSON 文件

## 文件结构

```
.
├── config.json              # PubMed 检索配置
├── api_config.json          # AI API 配置
├── pubmed-improved.py       # PubMed 文献下载脚本
├── pubmed-simple.py         # PubMed 下载脚本（简化版）
├── analyze_glycosylation.py # AI 分析脚本
├── output/                  # 输出目录（不上传 git）
│   ├── plant_glyco_simple.xlsx
│   └── glycosylation_sites_identified.xlsx
└── output-1/                # AI 分析中间结果（不上传 git）
    ├── batch_001.json
    └── ...
```
