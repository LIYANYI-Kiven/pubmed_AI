#!/usr/bin/env python3
"""
PubMed 文献 AI 筛选工具
根据用户自定义的筛选条件，通过 AI API 对文献进行批量分析和筛选。
筛选条件通过 api_config.json 中的 system_prompt / user_prompt_template 配置。
"""

import json
import os
import time
from typing import List, Dict
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
import requests


class LiteratureAnalyzer:
    """
    通用文献分析器，通过 AI API 对文献进行批量筛选。
    """

    def __init__(self, api_url: str, api_key: str, model: str,
                 system_prompt: str, user_prompt_template: str,
                 result_key: str, abstract_max_len: int = 500):
        """
        Args:
            api_url:               AI API 端点（兼容 OpenAI 格式）
            api_key:               API 密钥
            model:                 模型名称
            system_prompt:         系统提示词，描述 AI 的角色和任务
            user_prompt_template:  用户提示词模板，{n} 替换为文献数量，{articles} 替换为文献列表
            result_key:            JSON 结果中用于判断是否保留文献的布尔字段名
            abstract_max_len:      摘要最大字符数，超出则截断
        """
        self.api_url = api_url
        self.model = model
        self.system_prompt = system_prompt
        self.user_prompt_template = user_prompt_template
        self.result_key = result_key
        self.abstract_max_len = abstract_max_len
        self.headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }

    def analyze_batch(self, articles: List[Dict], batch_num: int) -> Dict:
        """
        分析一批文献，返回结构化结果。
        """
        prompt = self._build_prompt(articles)
        try:
            response = requests.post(
                self.api_url,
                headers=self.headers,
                json={
                    "model": self.model,
                    "messages": [
                        {"role": "system", "content": self.system_prompt},
                        {"role": "user",   "content": prompt}
                    ],
                    "temperature": 0.3
                },
                timeout=120
            )
            response.raise_for_status()

            result = response.json()
            content = result.get("choices", [{}])[0].get("message", {}).get("content", "{}")

            # 清理 markdown 代码块标记
            content = content.strip()
            if content.startswith("```json"):
                content = content[7:]
            elif content.startswith("```"):
                content = content[3:]
            if content.endswith("```"):
                content = content[:-3]
            content = content.strip()

            analysis_result = json.loads(content)
            return {
                "batch_num": batch_num,
                "success": True,
                "results": analysis_result.get("articles", [])
            }

        except requests.exceptions.HTTPError as e:
            msg = f"HTTP错误: {e}\n响应内容: {response.text[:300]}"
        except json.JSONDecodeError as e:
            raw = response.text[:300] if 'response' in locals() else "N/A"
            msg = f"JSON解析错误: {e}\n响应内容: {raw}"
        except Exception as e:
            msg = f"{type(e).__name__}: {e}"

        print(f"  ⚠️  批次 {batch_num} 失败: {msg}")
        return {"batch_num": batch_num, "success": False, "error": msg, "results": []}

    def _build_prompt(self, articles: List[Dict]) -> str:
        articles_text = []
        for i, article in enumerate(articles, 1):
            abstract = article.get("Abstract", "N/A")
            if not isinstance(abstract, str):
                abstract = "N/A"
            elif len(abstract) > self.abstract_max_len:
                abstract = abstract[:self.abstract_max_len] + "..."

            articles_text.append(
                f"{i}. PMID:{article.get('PMID', 'N/A')}\n"
                f"Title: {article.get('Title', 'N/A')}\n"
                f"Abstract: {abstract}\n"
            )

        # 只替换 {n} 和 {articles}，不处理其他花括号（如 JSON 示例中的 { }）
        prompt = self.user_prompt_template
        prompt = prompt.replace("{n}", str(len(articles)))
        prompt = prompt.replace("{articles}", "".join(articles_text))
        return prompt


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------

def load_excel(file_path: str) -> pd.DataFrame:
    print(f"\n[1/4] 读取文献数据: {file_path}")
    df = pd.read_excel(file_path)
    print(f"  ✅ 共 {len(df)} 篇文献")
    return df


def split_into_batches(df: pd.DataFrame, batch_size: int) -> List[List[Dict]]:
    print(f"\n[2/4] 分批（每批 {batch_size} 篇）...")
    records = df.to_dict("records")
    batches = [records[i:i + batch_size] for i in range(0, len(records), batch_size)]
    print(f"  ✅ 共 {len(batches)} 批")
    return batches


def analyze_batches(batches: List[List[Dict]], analyzer: LiteratureAnalyzer,
                    output_dir: str, sleep_sec: float) -> List[str]:
    print(f"\n[3/4] 调用 AI API 分析文献...")
    os.makedirs(output_dir, exist_ok=True)
    json_files = []

    for i, batch in enumerate(batches, 1):
        print(f"  处理批次 {i}/{len(batches)}...", end=" ", flush=True)
        result = analyzer.analyze_batch(batch, i)

        json_file = os.path.join(output_dir, f"batch_{i:03d}.json")
        with open(json_file, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        json_files.append(json_file)

        print("✅" if result["success"] else "❌")

        if i < len(batches):
            time.sleep(sleep_sec)

    print(f"  ✅ 全部批次处理完成")
    return json_files


def compile_results(json_files: List[str], original_df: pd.DataFrame,
                    output_file: str, result_key: str):
    print(f"\n[4/4] 整理筛选结果...")

    # 汇总所有 AI 分析结果
    all_results: Dict[str, Dict] = {}
    success_count = 0
    for json_file in json_files:
        with open(json_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        if data.get("success"):
            success_count += 1
            for article in data.get("results", []):
                pmid = str(article.get("pmid", ""))
                if pmid:
                    all_results[pmid] = article

    print(f"  成功批次: {success_count}/{len(json_files)}")
    print(f"  获得分析结果: {len(all_results)} 篇")

    # 筛选符合条件的文献
    filtered = []
    for _, row in original_df.iterrows():
        pmid = str(row.get("PMID", ""))
        if pmid in all_results and all_results[pmid].get(result_key):
            record = row.to_dict()
            record["AI_Confidence"] = all_results[pmid].get("confidence", "")
            record["AI_Reason"]     = all_results[pmid].get("reason", "")
            filtered.append(record)

    if not filtered:
        print("  ⚠️  未筛选出符合条件的文献")
        return

    result_df = pd.DataFrame(filtered)
    result_df.reset_index(drop=True, inplace=True)
    result_df.index += 1

    os.makedirs(os.path.dirname(output_file) or ".", exist_ok=True)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=True, index_label="No.", sheet_name="Results")
        ws = writer.sheets["Results"]

        # 表头样式
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions

        # 自动列宽（简单估算）
        col_widths = {"A": 6, "B": 12, "C": 60, "D": 80,
                      "E": 20, "F": 50, "G": 14, "H": 12, "I": 60}
        for col, width in col_widths.items():
            ws.column_dimensions[col].width = width

        # 隔行变色
        even_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        odd_fill  = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            row_fill = even_fill if row_idx % 2 == 0 else odd_fill
            for cell in row:
                cell.fill = row_fill
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            ws.row_dimensions[row_idx].height = 50

    print(f"  ✅ 筛选出 {len(result_df)} 篇文献")
    print(f"  📁 结果文件: {output_file}")


def load_api_config(path: str = "api_config.json") -> Dict:
    if not os.path.exists(path):
        raise FileNotFoundError(f"配置文件不存在: {path}")
    with open(path, "r", encoding="utf-8") as f:
        config = json.load(f)

    # 从外部文本文件加载提示词（优先级高于 JSON 内嵌）
    for key, file_key in [("system_prompt", "system_prompt_file"),
                           ("user_prompt_template", "user_prompt_file")]:
        file_path = config.get(file_key)
        if file_path:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"提示词文件不存在: {file_path}（配置项: {file_key}）")
            with open(file_path, "r", encoding="utf-8") as f:
                config[key] = f.read()

    return config


# ---------------------------------------------------------------------------
# 主函数
# ---------------------------------------------------------------------------

def main():
    print("=" * 70)
    print("  PubMed 文献 AI 筛选工具")
    print("=" * 70)

    config = load_api_config()

    INPUT_FILE   = config.get("input_file",  "output/pubmed_results.xlsx")
    OUTPUT_DIR   = config.get("output_dir",  "output-1")
    OUTPUT_FILE  = config.get("output_file", "output/filtered_results.xlsx")
    BATCH_SIZE   = config.get("batch_size",  10)
    SLEEP_SEC    = config.get("sleep_sec",   1.0)
    API_URL      = config.get("api_url",     "https://api.openai.com/v1/chat/completions")
    API_KEY      = config.get("api_key")     or os.getenv("OPENAI_API_KEY", "")
    MODEL        = config.get("model",       "gpt-4o-mini")
    RESULT_KEY   = config.get("result_key",  "is_relevant")
    ABSTRACT_LEN = config.get("abstract_max_len", 500)

    SYSTEM_PROMPT = config.get("system_prompt",
        "You are an expert literature analyst. Evaluate each article strictly based on the given criteria.")

    USER_PROMPT_TEMPLATE = config.get("user_prompt_template",
        """Analyze the following {n} articles and determine whether each one meets the screening criteria.

Return ONLY a JSON object in this exact format:
{{
  "articles": [
    {{
      "pmid": "PMID_VALUE",
      "is_relevant": true,
      "confidence": "high",
      "reason": "brief reason"
    }}
  ]
}}

Articles:
{articles}""")

    if not API_KEY or API_KEY == "your_api_key_here":
        print("\n⚠️  未设置 API 密钥，请在 api_config.json 中填写 api_key")
        return

    print(f"\n  输入文件 : {INPUT_FILE}")
    print(f"  输出目录 : {OUTPUT_DIR}")
    print(f"  输出文件 : {OUTPUT_FILE}")
    print(f"  批次大小 : {BATCH_SIZE}")
    print(f"  模型     : {MODEL}")
    print(f"  筛选字段 : {RESULT_KEY}")

    df       = load_excel(INPUT_FILE)
    batches  = split_into_batches(df, BATCH_SIZE)
    analyzer = LiteratureAnalyzer(
        api_url=API_URL, api_key=API_KEY, model=MODEL,
        system_prompt=SYSTEM_PROMPT,
        user_prompt_template=USER_PROMPT_TEMPLATE,
        result_key=RESULT_KEY,
        abstract_max_len=ABSTRACT_LEN
    )
    json_files = analyze_batches(batches, analyzer, OUTPUT_DIR, SLEEP_SEC)
    compile_results(json_files, df, OUTPUT_FILE, RESULT_KEY)

    print("\n" + "=" * 70)
    print("  ✅ 完成！")
    print("=" * 70 + "\n")


if __name__ == "__main__":
    main()
