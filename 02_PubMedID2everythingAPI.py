"""
PubMedID2everythingAPI.py - 通过 NCBI API 批量获取文献标题、PMCID 和全文链接

输入: data/pmid_list.txt（支持 txt / csv / xlsx）
输出:
  output/pubmed_api_results.xlsx      完整信息表（标题、PMCID、所有链接）
  output/pubmed_links.xlsx            精简链接表（PMID、全文链接、PubMed链接）
"""

import sys
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import os
from pubmed_tools.utils import read_pmids, save_excel, fetch_pubmed_records

INPUT_FILE = "data/pmid_list.txt"
OUTPUT_FULL  = "output/02_pubmed_api_results.xlsx"
OUTPUT_LINKS = "output/02_pubmed_links.xlsx"
# 如果 PMID 不在第一列，填写列名；否则保持 None
TARGET_COLUMN = None


def build_fulltext_link(pmcid: str, doi: str) -> str:
    """
    按优先级返回最佳全文链接：
      1. PMC 链接（开放获取，优先）
      2. DOI 链接（次选）
      3. 空字符串（均无）
    """
    if pmcid and pmcid != "N/A":
        return f"https://www.ncbi.nlm.nih.gov/pmc/articles/{pmcid}/"
    if doi and doi != "N/A":
        return f"https://doi.org/{doi}"
    return ""


def build_all_links(pmid: str, pmcid: str, doi: str) -> str:
    """组装完整链接字符串（用于详细结果表）"""
    links = [f"PubMed: https://pubmed.ncbi.nlm.nih.gov/{pmid}/"]
    if pmcid and pmcid != "N/A":
        links.append(f"PMC: https://www.ncbi.nlm.nih.gov/pmc/articles/{pmcid}/")
    if doi and doi != "N/A":
        links.append(f"DOI: https://doi.org/{doi}")
    return "\n".join(links)


def main():
    if not os.path.exists(INPUT_FILE):
        print(f"找不到输入文件：{INPUT_FILE}")
        return

    pmids = read_pmids(INPUT_FILE, TARGET_COLUMN)
    print(f"读取到 {len(pmids)} 个 PMID，开始获取文献信息...\n")
    if not pmids:
        return

    data_map = fetch_pubmed_records(pmids)

    full_results  = []  # 完整信息表
    links_results = []  # 精简链接表

    for pmid in pmids:
        info = data_map.get(pmid, {"title": "未找到", "pmcid": "N/A", "doi": "N/A"})
        pmcid = info["pmcid"]
        doi   = info["doi"]

        full_results.append({
            "PubMed ID":      pmid,
            "Article Title":  info["title"],
            "PMCID":          pmcid,
            "Full Text Links": build_all_links(pmid, pmcid, doi),
        })

        links_results.append({
            "PubMed ID":       pmid,
            "Full Text Link":  build_fulltext_link(pmcid, doi),
            "PubMed Link":     f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/",
        })

    save_excel(full_results,  OUTPUT_FULL)
    save_excel(links_results, OUTPUT_LINKS)


if __name__ == "__main__":
    main()
