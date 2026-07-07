"""
01_Filename2PubMedID.py

Look up PubMed IDs from article titles.

This version verifies every candidate PMID by fetching its PubMed title and
comparing it with the input title. That avoids accepting unrelated PubMed
search hits when Entrez returns a fuzzy match.
"""

from __future__ import annotations

import sys
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import html
import os
import re
import time
from difflib import SequenceMatcher

import pandas as pd
from Bio import Entrez

from pubmed_tools.config import API_DELAY
from pubmed_tools.utils import init_entrez, read_titles, save_excel


INPUT_FILE = "data/Name.txt"
OUTPUT_FOUND = "output/01_PMID_Results.xlsx"
OUTPUT_MISSING = "output/01_PMID_NotFound.xlsx"
OUTPUT_REALDOI = "data/01_RealDOI_v1.txt"

NOT_FOUND = "未找到"
REQUEST_FAILED = "请求失败"
_NOT_FOUND_VALUES = {NOT_FOUND, REQUEST_FAILED}

MAX_SEARCH_RESULTS = 10
TITLE_MATCH_THRESHOLD = 0.94


def _clean_title(title: str) -> str:
    return str(title or "").replace("\ufeff", "").strip()


def normalize_title(title: str) -> str:
    """Normalize titles for strict-but-practical comparison."""
    text = html.unescape(str(title or ""))
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.casefold()
    text = text.replace("β", "beta").replace("α", "alpha")
    text = text.replace("γ", "gamma").replace("δ", "delta")
    text = re.sub(r"[\W_]+", " ", text, flags=re.UNICODE)
    return re.sub(r"\s+", " ", text).strip()


def title_similarity(input_title: str, pubmed_title: str) -> float:
    left = normalize_title(input_title)
    right = normalize_title(pubmed_title)
    if not left or not right:
        return 0.0
    if left == right:
        return 1.0
    return SequenceMatcher(None, left, right).ratio()


def titles_match(input_title: str, pubmed_title: str) -> bool:
    return title_similarity(input_title, pubmed_title) >= TITLE_MATCH_THRESHOLD


def save_realdoi_txt(titles: list[str], output_txt: str = OUTPUT_REALDOI) -> list[str]:
    """Save unmatched titles as a one-title-per-line txt file."""
    seen = set()
    cleaned = []
    for title in titles:
        title = _clean_title(title)
        key = title.casefold()
        if title and key not in seen:
            seen.add(key)
            cleaned.append(title)

    os.makedirs(os.path.dirname(output_txt) or ".", exist_ok=True)
    with open(output_txt, "w", encoding="utf-8") as f:
        for title in cleaned:
            f.write(title + "\n")
    print(f"  未找到 PMID 的标题已保存至: {output_txt}")
    return cleaned


def export_realdoi_from_missing_xlsx(
    missing_xlsx: str = OUTPUT_MISSING,
    output_txt: str = OUTPUT_REALDOI,
) -> list[str]:
    """Export Article Title values from PMID_NotFound.xlsx to RealDOI.txt."""
    if not os.path.exists(missing_xlsx):
        raise FileNotFoundError(f"找不到未匹配结果文件: {missing_xlsx}")

    df = pd.read_excel(missing_xlsx, dtype=str).fillna("")
    if "Article Title" not in df.columns:
        raise ValueError(f"{missing_xlsx} 中缺少 Article Title 列")

    return save_realdoi_txt(df["Article Title"].tolist(), output_txt)


def _entrez_search(query: str, max_results: int = MAX_SEARCH_RESULTS) -> list[str]:
    handle = Entrez.esearch(db="pubmed", term=query, retmax=max_results)
    try:
        record = Entrez.read(handle)
    finally:
        handle.close()
    return [str(pmid) for pmid in record.get("IdList", [])]


def fetch_pubmed_titles(pmids: list[str]) -> dict[str, str]:
    """Return {pmid: article_title} for candidate PMIDs."""
    if not pmids:
        return {}

    handle = Entrez.efetch(db="pubmed", id=",".join(pmids), retmode="xml")
    try:
        records = Entrez.read(handle)
    finally:
        handle.close()

    titles: dict[str, str] = {}
    for article in records.get("PubmedArticle", []):
        medline = article.get("MedlineCitation", {})
        pmid = str(medline.get("PMID", ""))
        article_data = medline.get("Article", {})
        title = str(article_data.get("ArticleTitle", ""))
        if pmid and title:
            titles[pmid] = title
    return titles


def search_pmid_matches_by_title(title: str, max_results: int = MAX_SEARCH_RESULTS) -> list[dict]:
    """Search PubMed and keep only candidates whose fetched title matches."""
    title = _clean_title(title)
    if not title:
        return []

    seen = set()
    candidate_pmids: list[str] = []
    queries = [
        f'"{title}"[Title]',
        f"{title}[Title]",
    ]

    for query in queries:
        for pmid in _entrez_search(query, max_results=max_results):
            if pmid not in seen:
                seen.add(pmid)
                candidate_pmids.append(pmid)
        if candidate_pmids:
            break

    fetched_titles = fetch_pubmed_titles(candidate_pmids)
    matches = []
    for pmid in candidate_pmids:
        pubmed_title = fetched_titles.get(pmid, "")
        score = title_similarity(title, pubmed_title)
        if score >= TITLE_MATCH_THRESHOLD:
            matches.append({
                "PubMed ID": pmid,
                "Matched Title": pubmed_title,
                "Title Match Score": round(score, 4),
            })
    return matches


def search_pmid_by_title(title: str, max_results: int = MAX_SEARCH_RESULTS) -> str:
    """Backward-compatible API: return comma-separated verified PMIDs."""
    try:
        matches = search_pmid_matches_by_title(title, max_results=max_results)
        return ", ".join(match["PubMed ID"] for match in matches) if matches else NOT_FOUND
    except Exception as e:
        print(f"  检索出错: {e}")
        return REQUEST_FAILED


def main() -> None:
    init_entrez()
    titles = [_clean_title(title) for title in read_titles(INPUT_FILE)]
    titles = [title for title in titles if title]
    print(f"读取到 {len(titles)} 个文献标题，开始检索...\n")

    found = []
    missing = []

    for i, title in enumerate(titles, 1):
        print(f"[{i}/{len(titles)}] {title[:80]}...")
        try:
            matches = search_pmid_matches_by_title(title)
        except Exception as e:
            print(f"  -> {REQUEST_FAILED}: {e}")
            missing.append({"Article Title": title, "Notes": REQUEST_FAILED})
            time.sleep(API_DELAY)
            continue

        if not matches:
            print(f"  -> {NOT_FOUND}")
            missing.append({"Article Title": title, "Notes": NOT_FOUND})
        else:
            pmids = ", ".join(match["PubMed ID"] for match in matches)
            print(f"  -> {len(matches)} verified PMID(s): {pmids}")
            for match in matches:
                found.append({
                    "Article Title": title,
                    "PubMed ID": match["PubMed ID"],
                    "Matched Title": match["Matched Title"],
                    "Title Match Score": match["Title Match Score"],
                })

        time.sleep(API_DELAY)

    print(f"\n检索完成: 找到 {len(found)} 条，未找到 {len(missing)} 条")

    if found:
        save_excel(found, OUTPUT_FOUND)
    else:
        print("  无成功结果，跳过生成 01_PMID_Results.xlsx")

    if missing:
        save_excel(missing, OUTPUT_MISSING)
        save_realdoi_txt([item["Article Title"] for item in missing], OUTPUT_REALDOI)
    else:
        print("  全部找到，无需生成 01_PMID_NotFound.xlsx")


if __name__ == "__main__":
    main()
