#!/usr/bin/env python3
"""pubmed_search.py 功能测试"""
import sys, os, json, datetime
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pubmed_search import fetch_records, export_excel, export_json, record_to_json_entry

os.makedirs("test/output", exist_ok=True)

print("测试: 检索5篇文献...")
records = fetch_records(
    query="Arabidopsis N-glycosylation mass spectrometry",
    date_start="2022", date_end="2025",
    api_key=None, batch_size=5, sleep_sec=0.5
)

assert len(records) > 0, "应至少返回1篇文献"
print(f"获取 {len(records)} 篇文献")

r = records[0]
print(f"第一篇: {r['Title'][:60]}")
print(f"  PMID={r['PMID']}  PMCID={r['PMCID']}  DOI={r['DOI'][:40] if r['DOI'] else ''}")
print(f"  关键词: {r['Keywords'][:60]}")
print(f"  PubMed URL: {r['PubMed URL']}")
print(f"  Full Text: {r['Full Text URL']}")

# 导出
meta = {
    "query": "Arabidopsis N-glycosylation mass spectrometry",
    "date_start": "2022", "date_end": "2025",
    "fetched_at": datetime.datetime.now().isoformat()
}
export_excel(records, "test/output/pubmed_search_test.xlsx")
export_json(records, "test/output/pubmed_search_test.json", meta)

# 验证JSON结构
with open("test/output/pubmed_search_test.json", encoding="utf-8") as f:
    data = json.load(f)

assert "meta" in data
assert "articles" in data
assert data["meta"]["total"] == len(records)

entry = data["articles"][0]
print(f"\nJSON 字段: {list(entry.keys())}")
print(f"links: {entry['links']}")
assert "pmid" in entry
assert "abstract" in entry
assert "keywords" in entry
assert "mesh_terms" in entry
assert "links" in entry
assert "pubmed" in entry["links"]

print("\n✅ 所有测试通过")
