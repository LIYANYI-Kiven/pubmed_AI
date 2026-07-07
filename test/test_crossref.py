#!/usr/bin/env python3
"""
Crossref 文献检索功能测试
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from crossref_search import CrossrefClient, parse_item, export_excel, export_json

TEST_QUERY   = "plant protein glycosylation site"
TEST_DOI     = "10.1038/s41586-020-2649-2"  # Nature 2020 经典文章
OUT_EXCEL    = "test/output/crossref_results.xlsx"
OUT_JSON     = "test/output/crossref_raw.json"

os.makedirs("test/output", exist_ok=True)


def test_basic_search():
    print("=" * 60)
    print("测试 1: 基本关键词检索")
    print("=" * 60)
    print(f"查询: {TEST_QUERY}")

    client = CrossrefClient(mailto="test@example.com", sleep_sec=1.0)
    data = client.search(query=TEST_QUERY, rows=10)

    message = data.get("message", {})
    items = message.get("items", [])
    total = message.get("total-results", 0)

    assert isinstance(items, list), "items 应为列表"
    assert len(items) > 0, "应至少返回 1 条结果"

    print(f"✅ 数据库总计: {total:,} 条")
    print(f"✅ 本次返回: {len(items)} 条\n")

    print("前3条结果：")
    for i, item in enumerate(items[:3], 1):
        record = parse_item(item)
        print(f"  {i}. {record['Title'][:70]}")
        print(f"     作者: {record['First Author']}  期刊: {record['Journal']}")
        print(f"     发表: {record['Published']}  引用: {record['Cited By']}  DOI: {record['DOI']}")

    return items


def test_filter_by_year():
    print("\n" + "=" * 60)
    print("测试 2: 按年份过滤")
    print("=" * 60)

    client = CrossrefClient(mailto="test@example.com", sleep_sec=1.0)
    data = client.search(
        query="Arabidopsis glycosylation",
        rows=10,
        filter_from_year=2020,
        filter_to_year=2024,
        filter_type="journal-article",
    )

    items = data.get("message", {}).get("items", [])
    print(f"✅ 2020-2024 年返回 {len(items)} 条")

    for item in items[:3]:
        record = parse_item(item)
        pub_year = record["Published"][:4] if record["Published"] else ""
        print(f"   {pub_year}  {record['Title'][:60]}")

    return items


def test_sort_by_citations():
    print("\n" + "=" * 60)
    print("测试 3: 按引用次数排序")
    print("=" * 60)

    client = CrossrefClient(mailto="test@example.com", sleep_sec=1.0)
    data = client.search(
        query="plant protein glycosylation",
        rows=5,
        sort="is-referenced-by-count",
        order="desc",
    )

    items = data.get("message", {}).get("items", [])
    print(f"✅ 按引用次数降序返回 {len(items)} 条")

    cite_counts = []
    for item in items:
        record = parse_item(item)
        cite_counts.append(record["Cited By"])
        print(f"   引用: {record['Cited By']:5d}  {record['Title'][:60]}")

    # 验证引用次数是降序的
    for i in range(len(cite_counts) - 1):
        assert cite_counts[i] >= cite_counts[i+1], "引用次数应降序排列"
    print("✅ 引用次数排序验证通过")

    return items


def test_search_all_paginated():
    print("\n" + "=" * 60)
    print("测试 4: 分页获取多条结果")
    print("=" * 60)

    client = CrossrefClient(mailto="test@example.com", sleep_sec=1.0)
    items = client.search_all(
        query="glycosylation Arabidopsis",
        max_results=25,
        batch_size=20,
        filter_type="journal-article",
    )

    assert len(items) > 0, "应返回结果"
    print(f"✅ 分页获取共 {len(items)} 篇文献")

    return items


def test_get_by_doi():
    print("\n" + "=" * 60)
    print(f"测试 5: 通过 DOI 获取文献")
    print("=" * 60)
    print(f"DOI: {TEST_DOI}")

    client = CrossrefClient(mailto="test@example.com", sleep_sec=1.0)
    item = client.get_by_doi(TEST_DOI)

    assert item is not None, f"DOI {TEST_DOI} 应能获取到文献"
    record = parse_item(item)

    print(f"✅ 标题 : {record['Title'][:80]}")
    print(f"   期刊 : {record['Journal']}")
    print(f"   发表 : {record['Published']}")
    print(f"   引用 : {record['Cited By']}")
    print(f"   摘要 : {record['Abstract'][:100]}..." if record['Abstract'] else "   摘要 : 无")

    return item


def test_export():
    print("\n" + "=" * 60)
    print("测试 6: 导出 Excel 和 JSON")
    print("=" * 60)

    client = CrossrefClient(mailto="test@example.com", sleep_sec=1.0)
    items = client.search_all(
        query="plant protein N-glycosylation mass spectrometry",
        max_results=20,
        filter_from_year=2019,
        filter_type="journal-article",
    )

    records = [parse_item(item) for item in items]

    export_excel(records, OUT_EXCEL)
    export_json(records, items, OUT_JSON)

    assert os.path.exists(OUT_EXCEL), "Excel 文件未生成"
    assert os.path.exists(OUT_JSON),  "JSON 文件未生成"

    print(f"✅ Excel: {OUT_EXCEL}")
    print(f"✅ JSON : {OUT_JSON}")

    return records


if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("  Crossref 文献检索功能测试")
    print("=" * 60 + "\n")

    passed = 0
    failed = 0

    tests = [
        ("基本关键词检索",   test_basic_search),
        ("按年份过滤",       test_filter_by_year),
        ("按引用次数排序",   test_sort_by_citations),
        ("分页获取多条结果", test_search_all_paginated),
        ("通过 DOI 获取",    test_get_by_doi),
        ("导出 Excel/JSON",  test_export),
    ]

    for name, fn in tests:
        try:
            fn()
            passed += 1
        except AssertionError as e:
            print(f"❌ 断言失败 [{name}]: {e}")
            failed += 1
        except Exception as e:
            print(f"❌ 测试失败 [{name}]: {type(e).__name__}: {e}")
            failed += 1

    print(f"\n{'=' * 60}")
    print(f"  测试结果: {passed} 通过 / {failed} 失败")
    print(f"{'=' * 60}\n")
    sys.exit(0 if failed == 0 else 1)
