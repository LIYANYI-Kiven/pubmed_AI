#!/usr/bin/env python3
"""
LitSense 2.0 功能测试
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from litsense_search import LitSenseClient, results_to_df, export_excel, export_json

TEST_QUERY  = "N-glycosylation sites identified in Arabidopsis proteins by mass spectrometry"
OUT_EXCEL_S = "test/output/litsense_sentence.xlsx"
OUT_EXCEL_P = "test/output/litsense_passage.xlsx"
OUT_JSON_S  = "test/output/litsense_sentence_raw.json"
OUT_JSON_P  = "test/output/litsense_passage_raw.json"

os.makedirs("test/output", exist_ok=True)


def test_sentence_search():
    print("=" * 60)
    print("测试 1: 句子模式检索")
    print("=" * 60)
    print(f"查询: {TEST_QUERY}\n")

    client = LitSenseClient(sleep_sec=1.0, timeout=90.0)
    results = client.search(query=TEST_QUERY, mode="sentence", rerank=True)

    assert isinstance(results, list), f"返回类型应为 list，实际为 {type(results)}"
    print(f"✅ 返回 {len(results)} 条结果")

    if results:
        first = results[0]
        print(f"\n前3条结果：")
        for i, r in enumerate(results[:3], 1):
            print(f"  {i}. PMID: {r.get('pmid','')}  Score: {r.get('score',''):.4f}")
            text = r.get("text", r.get("sentence", ""))
            print(f"     [{r.get('section','')}] {text[:120]}...")

        df = results_to_df({"sentence": results}, "sentence")
        assert not df.empty, "DataFrame 不应为空"
        export_excel(df, OUT_EXCEL_S)
        export_json({"sentence": results}, OUT_JSON_S)

        assert os.path.exists(OUT_EXCEL_S), "Excel 文件未生成"
        assert os.path.exists(OUT_JSON_S),  "JSON 文件未生成"
        print(f"\n✅ 文件已生成: {OUT_EXCEL_S}")
    else:
        print("⚠️  无结果返回（查询可能过于严格）")

    return results


def test_passage_search():
    print("\n" + "=" * 60)
    print("测试 2: 段落模式检索")
    print("=" * 60)
    print(f"查询: {TEST_QUERY}\n")
    print("⏳ 段落检索平均需约 17 秒，请稍候...")

    client = LitSenseClient(sleep_sec=1.0, timeout=120.0)
    results = client.search(query=TEST_QUERY, mode="passage", rerank=True)

    assert isinstance(results, list), f"返回类型应为 list，实际为 {type(results)}"
    print(f"✅ 返回 {len(results)} 条结果")

    if results:
        print(f"\n前3条结果：")
        for i, r in enumerate(results[:3], 1):
            print(f"  {i}. PMID: {r.get('pmid','')}  Score: {r.get('score',''):.4f}")
            text = r.get("text", r.get("passage", ""))
            print(f"     [{r.get('section','')}] {text[:120]}...")

        df = results_to_df({"passage": results}, "passage")
        assert not df.empty
        export_excel(df, OUT_EXCEL_P)
        export_json({"passage": results}, OUT_JSON_P)

        assert os.path.exists(OUT_EXCEL_P)
        assert os.path.exists(OUT_JSON_P)
        print(f"\n✅ 文件已生成: {OUT_EXCEL_P}")
    else:
        print("⚠️  无结果返回")

    return results


def test_section_filter():
    print("\n" + "=" * 60)
    print("测试 3: 章节过滤（仅 results 节）")
    print("=" * 60)

    client = LitSenseClient(sleep_sec=1.0, timeout=90.0)
    results = client.search(
        query="glycosylation site identification mass spectrometry plant",
        mode="sentence",
        rerank=True,
        section="results",
    )

    print(f"✅ results 节返回 {len(results)} 条")
    for r in results[:3]:
        section = r.get("section", "")
        print(f"   section={section}  pmid={r.get('pmid','')}")

    return results


def test_multiple_queries():
    print("\n" + "=" * 60)
    print("测试 4: 多查询批量检索")
    print("=" * 60)

    queries = [
        "O-glycosylation in plant cell wall proteins",
        "N-glycan biosynthesis in Arabidopsis thaliana",
    ]

    client = LitSenseClient(sleep_sec=2.0, timeout=90.0)
    all_results = client.search_multiple(queries=queries, mode="sentence")

    assert len(all_results) == len(queries)
    total = sum(len(v) for v in all_results.values())
    print(f"✅ {len(queries)} 条查询，共返回 {total} 条结果")

    df = results_to_df(all_results, "sentence")
    print(f"   合并后 DataFrame: {len(df)} 行")

    return all_results


if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("  NCBI LitSense 2.0 功能测试")
    print("=" * 60 + "\n")

    passed = 0
    failed = 0

    tests = [
        ("句子模式检索",   test_sentence_search),
        ("段落模式检索",   test_passage_search),
        ("章节过滤",       test_section_filter),
        ("多查询批量检索", test_multiple_queries),
    ]

    for name, fn in tests:
        try:
            fn()
            passed += 1
        except AssertionError as e:
            print(f"❌ 断言失败: {e}")
            failed += 1
        except Exception as e:
            print(f"❌ 测试失败: {type(e).__name__}: {e}")
            failed += 1

    print(f"\n{'=' * 60}")
    print(f"  测试结果: {passed} 通过 / {failed} 失败")
    print(f"{'=' * 60}\n")
    sys.exit(0 if failed == 0 else 1)
