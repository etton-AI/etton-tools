# -*- coding: utf-8 -*-
"""6月第4周 4 票验证：提取 + 箱数/国家校验 + 命名 vs 标准答案。"""
import os, sys, tempfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import core

INPUT = r"C:/Users/berry/OneDrive/Desktop/提单和保函转换测试文档/6月份/6月第4周"
ANSWER = r"C:/Users/berry/OneDrive/Desktop/提单和保函转换测试文档/拓锐6月报关退税资料【发客户】/6月第4周"

for folder in sorted(os.listdir(INPUT)):
    d = os.path.join(INPUT, folder)
    if not os.path.isdir(d):
        continue
    if folder.startswith("."):
        continue
    pdfs, packing = [], None
    for f in os.listdir(d):
        p = os.path.join(d, f)
        if f.startswith("~$") or f.startswith("."):
            continue
        if f.lower().endswith(".pdf"):
            pdfs.append(p)
        elif f.lower().endswith(".xlsx"):
            packing = p

    print("=" * 100)
    print("票文件夹:", folder)

    # 单份底单也走合并
    tmpdir = tempfile.mkdtemp()
    merged = os.path.join(tmpdir, "merged.pdf")
    core.merge_declaration_pdfs(pdfs, merged)
    import pdfplumber
    with pdfplumber.open(merged) as pdf:
        print("  底单页数:", len(pdf.pages), "| 原始 PDF 数:", len(pdfs))

    rec = core.extract_customs_data(merged)
    print("  提单号:", rec.get("提单号"), "| 柜号:", rec.get("柜号"),
          "| 船名航次:", repr(rec.get("船名航次")))
    print("  箱数:", rec.get("箱数"), "| 总重量:", rec.get("总重量"))
    print("  起运地:", rec.get("起运地"), "| 目的港:", rec.get("目的港"),
          "| 运抵国:", rec.get("运抵国"))

    packing_meta = None
    if packing:
        packing_meta = core.parse_packing_list(packing)
        print("  箱货清单: 箱数=", packing_meta.get("total_boxes"),
              "国家=", packing_meta.get("country"), "渠道=", packing_meta.get("channel"))
        print("           品名数=", len(packing_meta.get("products_en", [])),
              "总体积=", packing_meta.get("total_volume"))

    warns = core.validate_ticket(folder, rec, packing_meta)
    if warns:
        for w in warns:
            print("  ⚠️", w)
    else:
        print("  ✓ 箱数/国家一致")

    # 命名 vs 标准答案
    for kind, ext in (("提单", ".pdf"), ("电放保函", ".docx"), ("底单", ".pdf")):
        mine = core.derive_output_name(folder, kind) + ext
        ans_dir = os.path.join(ANSWER, folder)
        matched = mine in os.listdir(ans_dir) if os.path.isdir(ans_dir) else False
        print(f"  命名[{kind}]: {mine}  → 与标准答案一致={matched}")
