# -*- coding: utf-8 -*-
"""批量工作流端到端自测：拆分底单合并 + 提取 + 校验 + 批量生成 ZIP。"""
import os, sys, tempfile, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import core

BASE = r"C:/Users/berry/OneDrive/Desktop/提单和保函转换测试文档/7月份第1周/7月份第1周"
TPL = r"C:/Users/berry/OneDrive/Desktop/提单和保函转换测试文档/提单模板.docx"

tickets = []
for folder in sorted(os.listdir(BASE)):
    d = os.path.join(BASE, folder)
    if not os.path.isdir(d):
        continue
    pdfs, packing = [], None
    for f in os.listdir(d):
        p = os.path.join(d, f)
        if f.lower().endswith(".pdf"):
            pdfs.append(p)
        elif f.lower().endswith(".xlsx"):
            packing = p
    print("=" * 90)
    print("票文件夹:", folder)
    print("  PDF 数:", len(pdfs), [os.path.basename(p) for p in pdfs])

    tmpdir = tempfile.mkdtemp()
    merged = os.path.join(tmpdir, "merged.pdf")
    core.merge_declaration_pdfs(pdfs, merged)
    import pdfplumber
    with pdfplumber.open(merged) as pdf:
        print("  合并后页数:", len(pdf.pages))

    rec = core.extract_customs_data(merged)
    print("  提单号:", rec.get("提单号"), "| 箱数:", rec.get("箱数"),
          "| 总重量:", rec.get("总重量"))
    print("  起运地:", rec.get("起运地"), "| 目的港:", rec.get("目的港"),
          "| 运抵国:", rec.get("运抵国"))
    print("  运输工具:", repr(rec.get("运输工具")), "| 柜号:", rec.get("柜号"))

    packing_meta = None
    if packing:
        packing_meta = core.parse_packing_list(packing)
        print("  箱货清单: 箱数=", packing_meta.get("total_boxes"),
              "国家=", packing_meta.get("country"), "渠道=", packing_meta.get("channel"))
        print("           品名数=", len(packing_meta.get("products_en", [])),
              "总体积=", packing_meta.get("total_volume"))

    warns = core.validate_ticket(folder, rec, packing_meta)
    if warns:
        print("  ⚠️ 校验提醒:")
        for w in warns:
            print("     -", w)
    else:
        print("  ✓ 箱数/国家一致")

    print("  命名示例: 提单 →", core.derive_output_name(folder, "提单"))
    print("           底单 →", core.derive_output_name(folder, "底单"))
    tickets.append({"folder": folder, "record": rec, "merged_pdf": merged})

print("=" * 90)
print("测试 generate_batch（前 2 票，含 1 票拆分底单）...")
out_dir = os.path.join(tempfile.mkdtemp(), "out")
zip_path, results = core.generate_batch(tickets[:2], TPL, out_dir)
print("ZIP:", zip_path)
with zipfile.ZipFile(zip_path) as zf:
    for n in zf.namelist():
        print("  ZIP 内含:", n)
for r in results:
    print("  bl   :", os.path.basename(r["bl_pdf"]), "存在=", os.path.exists(r["bl_pdf"]))
    print("  telex:", os.path.basename(r["telex_docx"]), "存在=", os.path.exists(r["telex_docx"]))
    print("  ddan :", os.path.basename(r["ddan_pdf"]), "存在=", os.path.exists(r["ddan_pdf"]))
    print("  预览 :", os.path.basename(r["telex_preview_pdf"]), "存在=", os.path.exists(r["telex_preview_pdf"]))
print("DONE")
