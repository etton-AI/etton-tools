"""
端到端验证：模拟"审核流程"（不启动 Flask，直接调用 core 引擎）
① 报关底单PDF → 提取字段
② 生成审核用Excel（模拟人工可编辑的中间产物）
③ 从审核Excel → 生成 提单+保函 → ZIP
"""
import os, sys, tempfile, zipfile
sys.path.insert(0, os.path.dirname(__file__))

from core import (
    extract_customs_data, build_review_excel, parse_review_excel,
    generate_from_review_xlsx, fill_telex_docx,
)

PDF = "/data/inputs/QZKT26060001-24件报关底单.pdf"
TPL = "/data/inputs/提单模板.docx"
base = tempfile.mkdtemp()

print("=" * 56)
print("① 报关底单 → 提取字段")
print("=" * 56)
rec = extract_customs_data(PDF)
for k, v in rec.items():
    print(f"  {k}: {v}")

print("\n" + "=" * 56)
print("② 生成「审核用 Excel」（供前端展示/人工编辑）")
print("=" * 56)
xlsx = os.path.join(base, "审核表.xlsx")
build_review_excel([rec], xlsx)
print("审核表:", xlsx)

# 模拟人工审核：修改/补全字段（例如补总体积、改目的港）
records = parse_review_excel(xlsx)
records[0]["总体积"] = "12.5"          # 人工补填
records[0]["目的港"] = "The U.S."      # 人工确认
print("审核后(模拟编辑):", records[0]["总体积"], records[0]["目的港"])

print("\n" + "=" * 56)
print("③ 确认生成 → 提单 + 电放保函 → ZIP")
print("=" * 56)
out_dir = os.path.join(base, "out")
# 重新写回审核xlsx（模拟前端回传）
build_review_excel(records, xlsx)
zip_path, bl_cnt, telex_cnt = generate_from_review_xlsx(xlsx, TPL, out_dir, to_pdf=False)
print(f"生成完成: {zip_path}")
print(f"  提单 {bl_cnt} 份, 保函 {telex_cnt} 份")

print("\n" + "=" * 56)
print("产物清单:")
print("=" * 56)
for f in sorted(os.listdir(out_dir)):
    print("  -", f)

# 校验保函结构
tl = os.path.join(base, "telex_check.docx")
fill_telex_docx(tl, records[0])
from docx import Document
doc = Document(tl)
txt = "\n".join(p.text for p in doc.paragraphs)
print("\n电放保函关键字段校验:")
for key in ["广州拓锐科技有限公司", "ETTON TECHNOLOGY LOGISTICS (ZHONGSHAN)", rec["提单号"], rec["运输工具"], rec["目的地"], rec.get("consignee","")]:
    ok = key in txt
    print(f"  [{'✓' if ok else '✗'}] {key}")

print("\n🎉 审核流程端到端验证通过" if bl_cnt and telex_cnt else "⚠️ 异常")
