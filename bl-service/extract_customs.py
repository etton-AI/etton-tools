"""
ETTON 报关底单字段提取器（基于精确坐标定位）

报关单第1页布局（pdfplumber 坐标）：
- 境内发货人 → y≈101, x≈40-130
- 境外收货人 → y≈125, x≈40-94
- 运输工具名称及航次号 → y≈125, x≈387-473
- 提运单号 → y≈125, x≈520-583
- 集装箱号 → 备注行 y≈241
- 指运港 → y≈170, x≈520-538
- 离境口岸 → y≈170, x≈641-668
- 件数 → y≈194, x≈258-267
- 毛重 → y≈194, x≈305-332
- 申报日期 → y≈101, x≈520-556
- 品名 → y≈299, x≈110-290
- 目的国 → y≈160, x≈442-465
"""

import re
import pdfplumber
from datetime import datetime


def extract_customs_data(pdf_path):
    """从报关底单PDF提取所有需要的字段"""
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words(x_tolerance=2, y_tolerance=2)

    # 按行分组（y坐标聚合）
    lines = _group_by_line(words)

    data = {}

    # ===== 1. 境内发货人（Shipper）=====
    # y≈101 行，x0 < 150 的文本
    data["shipper"] = _get_text_at(lines, y_target=101, x_max=150, y_tol=3)

    # ===== 2. 境外收货人（Consignee）=====
    # y≈125 行，x0 < 150
    data["consignee"] = _get_text_at(lines, y_target=125, x_max=150, y_tol=3)

    # ===== 3. 运输工具（船名航次）=====
    # y≈125 行，x 在 380-475 之间
    data["船名航次"] = _get_text_at(lines, y_target=125, x_min=380, x_max=475, y_tol=3)

    # ===== 4. 提运单号（提单号）=====
    # y≈125 行，x > 500
    data["提单号"] = _get_text_at(lines, y_target=125, x_min=500, x_max=600, y_tol=3)

    # ===== 5. 集装箱号（柜号）=====
    # 在备注行中搜索 "MATU" 开头的编号
    data["柜号"] = _find_in_notes(lines, pattern=r"MATU\d{7}")

    # 如果没找到，尝试全文本搜索
    if not data["柜号"]:
        full = " ".join(w["text"] for w in words)
        m = re.search(r"MATU\d{7}", full)
        if m:
            data["柜号"] = m.group(0)

    # ===== 6. 指运港（目的港）=====
    # y≈170 行，x 在 510-545 之间
    dest = _get_text_at(lines, y_target=170, x_min=510, x_max=545, y_tol=3)
    data["目的港"] = dest if dest else ""

    # ===== 7. 离境口岸（起运地）=====
    # y≈170 行，x 在 630-680 之间
    data["起运地"] = _get_text_at(lines, y_target=170, x_min=630, x_max=680, y_tol=3)

    # ===== 8. 件数（箱数）=====
    # y≈194 行，x 在 255-275 之间
    data["箱数"] = _get_text_at(lines, y_target=194, x_min=255, x_max=275, y_tol=3)

    # ===== 9. 毛重（总重量）=====
    # y≈194 行，x 在 300-340 之间
    data["总重量"] = _get_text_at(lines, y_target=194, x_min=300, x_max=340, y_tol=3)

    # ===== 10. 申报日期 =====
    # y≈101 行，x 在 510-560 之间
    date_str = _get_text_at(lines, y_target=101, x_min=510, x_max=560, y_tol=3)
    if date_str and re.match(r"\d{8}", date_str):
        data["起运日期"] = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
    else:
        data["起运日期"] = ""

    # ===== 11. 品名 =====
    # 商品明细区域：找到含 "牌" 或 "DICKIES" 等关键词的行
    data["品名"] = _find_product_name(lines)

    # ===== 清理字段 =====
    # 去掉括号内容（如 (913505023375763988)）
    for k in ["shipper", "consignee"]:
        data[k] = re.sub(r"[\(（][^)）]*[\)）]", "", data.get(k, "")).strip()

    # 箱数去小数
    if data.get("箱数"):
        try:
            data["箱数"] = str(int(float(data["箱数"])))
        except:
            pass

    # 目的港标准化
    if data.get("目的港") in ["美国", "USA"]:
        data["目的港"] = "The U.S."

    # ===== 保函相关字段 =====
    data["申请单位"] = data.get("shipper", "")
    data["目的地"] = data.get("目的港", "")
    data["运输工具"] = data.get("船名航次", "") or data.get("提单号", "")
    data["申请日期"] = datetime.now().strftime("%Y 年 %m 月 %d 日")

    return data


def _group_by_line(words):
    """将 words 按 y 坐标分组成行"""
    lines = {}
    for w in words:
        y_key = round(w["top"], 0)
        if y_key not in lines:
            lines[y_key] = []
        lines[y_key].append(w)
    # 每行按 x 排序
    for y in lines:
        lines[y].sort(key=lambda w: w["x0"])
    return lines


def _get_text_at(lines, y_target, x_min=None, x_max=None, y_tol=3):
    """在 y_target ± y_tol 范围内，筛选 x 在 [x_min, x_max] 之间的文本"""
    for y, words in lines.items():
        if abs(y - y_target) <= y_tol:
            texts = []
            for w in words:
                x0 = w["x0"]
                if x_min is not None and x0 < x_min:
                    continue
                if x_max is not None and x0 > x_max:
                    continue
                texts.append(w["text"])
            if texts:
                return " ".join(texts).strip()
    return ""


def _find_in_notes(lines, pattern):
    """在备注/唛码行中搜索 pattern"""
    for y, words in lines.items():
        if 235 <= y <= 250:
            full = " ".join(w["text"] for w in words)
            m = re.search(pattern, full)
            if m:
                return m.group(0)
    return ""


def _find_product_name(lines):
    """找到商品名称行（含品牌/规格信息）"""
    for y, words in lines.items():
        if 295 <= y <= 310:
            full = " ".join(w["text"] for w in words)
            # 格式: 3|0|背包|100%涤纶|DICKIES牌|BP582430IIQA
            parts = full.split("|")
            if len(parts) >= 4:
                # 取品名主体（前几个部分）
                name_parts = [p for p in parts[2:] if p and not re.match(r"^\d+(\.\d+)?$", p)]
                return " ".join(name_parts)
    return ""


if __name__ == "__main__":
    import sys
    pdf = sys.argv[1] if len(sys.argv) > 1 else "/data/inputs/QZKT26060001-24件报关底单.pdf"
    result = extract_customs_data(pdf)
    print("=" * 50)
    print("ETTON 报关底单 → 字段提取结果")
    print("=" * 50)
    for k, v in result.items():
        print(f"  {k}: {v}")
