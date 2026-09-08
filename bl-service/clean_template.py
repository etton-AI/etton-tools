"""
修复 ETTON 提单模板中「嵌套/重复合并域」（如「目的港」被做成 «目的港»«目的港»）。
重复域会导致 mailmerge 填充后 LibreOffice 渲染出 ".US..S." 之类重复字符。

dedupe_template(): 打开模板，把「段落内连续重复的同名 MERGEFIELD 域」去重，
                   只保留每个域第一次出现，写出新模板。幂等，可多次调用。
"""
import os, re, zipfile, shutil
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def qn(t): return f"{{{W}}}{t}"


def dedupe_template(template_path, fixed_path):
    """去重模板中的重复合并域，写出到 fixed_path。"""
    os.makedirs(os.path.dirname(fixed_path) or ".", exist_ok=True)

    with zipfile.ZipFile(template_path) as zin:
        orig_xml = zin.read("word/document.xml")
        infos = list(zin.infolist())

    root = etree.fromstring(orig_xml)
    for para in root.iter(qn("p")):
        seen = set()
        for run in list(para.findall(qn("r"))):
            instr = run.find(qn("instrText"))
            if instr is not None:
                m = re.search(r"MERGEFIELD\s+(\S+)", (instr.text or ""))
                if m:
                    name = m.group(1)
                    if name in seen:
                        para.remove(run)   # 移除重复域
                    else:
                        seen.add(name)
    new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    with zipfile.ZipFile(fixed_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in infos:
            data = orig_xml if item.filename == "word/document.xml" else None
            with zipfile.ZipFile(template_path) as z:
                data = z.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_xml
            zout.writestr(item, data)

    # 校验
    with zipfile.ZipFile(fixed_path) as z:
        z.read("word/document.xml")
    return fixed_path


if __name__ == "__main__":
    src = "/data/inputs/提单模板.docx"
    out = "/data/workspace/提单模板_已修复.docx"
    dedupe_template(src, out)

    from mailmerge import MailMerge
    doc = MailMerge(out)
    fields = sorted(set(doc.get_merge_fields()))
    print("修复后合并域:", fields)
    assert "目的港" in fields
    print("✅ 模板修复完成")
