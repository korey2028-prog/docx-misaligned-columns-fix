#!/usr/bin/env python3
"""post_save_patch.py — editor_sdk 保存 docx 后的已知副作用修补。

用法:
  python3 post_save_patch.py <file.docx> --cantsplit-marker "<表格内特征文本>"
      [--backup-dir /tmp] [--field-keywords HYPERLINK,PAGEREF,TOC,SEQ,MERGEFIELD]

自动完成（写前先备份原文件）:
  1. fix-instrtext : 把被 editor_sdk 拍平成可见 w:t 的域指令恢复为 w:instrText
     （仅当该 run 位于 fldChar begin 与 separate/end 之间才改，避免误伤普通文本）。
  2. add-cantsplit : 给包含 --cantsplit-marker 文本的表格的所有行加 <w:cantSplit/>，
     防止单题行在页边界被拆成两页。

退出码: 0=成功, 2=未找到 cantsplit 目标表格, 其他=异常。
"""

import argparse
import shutil
import sys
import time
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
DEFAULT_FIELD_KEYWORDS = (
    "HYPERLINK", "PAGEREF", "TOC", "SEQ", "MERGEFIELD", "INCLUDEPICTURE",
    "REF ", "DATE", "TIME", "PAGE ", "IF ", "SYMBOL",
)


def load_document(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}
    root = etree.fromstring(data["word/document.xml"])
    return names, data, root


def save_document(path, names, data, root):
    data["word/document.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, data[n])


def fix_instrtext(root, keywords):
    """把域结构内被拍平成 w:t 的指令文本恢复为 instrText。"""
    fixed = 0
    for r in root.iter(f"{{{W}}}r"):
        texts = r.findall(f"{{{W}}}t")
        if not texts:
            continue
        content = "".join(t.text or "" for t in texts)
        if not any(content.startswith(k) for k in keywords):
            continue
        # 确认 run 处于域结构内：向前找同级/邻近的 fldChar begin，向后找 separate/end
        parent = r.getparent()
        runs = [el for el in parent.iter(f"{{{W}}}r")]
        try:
            pos = runs.index(r)
        except ValueError:
            continue
        before = runs[max(0, pos - 12):pos]
        after = runs[pos + 1:pos + 12]
        has_begin = any(x.find(f"{{{W}}}fldChar") is not None
                        and x.find(f"{{{W}}}fldChar").get(f"{{{W}}}fldCharType") == "begin"
                        for x in before)
        has_sep_or_end = any(x.find(f"{{{W}}}fldChar") is not None
                             and x.find(f"{{{W}}}fldChar").get(f"{{{W}}}fldCharType") in ("separate", "end")
                             for x in after)
        if not (has_begin and has_sep_or_end):
            continue
        t = texts[0]
        instr = etree.Element(f"{{{W}}}instrText")
        instr.set(XML_SPACE, "preserve")
        instr.text = t.text
        idx = list(r).index(t)
        r.remove(t)
        r.insert(idx, instr)
        fixed += 1
    return fixed


def add_cantsplit(root, marker):
    """给包含 marker 文本的表格所有行加 cantSplit；返回 (表数, 行数)。"""
    def text_of(el):
        return "".join(t.text or "" for t in el.iter(f"{{{W}}}t"))

    tables = rows = 0
    for tbl in root.iter(f"{{{W}}}tbl"):
        if marker not in text_of(tbl):
            continue
        tables += 1
        for tr in tbl:
            if etree.QName(tr).localname != "tr":
                continue
            trpr = tr.find(f"{{{W}}}trPr")
            if trpr is None:
                trpr = etree.Element(f"{{{W}}}trPr")
                tr.insert(0, trpr)
            if trpr.find(f"{{{W}}}cantSplit") is None:
                trpr.insert(0, etree.Element(f"{{{W}}}cantSplit"))
                rows += 1
    return tables, rows


def left_align_table(root, marker):
    """把包含 marker 文本的表格内所有段落强制 jc=left。

    背景：很多试卷文档的默认段落样式是两端对齐（jc=both）。表格单元格
    段落若继承两端对齐，长选项在格内折行时首行会被拉伸出难看的大空隙。
    单行单元格看不出问题，属于潜伏雷。
    """
    def text_of(el):
        return "".join(t.text or "" for t in el.iter(f"{{{W}}}t"))

    tables = paras = 0
    for tbl in root.iter(f"{{{W}}}tbl"):
        if marker not in text_of(tbl):
            continue
        tables += 1
        for p in tbl.iter(f"{{{W}}}p"):
            ppr = p.find(f"{{{W}}}pPr")
            if ppr is None:
                ppr = etree.Element(f"{{{W}}}pPr")
                p.insert(0, ppr)
            jc = ppr.find(f"{{{W}}}jc")
            if jc is None:
                jc = etree.Element(f"{{{W}}}jc")
                # w:jc 在 pPr 中的合法位置靠后，直接 append 即可（schema 为 choice 序列）
                ppr.append(jc)
            if jc.get(f"{{{W}}}val") != "left":
                jc.set(f"{{{W}}}val", "left")
                paras += 1
    return tables, paras


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("docx", type=Path)
    ap.add_argument("--cantsplit-marker", default=None,
                    help="目标表格内出现的特征文本（用于定位表格）")
    ap.add_argument("--left-align-marker", default=None,
                    help="把含该特征文本的表格内所有段落强制 jc=left")
    ap.add_argument("--field-keywords", default=",".join(k.strip() for k in DEFAULT_FIELD_KEYWORDS),
                    help="逗号分隔的域指令关键词")
    ap.add_argument("--backup-dir", default="/tmp")
    args = ap.parse_args()

    if not args.docx.exists():
        print(f"file not found: {args.docx}", file=sys.stderr)
        return 1

    backup = Path(args.backup_dir) / f"{args.docx.stem}.patch-bak-{int(time.time())}{args.docx.suffix}"
    shutil.copy(args.docx, backup)

    names, data, root = load_document(args.docx)

    n_instr = fix_instrtext(root, [k for k in args.field_keywords.split(",") if k])
    n_tables = n_rows = 0
    if args.cantsplit_marker:
        n_tables, n_rows = add_cantsplit(root, args.cantsplit_marker)
    la_tables = la_paras = 0
    if args.left_align_marker:
        la_tables, la_paras = left_align_table(root, args.left_align_marker)

    save_document(args.docx, names, data, root)
    print(f"backup: {backup}")
    print(f"instrText restored: {n_instr}")
    print(f"cantsplit: tables={n_tables} rows={n_rows}")
    print(f"left-align: tables={la_tables} paragraphs={la_paras}")
    if args.cantsplit_marker and n_tables == 0:
        print(f"WARN: no table matched marker {args.cantsplit_marker!r}")
        return 2
    if args.left_align_marker and la_tables == 0:
        print(f"WARN: no table matched marker {args.left_align_marker!r}")
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
