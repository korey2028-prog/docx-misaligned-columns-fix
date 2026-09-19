#!/usr/bin/env python3
"""verify_docx_diff.py — 两个 docx 的元素级全文档差异核验。

用法:
  python3 verify_docx_diff.py <original.docx> <modified.docx> [--context 80]

对正文 body 的顶层元素（段落/表格/其他）构建 (tag, 可见文本, instrText) 三元组序列，
用 SequenceMatcher 对齐并输出全部非 equal 差异块。用于"限定区域修复"验收：
合格标准 = 仅目标区出现 replace（旧段落 → 新表格），目标区外零差异。

注意：bookmark id 重编号、属性显式化（如 kinsoku 加 w:val="1"）属于编辑器往返的
可接受噪声，本脚本只比文本与 instrText，不比属性；如需属性级核验另做。
"""

import argparse
import sys
import zipfile
import difflib
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def load_body(path):
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    return root.find(f"{{{W}}}body")


def element_signature(el):
    tag = etree.QName(el).localname
    text = "".join(t.text or "" for t in el.iter(f"{{{W}}}t"))
    instr = "".join(x.text or "" for x in el.iter(f"{{{W}}}instrText")).strip()
    if tag == "tbl":
        rows = []
        for tr in el:
            if etree.QName(tr).localname == "tr":
                cells = [
                    "".join(t.text or "" for t in tc.iter(f"{{{W}}}t"))
                    for tc in tr if etree.QName(tc).localname == "tc"
                ]
                rows.append("|".join(cells))
        text = "\n".join(rows)
    return (tag, text, instr)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("original", type=Path)
    ap.add_argument("modified", type=Path)
    ap.add_argument("--context", type=int, default=100, help="每条差异预览的字符数")
    args = ap.parse_args()

    old = [element_signature(el) for el in load_body(args.original)]
    new = [element_signature(el) for el in load_body(args.modified)]
    print(f"elements: original={len(old)} modified={len(new)}")

    sm = difflib.SequenceMatcher(a=old, b=new, autojunk=False)
    non_equal = [op for op in sm.get_opcodes() if op[0] != "equal"]
    print(f"non-equal opcodes: {len(non_equal)}")
    for tag, i1, i2, j1, j2 in non_equal:
        print(f"== {tag} old[{i1}:{i2}] new[{j1}:{j2}]")
        for k in range(i1, i2):
            print(f"  OLD {k} {old[k][0]}: {old[k][1][:args.context]!r}"
                  + (f"  [instr: {old[k][2][:60]!r}]" if old[k][2] else ""))
        for k in range(j1, j2):
            print(f"  NEW {k} {new[k][0]}: {new[k][1][:args.context]!r}"
                  + (f"  [instr: {new[k][2][:60]!r}]" if new[k][2] else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
