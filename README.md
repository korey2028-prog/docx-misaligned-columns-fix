# docx-misaligned-columns-fix

一个 [WorkBuddy](https://www.workbuddy.cn) Skill：修复本地 Word/docx 中的**"伪对齐"错位排版**——把用连续制表符/空格硬凑的选项列，改成**无框线、固定列宽的表格**，实现每题一行、各列上下严格对齐。

## 解决什么问题

试卷类文档里常见这种排版：

```text
16.  A. congratulation⇥⇥⇥B. invitation⇥⇥⇥⇥C. rejection⇥⇥⇥⇥D. complaint
17.  A. taking up⇥⇥B. preparing for⇥⇥C. calling off⇥⇥D. passing down
```

连续 Tab（或空格）只能"碰巧"对齐：换一台机器、换一个字体渲染环境就错位。本 skill 把它改成五列无框线表格：

```text
16. | A. congratulation | B. invitation     | C. rejection      | D. complaint
17. | A. taking up      | B. preparing for  | C. calling off    | D. passing down
```

- 每题固定占一行，A/B/C/D 四列各自上下严格对齐；
- 列宽按目标字体（Pillow 实测）计算，保证**所有内容单行放下**，不截断、不压字号；
- 无边框、无底色、无表头，视觉上和正文融为一体；
- 原文内容零改动（逐项 diff 核验）。

## 安装

复制到 WorkBuddy 的用户级 skill 目录：

```bash
git clone https://github.com/korey2028-prog/docx-misaligned-columns-fix.git
mkdir -p ~/.workbuddy/skills
cp -r docx-misaligned-columns-fix ~/.workbuddy/skills/
```

依赖：`python3`（编辑通道由 WorkBuddy 内置的 `tencent-local-office-edit` skill 提供），列宽实测需要 `pip install pillow lxml`。

## 工作流程（五阶段）

1. **只读侦察（XML）** — 解包 docx 用 lxml 定位目标区：确认伪对齐手段（Tab/空格/自动编号）、题号来源（numPr 自动编号要查 numbering.xml 的 start 值）、所在节的页面宽度与边距。
2. **列宽实测** — 用 Pillow 按目标字体逐列实测最长文本（dxa 刻度 = pt×20），等宽选项列 + 内边距余量，总宽不超过正文可用宽度；放不下的块保留原样并列入异常清单，绝不硬塞。
3. **编辑器写入** — 走 editor_sdk 通道：`doc_find` 定位（paragraph_id 与源 XML 的 paraId 交叉核对）→ **从后往前**删旧段（前方坐标不漂移）→ 插表 → 批量填内容 → 设无边框/固定布局 → 保存。
4. **保存后 XML 补丁** — 修两个编辑器已知副作用（见下）。
5. **核验** — 元素级全文档 diff 确认目标区外零改动 + 逐行内容核对 + 诚实标注"视觉验收待确认"。

## 两个脚本

| 脚本 | 用途 |
|---|---|
| `scripts/post_save_patch.py` | 保存后修补：①把被编辑器拍平成可见文字的域指令（`HYPERLINK` 等）恢复为 `instrText`；②给目标表格所有行加 `<w:cantSplit/>`，防止单题行在页边界被拆成两页。写前自动备份。 |
| `scripts/verify_docx_diff.py` | 验收：对两个 docx 做元素级（tag + 可见文本 + instrText）SequenceMatcher diff，确认除目标区外零差异。 |

```bash
# 保存后修补（自动备份到 /tmp）
python3 scripts/post_save_patch.py 成品.docx --cantsplit-marker "A. congratulation"

# 与原件做全文档差异核验
python3 scripts/verify_docx_diff.py 原件.docx 成品.docx
```

## 已知坑（实测踩过）

- **编辑器会把域指令拍平**：editor_sdk 保存时可能把域的 `instrText` 变成可见文字（实测出现在离目标区很远的段落），必须全文档扫描恢复并如实报告。
- **cantSplit 设置不了**：编辑器表格工具的 `cant_split=true` 语义是"允许跨页断行"，想"禁止拆分"只能保存后 XML 补丁。
- **file_id 会过期**：预览面板重开会换 UUID，编辑器报 `document is not open` 就重新取。
- **写操作后坐标全部失效**：批量删除务必从后往前，可免去逐次重查。
- **自动编号题号随段落删除而消失**：表格首列需显式写回原编号。
- **多节文档各节页边距/分栏不同**：列宽必须按目标区所在节重算。

完整操作序列见 [references/playbook.md](references/playbook.md)。

## 适用范围

- 试卷完形填空 / 单选题的 A/B/C/D 选项区
- 双语词汇表、术语对照表等任何"Tab/空格硬凑列"的段落
- ⚠️ 仅处理**可编辑的 docx**；PDF/扫描件不在范围内

## License

[MIT](LICENSE)
