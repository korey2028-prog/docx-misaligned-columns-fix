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

## 安装与使用

### 方式一：WorkBuddy 用户（最省事）

把本仓库克隆到 `~/.workbuddy/skills/docx-misaligned-columns-fix/`，重启会话即可。之后只要对 WorkBuddy 说"帮我把这份 docx 里的完形选项对齐"，skill 会自动触发，编辑、补丁、核验全流程自动走。

### 方式二：其他 AI Agent（豆包 / Codex / 任何能操作本地文件的 Agent）

需要：`python3` + `pip install lxml pillow`（唯一两个第三方依赖）。

把仓库链接发给你的 Agent，并附上这段话：

> 请阅读这个仓库的 SKILL.md 和 references/playbook.md，按照里面的方法论处理我给你的 docx：
> 把"Tab/空格硬凑对齐"的选项列改成无框线固定表格。要求：① 先做 XML 只读侦察（表格方案不依赖制表位，微信预览也能对齐）；② 列宽用 Pillow 按文档实际字体实测，不许估算；③ 编辑完成后运行 scripts/post_save_patch.py 打补丁（域指令恢复 + cantSplit + 单元格左对齐），再运行 scripts/verify_docx_diff.py 与原件做全文档 diff 核验；④ 全程从副本改，不覆盖原件；⑤ 核验不通过的文件不许标记完成。

两个脚本脱离 WorkBuddy 也能独立使用：

```bash
# 保存后修补：域指令恢复 + 行禁止跨页拆分 + 单元格强制左对齐（自动备份）
python3 scripts/post_save_patch.py 成品.docx \
    --cantsplit-marker "A. congratulation" --left-align-marker "A. congratulation"

# 与原件做元素级全文档差异核验（验收标准：除目标区外零差异）
python3 scripts/verify_docx_diff.py 原件.docx 成品.docx
```

### 方式三：不用 AI，手动改

方法论照旧适用：无框线表格 + 固定列宽 + 按真实字体测宽。改完用 `verify_docx_diff.py` 自查一遍再交。

## 已知坑（实测踩过）

- **微信 docx 预览不支持制表位**：Tab/制表位方案在 Word/WPS/LibreOffice 里都对，微信预览会把 tab 当普通空格——**列对齐只有无框线表格这一条可靠路径**。
- **默认段落样式可能是两端对齐（jc=both）**：表格单元格继承它后，长选项折行时首行被拉伸（单行看不出来，潜伏雷）。用 `--left-align-marker` 强制左对齐。
- **编辑器会把域指令拍平**：editor_sdk 保存时可能把域的 `instrText` 变成可见文字（实测出现在离目标区很远的段落），必须全文档扫描恢复并如实报告。
- **cantSplit 设置不了**：编辑器表格工具的 `cant_split=true` 语义是"允许跨页断行"，想"禁止拆分"只能保存后 XML 补丁。
- **file_id 会过期**：预览面板重开会换 UUID，编辑器报 `document is not open` 就重新取。
- **写操作后坐标全部失效**：批量删除务必从后往前，可免去逐次重查。
- **自动编号题号随段落删除而消失**：表格首列需显式写回原编号。
- **多节文档各节页边距/分栏不同**：列宽必须按目标区所在节重算。
- **核验假阳性**：多套卷子并存时题号重复（多份卷都有"23. A. …"），按题号匹配会误报——核验要带上下文或按唯一选项文本定位。
- **WPS 特有属性 `firstLineChars`**：Word 与 LibreOffice 解释不同（差约 20pt），跨平台 tab 起点漂移的元凶之一；表格方案天然免疫。

完整操作序列见 [references/playbook.md](references/playbook.md)。

## 适用范围

- 试卷完形填空 / 单选题的 A/B/C/D 选项区
- 双语词汇表、术语对照表等任何"Tab/空格硬凑列"的段落
- ⚠️ 仅处理**可编辑的 docx**；PDF/扫描件不在范围内

## License

[MIT](LICENSE)
