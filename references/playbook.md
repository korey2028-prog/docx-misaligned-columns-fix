# 错位排版修复 Playbook（完整操作序列）

以"试卷完形填空选项 → 5 列无框线表格"为完整实例。所有编辑通过 `tencent-local-office-edit` skill 的 edsdk.py 完成；XML 侦察/补丁用 lxml。

约定：`EDS = /Applications/WorkBuddy.app/Contents/Resources/app.asar.unpacked/resources/plugins/workbuddy-builtin/skills/tencent-local-office-edit/edsdk.py`，调用格式 `python3 EDS call <tool> --json '{...}'`。

---

## 阶段 0：保护原件

```bash
cp 原件.docx 原件_完形选项样板_v1.docx   # 一切编辑都在副本上做
```

## 阶段 1：只读侦察（XML）

解包（只读，不动原文件）：

```bash
mkdir -p /tmp/recon && cd /tmp/recon && unzip -o -q 副本.docx
```

用 lxml 提取顶层元素序列（`w:p` 取全部 `w:t` 拼接，`w:tbl` 记行数），据此定位：

1. **年份/章节标题** → 各卷起始元素索引；
2. **目标区标记**（如"完形填空"）→ 选项段区间；
3. 逐段打印区间内容，确认伪对齐手段：统计每段的 `w:tab` 个数、连续空格、`w:br`。

### 关键检查项

| 检查项 | 方法 | 为什么 |
|---|---|---|
| 题号来源 | 段落 `pPr/numPr/numId` → numbering.xml 找 `w:num` → abstractNum → `w:start` | 题号是渲染出来的，段落文本里没有；删段=删题号，表格首列要显式补 |
| 段内序列 | 遍历 run：`w:t`/`w:tab`/`w:br` 按顺序拼接 | 确认文本真实内容与拆分方式（option 文本可能被拆成多个 run，如 "c"+"ongratulation"） |
| 可用宽度 | 目标区**之后**最近的 `pPr/sectPr`（段落级分节符）或 body 末尾 sectPr：`pgSz@w − pgMar@left − pgMar@right` | 多节文档各节不同；注意分栏（cols num>1）时按栏宽算 |
| 自定义制表位/缩进 | `pPr/tabs`、`pPr/ind` | 决定旧对齐的原理，也影响表格替换后的缩进预期 |
| 字体字号 | 段落 rPr：`rFonts@ascii`、`sz`（半磅，24=12pt） | 表格单元格要显式复刻，否则吃默认样式 |

## 阶段 2：列宽实测（Pillow）

```python
from PIL import ImageFont
f = ImageFont.truetype("/System/Library/Fonts/Supplemental/Times New Roman.ttf", 12*20)  # 12pt，dxa 刻度
w = f.getlength("A. congratulation")   # 返回 dxa
```

- 对每列取**最长文本**（含 "A. " 前缀）实测；macOS 字体目录：`/System/Library/Fonts/Supplemental/`。
- 列宽 = max 文本宽 + 内边距（左 120 + 右 80 dxa）；所有列求和 ≤ 可用宽度；放不下时按指令在总宽内压缩等宽列，仍放不下 → 该块保留原样并列入异常清单，**不硬塞**。
- 优先等宽选项列（同一选项块内 A/B/C/D 列等宽），题号列单独窄列。

## 阶段 3：编辑器写入

```text
1. present_files(副本路径)                    # 必须先呈现；open_file 不会出现在预览
2. edsdk call get_pool_status --json '{"file_path":"<同上路径>"}'   # 取 UUID file_id
3. 逐段 doc_find（唯一性子串）拿每段 begin/paragraph_id
   → 与源 XML 的 w14:paraId 交叉核对，防定位错段
4. doc_delete_paragraph：按 begin **从大到小**逐个删   # 从后往前，前方坐标不漂移，无需重查
5. doc_insert_table idx=<首段原begin> row_count=N col_count=5
   → 返回 table_info.id 即 table_id（8 字符）
6. doc_set_table_cells：一次传全部 N×5 个 cell 的 text，
   common_cell_properties.text_format = {font_family, font_size}（pt）
7. doc_set_table_properties：
   mode=manual, col_widths_dxa=[...], width={type:dxa,value:总宽},
   alignment=left, cell_v_align=top, row_height_auto=true,
   cell_margin={left:120,right:80,top:0,bottom:0},
   borders={top/bottom/left/right/inside_h/inside_v 全部 {val:"none"}}
8. edsdk call save_file --json '{"file_id":"..."}'
```

注意：

- `doc_set_table_cells` 的 cells 元素：`{row, col, text}`（1-based）；不传 text=保留，传空串=清空。
- 新插表格默认样式是 **Table Grid（带边框）**，第 7 步必须覆盖为 none。
- 写操作会让旧 idx 失效；本序列中删除全部完成后再插表，插表 idx 用删除前记录的首段 begin 仍有效。
- 编辑器中途报 `document is not open` = 实例过期（预览重开/池清理），重新 present_files + get_pool_status。

## 阶段 4：保存后 XML 补丁（两个已知副作用）

跑 `scripts/post_save_patch.py <成品.docx> --cantsplit-marker "<表格特征文本>"`，或手动：

1. **instrText 拍平修复**：全文档扫 `w:t`，凡文本以域指令关键词开头（`HYPERLINK`、`PAGEREF`、`TOC`、`SEQ`、`MERGEFIELD`…）且所在 run 处于 `fldChar begin` 与 `separate/end` 之间 → 把 `w:t` 改回 `w:instrText`（保留 rPr 与 xml:space="preserve"）。实测 editor_sdk 在本项目把 `HYPERLINK "javascript:;" \o "机器发音"` 拍平过一次，出现在**远离目标区**的词汇区——非目标区回归必须修复并报告。
2. **cantSplit**：给目标表格每一行 `trPr` 插入 `<w:cantSplit/>`（工具的 `cant_split=true` 反而是允许拆分，无法靠工具设置"禁止"）。
3. 补丁前自动备份；补丁后必须重跑阶段 5。

## 阶段 5：核验

1. **全文档 diff**：`scripts/verify_docx_diff.py 原件.docx 成品.docx`——元素级 (tag, 文本, instrText) 序列 SequenceMatcher。合格标准：仅目标区出现 replace（旧段→表格）；目标区外零差异（bookmark id 重编号/属性规范化属可接受噪声，需说明）。
2. **内容逐项核对**：表格每行 5 格与原段文本一一对应；编号连续且等于原自动编号值。
3. **样式核对**：抽查单元格 run 的 rFonts/sz 与原文一致。
4. **视觉验收**：无法渲染页面时，明确标注"尚未完成视觉验收"，请用户在预览中翻查目标区及**后续页面**（行高变化会移动后文分页位置）。
5. 删除/关闭编辑器实例避免过期副本覆盖补丁（close_file 报 no editor open 即已清理）。

## 批量阶段的追加规则

- 全部目标块先做**只读勘察清单**：文件名、目标区位置、版式组、是否本来合规、拟用列宽。用户确认后再动手。
- 相同版式复用同一组参数；新版权式重新走阶段 1–2。
- 每份文件独立走阶段 3–5，不得以抽查代替逐份核验。
- 输出文件另存不覆盖；失败/待审文件单独标记。
