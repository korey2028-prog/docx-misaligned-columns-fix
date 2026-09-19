---
name: docx-misaligned-columns-fix
description: 修复本地 Word/docx 中"伪对齐"错位排版——用连续制表符/空格硬凑的选项列、词汇表列等，改为 N 列无框线固定表格实现严格列对齐。典型场景：试卷完形填空 A/B/C/D 选项、单选选项、双语词汇表。走 tencent-local-office-edit 的 editor_sdk 通道编辑，保存后做 XML 级修补并核验。当用户要求"对齐选项""修复错位排版""选项一题一行上下对齐""Tab 对齐改表格"时使用。
agent_created: true
---

# docx 错位排版修复（伪对齐 → 无框线固定表格）

修复 docx 中靠连续 Tab/空格硬凑的"伪对齐"列：改为**无框线、固定列宽表格**，保证每行内容单行且各列上下严格对齐。三条铁律：**不重建整卷、不用空格/Tab 撑齐、未核验的文件不算完成**。

## 前置

- 先加载 `tencent-local-office-edit` skill（editor_sdk 编辑通道）。
- 再读本 skill 的 `references/playbook.md`——里面有完整命令序列、参数模板和踩坑细节。
- 用户有"限定区域修复"约束时：先样板后批量；只动目标区；从副本改，不覆盖原件。

## 核心流程（五阶段）

1. **只读侦察（XML，不动文件）**：解包 docx 用 lxml 定位目标区；确认①伪对齐手段（Tab/空格/自动编号）、②编号来源（numPr 自动编号则查 numbering.xml 的 start 值）、③所在节 pgSz/pgMar/cols（决定可用宽度）、④目标区起止元素索引。大文档不要用编辑器逐段扫描，XML 侦察快得多。
2. **列宽实测**：用 Pillow 按目标字体实测各列最长文本宽度；dxa 刻度 = font_size(pt)×20。总列宽 ≤ pgSz.w − pgMar.left − pgMar.right；优先等宽选项列；内边距（左 120/右 80 dxa）计入余量。
3. **编辑器写入**：present_files 呈现副本 → get_pool_status 拿 file_id → doc_find 逐段定位 → **从后往前** doc_delete_paragraph 删旧段 → doc_insert_table 插空表 → doc_set_table_cells 批量填内容（common_cell_properties 统一字体字号）→ doc_set_table_properties 设无边框/fixed/列宽/内边距 → save_file。
4. **保存后 XML 补丁**：跑 `scripts/post_save_patch.py`（自动备份）修两个 editor_sdk 已知副作用：①域 instrText 被拍平成可见文字；②表格行加 `<w:cantSplit/>` 防单行跨页拆分。
5. **核验**：跑 `scripts/verify_docx_diff.py <原件> <成品>` 做元素级全文档 diff，确认除目标区外零改动；逐行核对内容对应关系；再用 LibreOffice 渲染 PDF 做视觉验收（文字层/几何层/目检三级检查，见 playbook 阶段 5），并量化报告分页位移。确实无法渲染时才标注"视觉验收待确认"。

## 已知坑（必读）

- **instrText 拍平**：保存时域指令（HYPERLINK 等）可能变成可见文字，出现在非目标区——必须全文档扫描修复，并如实报告该回归。
- **cantSplit**：doc_set_table_properties 的 `cant_split=true` 语义反直觉（= 允许跨页断行）；工具无法设置"禁止拆分"，靠保存后 XML 补丁。
- **file_id 会变**：present_files 重新注册会换 UUID；报 "document is not open" 即实例过期，重新 get_pool_status 取新 id。
- **坐标失效**：任何写操作后旧 idx 立即失效；批量删除从后往前可免重查。
- **自动编号段落**：删 numPr 段落后自动题号消失，表格首列须显式写入原编号文本；不同编号实例的 start 值不同（未必从 1 开始）。
- **分节差异**：多节文档各节页边距/分栏不同，列宽必须按目标区所在节重算。
- **doc_find 跨 run**：编辑器查找能匹配跨 run 拆分的文本，返回的 paragraph_id 与源 XML 的 w14:paraId 一一对应，可用来双重校验定位是否正确。
