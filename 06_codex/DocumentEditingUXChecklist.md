# Whale Editor 文稿编辑 UX 与 OpenXML 修改清单

更新时间：2026-05-02

## 产品定位

Whale Editor 不是浏览器里的完整 Word / PowerPoint 复制品。核心体验应该是：

1. 用户上传 DOCX / PPTX。
2. 系统把文件当作 OOXML / OpenXML 包读取，生成可定位 manifest。
3. 右侧对话把用户指令、批注、参考材料匹配成局部任务。
4. Agent 生成结构化 patch，底层修改 `word/document.xml`、`ppt/slides/slide*.xml` 等文件。
5. 左侧实时显示可信预览和每一步可见变化。
6. 完成后预览可看，真实下载走支付解锁。

## 文稿编辑网站需要注意什么

| 项目 | 风险 | Whale Editor 当前修改 |
| --- | --- | --- |
| 原文件安全 | 用户怕原稿被破坏 | 保持上传原件不直接覆盖，修改稿存在 session buffer，下载另存 |
| 预览真实性 | 浏览器预览容易被误认为 Word 级排版 | 左侧继续定位为 manifest preview，并在“结构”视图说明不是完整 WYSIWYG |
| 格式保真 | 全段重写会破坏 run、编号、表格 | 已走 OOXML patch，优先小范围替换既有文本节点 |
| 结构可定位 | AI 只看纯文本容易改错位置 | manifest 暴露 paragraphId、tableCellId、comment anchor、slideId、shapeId |
| 批注处理 | 批注可能没有精确锚点 | 批注视图列出锚定段落；无锚点时仍作为审阅上下文 |
| 审阅透明 | 用户不知道 AI 正在改哪里 | 右侧显示规划、步骤、操作明细；左侧高亮最近修改对象 |
| 失败可恢复 | patch 失败后不应直接丢任务 | 失败在步骤内展示，并保留重新读取和继续修改入口 |
| 下载权限 | 用户未付款不能拿到完整文件 | 下载接口返回 402，完成按钮为“解锁下载”，预览接口与下载二进制分离 |
| 文件隐私 | 文档可能含商业/论文内容 | 本地 API key 仅浏览器保存；后续还需接入文件保留清理 |
| 成本波动 | cache 命中不是用户可理解权益 | 报价按文档价值和修改窗口，不把缓存命中写成承诺 |

## UX 交互需要注意什么

| 交互点 | 应做法 | Whale Editor 当前修改 |
| --- | --- | --- |
| 上传入口 | 入口小，但拖拽区域明确 | 右侧只保留加号上传，左侧工作区也支持拖入 |
| 左右布局 | 文档优先，AI 侧栏固定 | 左侧编辑器 + 右侧悬浮侧栏已稳定 |
| 视图切换 | 编辑、批注、审阅、结构要可点可理解 | 新增“结构”视图，和编辑/批注/审阅同级 |
| 指令输入 | 支持自然语言、批注、参考材料 | 输入匹配层固定为 Office Input Matching Contract |
| 进程反馈 | 不只显示思考中，要显示可执行步骤 | 右侧显示规划、步骤、操作、缓存命中率 |
| 选区上下文 | 用户点某段应能带入指令 | 选中文稿可带入右侧输入 |
| 状态语言 | 中文站点不要混入无意义英文 | 前端已有中文兜底，manifest ID / 操作名保留原文 |
| 暗色模式 | 编辑区也必须跟随暗色 | 已修复左侧文稿、网格、工具栏、表格暗色 |
| 报价窗口 | 上传后不要无限期占用修改资格 | 上传后 10 分钟报价预览，付款后按套餐给修改窗口 |
| 完成后动作 | 预览、继续修改、下载要分开 | 完成后左侧刷新 session preview，下载单独锁住 |

## OpenXML 修改链路

| 阶段 | 实现要点 | 当前状态 |
| --- | --- | --- |
| 解包 | DOCX / PPTX 本质是 zip 包 | 已支持读取 DOCX / PPTX manifest |
| 抽取 | 读取 document.xml、styles、numbering、comments、slides、rels | 已抽取段落、表格、批注、PPT slide/shape/notes 摘要 |
| 固定 ID | 给每个可修改对象稳定 ID | 已使用 `docx.p.*`、`docx.tbl.*`、`pptx.slide.*`、`pptx.shape.*` |
| 规划 | 用户输入匹配 manifest ID | planner prompt 要求引用目标 ID 和格式保真约束 |
| Patch | 生成结构化操作而不是整文件文本 | 已有 `replace_text_in_paragraph`、`update_table_cell_text`、`replace_shape_text` 等路径 |
| 验证 | 检查 fileType、targetId、包结构 | patch validator 和 inspect warning/error 已接入 |
| 回写 | 更新包内 XML 后重新打包 | 完成后保存到 session buffer，并用 inspect 刷新预览 |

## DeepSeek 缓存命中设计

DeepSeek 官方 API 文档当前口径：

- Context Caching 默认开启，不需要额外改代码。
- 命中要求后续请求完整匹配已持久化的缓存前缀单元。
- 响应 usage 里提供 `prompt_cache_hit_tokens` 和 `prompt_cache_miss_tokens`。
- 缓存是 best-effort，不保证 100% 命中。
- 缓存不再使用后通常几小时到几天自动清理。
- 2026-05-02 官方价格页显示 `deepseek-v4-flash` 与 `deepseek-v4-pro` 区分 cache hit / cache miss 输入价格，cache hit 明显低于 cache miss。

参考来源：

- https://api-docs.deepseek.com/guides/kv_cache
- https://api-docs.deepseek.com/quick_start/pricing

Whale Editor 对应策略：

| 策略 | 做法 | 目的 |
| --- | --- | --- |
| 稳定 system prompt | system prompt、工具 schema、patch schema 不随用户输入变化 | 让高价值前缀可复用 |
| 稳定 manifest 排序 | 对 manifest 使用稳定排序和稳定 JSON | 减少同一文档追改时前缀漂移 |
| 动态信息后置 | 用户本轮指令、选区、错误修复放在后缀 | 保留前缀缓存命中机会 |
| 同文档 session | 继续修改复用最新 session 文档 | 用户连续追改时更容易热缓存 |
| 指标可见 | UI 展示 cache hit / miss token 和命中率 | 方便后续定价和成本复盘 |
| 不承诺 TTL | 报价不写“缓存一定保留” | 避免把 best-effort 能力变成用户权益 |

## 本轮已逐项修改

- 左侧工具栏新增 `结构` 视图。
- `结构` 视图展示 OpenXML 修改管线：解包、manifest、输入匹配、patch、验证回写。
- `结构` 视图展示当前文档的 DOCX/PPTX 结构指标。
- `结构` 视图展示参与修改的关键包内文件路径。
- `结构` 视图展示 DeepSeek 缓存设计：稳定前缀、动态后缀、同文档追改、定价口径。
- `结构` 视图展示 UX 护栏：预览边界、最小修改、可回溯、付费锁。
- 左侧底部状态条从 `OOXML Patch` 调整为更标准的 `OpenXML Patch`，并强调 manifest 稳定前缀。

## 后续还要补强

- 更完整的 Word/PPT 排版预览，尤其是页眉页脚、脚注、图片、PPT 坐标和主题字体。
- body child sequence，确保段落/表格交错顺序完全保真。
- run/span 级批注锚点，减少段落级锚点的误差。
- 跨 run 替换时的精确切分和格式分布。
- 原生 Word tracked-change markup；在实现前不能宣称支持真正修订模式。
- 后端强制 2 小时修改窗口、24 小时下载保留和文件清理。
- 真实 DeepSeek key + 真实 DOCX/PPTX 的端到端成功链路复测。
