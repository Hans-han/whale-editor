# Whale Editor Web Workspace Status

更新时间：2026-04-30

## 已完成内容

- 将网站入口改成左侧文档编辑器工作区、右侧固定 AI 对话框的布局。
- 保留原有 DeepSeek API Key、文档上传、参考材料、流式执行、下载结果能力。
- 将执行进程放入右侧对话流：用户指令、规划、步骤执行、操作明细、缓存命中率、错误信息和最终动作在同一面板显示。
- 修复“继续修改”语义：前端会进入 follow-up 模式，后端在完成后把 session 缓存更新为最新生成文档，后续指令会复用当前文档。
- 按 Word 侧栏式文档助手的公开产品行为线索重写 executor system prompt：强调选区式最小编辑、样式继承、编号保真、表格几何保真、PPT 形状几何保真、模板填充和 review context 保留。
- 更新 planner prompt：规划阶段必须把 `styleId/listId/table geometry/slide layout` 等保真约束写入具体步骤。
- 新增固定输入匹配层 `prompts/office-input-matcher.md`：所有直接输入、PDF/文档参考、模板、批注、修订请求都会先按统一分类和优先级进入 planner。
- DOCX manifest 现在抽取批注文本、作者、日期和锚定段落 ID；DOCX 参考材料抽取文本时也会附带批注内容。
- DOCX manifest 已扩展到底层 OOXML 字段：run 字体/字号/粗体/斜体/下划线/颜色、段落编号、段落对齐/缩进/间距、表格宽度/grid columns、单元格宽度/合并信息。
- DOCX patch 的 `replace_text_in_paragraph` 已改为直接修改既有 `w:t` 文本节点，而不是重建整段；小范围替换会保留原 run 容器和 run 属性。
- 修复 `w:t` 首尾空格保真：被 patch 修改过的文本节点会写入 `xml:space="preserve"`，避免 Word 打开后吞掉边界空格。
- 修复 OOXML 布尔 run 属性解析：`w:val="off"` 和 `w:val="no"` 现在会识别为关闭，而不是误判成开启。
- 修复左侧工具栏“批注/审阅”不可点击的问题：三个工具栏按钮现在有真实切换状态、`aria-pressed` 和工作区视图。
- 新增本地 `/api/document/inspect` 接口：上传 DOCX/PPTX 后可读取 manifest 摘要、Word 批注、批注锚定段落、包验证 warning/error，用于左侧“批注/审阅”视图。
- 左侧“批注”视图能列出真实 Word 批注，并提供“按全部批注修改”快捷动作；“审阅”视图能显示段落/表格/批注/编号或 PPT slide/shape/notes 摘要，并把审阅修复指令填入右侧输入框。
- 修复中文文件名 mojibake：后端会把 multipart 上传中被 latin1 / Windows-1252 错解的 UTF-8 文件名还原为中文，`/api/document/inspect`、自动执行会话、参考材料名称和下载文件名都走同一规则。
- 修复执行步骤英文展示：planner prompt 已要求 `rationale/title/task` 使用简体中文；前端进度面板对常见英文 planning 文案做中文化兜底，保留 manifest ID 和真实文档英文词。
- 左侧“编辑”已改为实时文稿预览：上传后立即调用 inspect 渲染 DOCX 段落/表格或 PPTX slide/shape 摘要，不再只显示文件名状态页。
- 执行过程中，前端会根据 `step_done.operations` 直接更新左侧预览里的段落、单元格、幻灯片标题或形状文本，并高亮刚修改的对象。
- 右侧 AI 面板改成深海鲸鱼蓝文档侧栏；上传入口缩小为紧凑按钮，同时左侧工具栏也提供小型上传按钮。
- 修复执行结束时的 `流式读取失败：network error`：后端不再把完整修改稿以 base64 塞进 NDJSON 流，完成事件只返回 session 下载地址；前端收到完成事件后显示下载按钮，并且不会把连接尾部异常覆盖成失败。
- 新增 `/api/session/:id/download` 下载入口：修改后的 DOCX/PPTX 保存在当前 session buffer 中，下载时再按正确 MIME 和中文文件名返回。
- 已用浏览器模拟完整流式完成事件验证：右侧状态显示“全部步骤已完成”，错误框为空，下载按钮可见，左侧预览同步更新。
- 已完成一轮多 agent 网站评分，结果写入 `06_codex/WebsiteScoring.md`：综合约 73/100，下一阶段优先补选区联动、完成后重新预览、批注闭环、可访问性和生产可靠性护栏。
- 已按评分结果完成第一轮修复：选区上下文带入右侧输入、批注/审阅项定位到左侧正文、执行完成后自动切回编辑并重新 inspect 修改稿、移动端 390px 不越界、页面内错误提示和 ARIA 状态补齐、favicon/meta description 补齐。
- 已完成后端可靠性护栏：安全响应头、CORS allowlist、关闭 `X-Powered-By`、session 总数量/总字节上限、自动执行完成时直接用 Buffer 更新 session 并保留旧 base64 兼容。
- 修复本地自签名 HTTPS 可能打不开的问题：HSTS 仅在 `ENABLE_HSTS=1` 时启用；本地调试同时提供无证书的 HTTP 入口。
- 品牌与配色已统一：网页内不再出现第三方助手名称，右侧侧栏和主按钮统一为深海鲸鱼蓝视觉系统。
- 已重启本地服务：网站优先打开 `http://127.0.0.1:3001/`；需要 HTTPS/Office add-in 时使用 `https://127.0.0.1:3000/`。
- 已补齐全站鲸鱼蓝统一色板：背景、文稿区、右侧对话框、按钮、状态、错误提示和 favicon 都收敛到同一蓝色系统，不再混用暖色/绿色/紫红错误态。
- 已查证 DeepSeek V4 官方缓存口径：Context Caching 默认开启，未使用的缓存会在一段时间后自动清理，官方说明为“几小时到几天”，不是可承诺的固定 TTL。
- 已新增 `06_codex/PricingAndSessionRules.md`：固定 MVP 定价、上传后预览期、付款后修改窗口、下载保留时间、套餐边界和后续报价卡展示规则。
- 已把定价规则接入网页：上传文档后右侧会显示推荐套餐、价格、修改窗口、剩余轮次、下载保留时间和 10 分钟报价倒计时。
- 已新增外观切换：顶部可在“跟随系统 / 亮色 / 暗色”之间切换，选择会保存在本浏览器，暗色模式继续使用鲸鱼蓝色板。
- 已收缩左侧文稿框高度：左侧不再模拟很长的 Word 页面，改为紧凑工作区；用户也可以直接把文档拖到左侧区域载入。
- 已简化右侧上传入口：未上传时只显示一个加号按钮，上传后展开为文件名、大小和移除入口。
- 已接入下载支付锁：修改完成后前端显示“解锁下载”，后端 `/api/session/:id/download` 在未支付时返回 HTTP `402`，支付成功后才返回真实 DOCX/PPTX。
- 已接入 Stripe Checkout Session：`/api/checkout/session` 创建结算页，`/api/checkout/confirm` 和 `/api/stripe/webhook` 可把当前文档 session 标记为已支付。
- 已把修改稿预览和下载二进制分离：完成后左侧刷新走 `/api/session/:id/inspect`，不再用下载 URL 拉取文件，避免绕过支付锁。
- 已完成一轮视觉美化参考整理，新增 `06_codex/DesignReferences.md`，参考 Word Copilot、Grammarly、Notion AI 和 Canva Magic Write 的文档优先/侧栏助手/轻入口模式。
- 已美化网页外观：左侧文档区增加尺标、纸张边线、轻网格背景和更稳的文稿阴影；右侧侧栏改成更精致的深海蓝层级，首页说明文字收短。

## 当前采用的统一口径

- 网页是 Office 风格工作台，不再是单页上传表单。
- 左侧 `document-page` 是网页内编辑/状态工作区；真实文件修改仍由 OOXML patch 写回并下载。
- 左侧工具栏现在是三种工作区视图：编辑、批注、审阅；批注/审阅读取真实 Office manifest，但不直接在浏览器里改原文件。
- 编辑视图的“实时显示”是基于 manifest preview + patch operation 的可见文本同步；最终真实文件仍由 OOXML patch 写回并通过下载交付。
- 右侧 `ai-panel` 是核心交互面板，所有 agent 进程都以结构化状态展示。
- 继续修改优先复用 session 中的最新文档，不重新上传原始文件。
- 提示词层面不伪造任何第三方内部 system prompt；采用“Office-native formatting contract”来复刻可观察行为。
- 输入匹配层固定为 `Office Input Matching Contract v1`：直接指令最高优先级；批注在用户要求处理批注时变成局部任务；参考文件默认是支撑材料，除非用户指定为模板/规范/最终依据。
- “格式准”的主路径是低层 OOXML manifest + 结构化 patch：prompt 负责选目标和约束，真实修改由 `word/document.xml`、`word/comments.xml` 等包内文件完成。
- 小范围文本替换优先使用 `replace_text_in_paragraph`；整段重写、表格单元格替换、插入段落仍属于更高风险操作，需要依赖 manifest 约束和验证。
- UI 展示语言统一为中文；只有真实文档内容、manifest ID、style ID、操作名等需要精确引用的内容允许保留原文。
- 流式响应只传规划、步骤、操作和轻量完成状态；最终文件通过 session 下载接口获取，避免大文件 base64 让浏览器流读取在最后报网络错误。
- 执行完成后的默认体验是回到“编辑”预览，并尽量用 session 最新文档刷新左侧结果；下载按钮仍保留作为最终交付入口。
- 定价口径按冷启动/热缓存拆分：同一文档、同一会话内的连续修改可以按较高缓存命中率估算；跨天、换文档、长时间不再使用的任务必须按可能 cache miss 重新计价。
- MVP 默认主套餐为 `¥19.9 / 文档`，付款后给 `2 小时` 活跃修改窗口和 `24 小时` 下载保留；小文档用 `¥9.9`，复杂文档用 `¥39.9` 或 `¥69 起` 自动报价。
- 报价卡第一阶段已接入下载支付锁和 Stripe Checkout；权益扣减、2 小时修改窗口强制过期、24 小时下载保留清理仍未实现。
- 外观主题由前端 `localStorage` 保存；默认跟随系统，不影响文档 patch、上传、下载和会话复用逻辑。
- 左侧工作区只承担拖入、预览、选区和审阅定位，不作为完整长页面 WYSIWYG 编辑器。
- 当前视觉方向是“文档画布 + 深海蓝 AI 侧栏”，不是营销 landing page。

## 尚未解决的问题

- 左侧编辑区目前已能实时显示正文摘要和 patch 后的可见文本变化，但仍不是 Word 的完整 WYSIWYG 排版引擎。
- 文档内容仍不能直接在左侧所见即所得编辑后同步给 agent。
- 未在真实 DeepSeek API 成功路径下跑完整 DOCX/PPTX 修改，因为这需要有效 API key 和真实文档输入。
- 当前 patch layer 还不能生成原生 Word tracked-change markup；提示词已明确不能假装支持，只能保持操作明细可审计。
- 批注锚点目前能映射到段落级，尚未精确到 run/span 级选区。
- 跨多个 run 的替换目前会保留 run 容器并返回 warning，但替换文本的格式分布仍需要人工视觉复核；尚未实现 Word 级别的精确选区切分。
- manifest 和 patch 目前仍把段落、表格分别按集合读取；段落/表格交错顺序的精确建模还需要 preserveOrder 或 body child sequence 抽象。
- `replace_paragraph_text`、`update_table_cell_text`、`insert_paragraph_after` 仍比 substring patch 更容易影响内部 run/多段结构；后续应继续下沉到 run/cell paragraph 级。

## 下一阶段动作

- 接入更完整的真实 DOCX/PPTX 预览：在现有批注/审阅 manifest 基础上，把标题、段落、run 样式和 slide 摘要渲染得更接近原文。
- 用一份真实 `.docx` 和有效 DeepSeek key 跑完整成功路径，确认下载文件、缓存命中率和继续修改链路。
- 下一步继续提升“格式特别准”，应优先实现 body child sequence、run/span 级批注锚点、跨 run 精确切分、表格单元格多段保留、PPT 文本框字体与 autofit 属性。
- 下一步可以把输入匹配结果显式显示在右侧进度面板，例如“识别到 3 条 REVIEW_COMMENT、1 个 STYLE_SPEC”。
- 继续补上线付费闭环：把 Stripe Price/Product 配置、订单记录、权益扣减、退款/失败状态和 24 小时文件清理落到持久化存储中。
