# Whale Editor 网站评分与功能规划

更新时间：2026-05-02

## 评分标准来源

- Lighthouse / Chrome DevTools：用于 Performance、Best Practices、SEO 的 0-100 评分框架。参考口径：90-100 为良好，50-89 需要改进，0-49 较差；Performance 由 FCP、Speed Index、LCP、TBT、CLS 等指标加权。
- WCAG 2.2：用于可访问性底线。重点看可感知、可操作、可理解、健壮，以及 AA 层级中与焦点、目标尺寸、颜色对比、状态反馈相关的规则。
- Nielsen Norman Group 10 Usability Heuristics：用于产品体验判断。重点看系统状态可见性、贴近用户语言、一致性、错误预防、错误恢复、减少记忆负担、保持界面聚焦。

## 本轮评分

| 维度 | 分数 | 结论 |
| --- | ---: | --- |
| 性能与工程质量 | 78/100 | 首屏轻，测试通过；主要扣分在安全头、CORS、favicon 404、session 内存边界和 base64 中转风险。 |
| 可访问性与中文本地化 | 66/100 | 基础可用，但移动端裁切、编辑区语义、状态/错误语义、颜色对比和英文残留还不够。 |
| 产品体验 / Word 侧栏助手接近度 | 76/100 | 左文稿右 AI 的核心形态已成型；缺“选区联动”和“完成后重新预览修改稿”。 |
| 综合判断 | 73/100 | 当前是可演示的原型，但离稳定可用的文档工作台还差一轮质量和交互闭环。 |

## 实测证据

- 本地入口：`https://127.0.0.1:3000/`。
- 日常网站入口：`http://127.0.0.1:3001/`；Office add-in / HTTPS 入口：`https://127.0.0.1:3000/`。
- Chrome headless 页面实测：首屏总传输约 88 KB，`index.html` 8.1 KB、`app.css` 29.5 KB、`app.js` 50.3 KB。
- 上传测试 DOCX 后：左侧能显示文稿预览；批注视图能读取真实 Word 批注；审阅视图能列出段落、批注和包检查。
- 桌面 Tab 顺序能到达 API Key、工具栏、上传、文档区、dropzone、指令框、添加材料。
- 移动端 390px 下无页面级横向滚动，但工具栏和文档页存在局部裁切风险。
- `/api/health` 返回服务可用，但 `libreoffice.available=false`，所以 `.doc/.ppt` 转换不能作为当前主卖点。
- 视觉回归：全站色彩已统一为鲸鱼蓝，右侧侧栏、主按钮、状态、错误提示和 favicon 不再混用暖色/绿色/紫红色。

## 必须加的 Feature

1. 选区联动
   - 左侧选中段落、表格单元、批注锚点后，右侧自动带入目标 ID、原文、样式、批注上下文。
   - 这是最接近 Word 侧栏式文档助手体感的核心功能，比继续堆按钮更重要。

2. 完成后重新预览修改稿
   - 执行完成后自动 inspect session 里的最新文件。
   - 左侧展示修改后的预览，再显示下载入口，减少“下载才知道效果”的不确定性。

3. 批注工作流闭环
   - 批注列表点击后定位左侧段落。
   - 执行后显示批注处理状态：已处理、未处理、需人工复核。

4. 移动端和可访问性修复
   - 工具栏在 390px 下必须可见可点。
   - `document-page` 补充明确可访问名称和编辑/预览语义。
   - 状态区使用 `role=status` / `aria-live`，错误区使用 `role=alert`。
   - 颜色对比和目标尺寸按 WCAG 2.2 AA 修正。

5. 生产可靠性护栏
   - 加安全响应头、收紧 CORS、关闭 `X-Powered-By`。
   - 修复 favicon 404、增加 meta description。
   - session 缓存加总数量和总字节上限。
   - 去掉最终文件 base64 中转，直接把 Buffer 或临时文件交给 session。

6. 页面质量监控脚本
   - 固定检查资源预算、控制台错误、移动端裁切、键盘 Tab、NDJSON 是否收到终端事件。
   - 每次改 UI 后用同一套脚本回归。

## 暂缓或不加的 Feature

1. 暂缓完整 Word WYSIWYG 编辑器
   - 成本高，且容易和 OOXML patch 主路径冲突。
   - 当前更应该做“可信预览 + 选区上下文 + patch 回写”。

2. 暂缓原生 Word tracked changes
   - 目前还没有稳定的 Word 级修订标记生成链路。
   - 先用可审计操作记录和修改后预览替代。

3. 暂缓重型前端框架迁移
   - 当前静态页面体积小、性能好。
   - 迁移不会解决当前最关键的文档交互问题。

4. 暂缓更多动画和视觉装饰
   - 当前扣分不在视觉炫酷度，而在状态语义、移动端、可访问性和可靠性。

5. 暂缓把 `.doc/.ppt` 当主功能宣传
   - 当前 LibreOffice 不可用，旧格式转换路径不稳定。
   - 先把 `.docx/.pptx` 做准。

## 下一阶段实施顺序

1. 先修低成本质量项：favicon、meta description、安全头、CORS、移动端裁切、ARIA 状态、错误区。
2. 实现左侧选区联动，把选中的 paragraph/table/comment context 填入右侧输入。
3. 执行完成后自动重新 inspect 最新 session 文件，并刷新左侧预览。
4. 批注列表和左侧预览做互相定位，增加处理状态。
5. 移除 base64 中转，给 session cache 加容量上限。
6. 写页面质量监控脚本，把本轮评分证据自动化。

## 当前不确定项

- 未跑真实 DeepSeek 成功链路；需要有效 API Key 和真实 DOCX/PPTX 样本文档。
- 没有官方 Lighthouse JSON，本轮性能分基于 Chrome headless、响应头、资源体积、测试和代码审查估算。
- `.doc/.ppt` 是否要支持，取决于后续是否安装并验证 LibreOffice。

## DeepSeek V4 缓存与定价口径

- 官方 Context Caching 是默认开启的前缀缓存，缓存单位是 64 tokens，命中取决于 prompt 前缀是否完全复用。
- 官方没有给固定缓存 TTL；文档口径是未长期使用的缓存会自动清理，持续时间可能从几小时到几天。
- 定价不要承诺“跨天一定缓存命中”。建议按首轮冷启动价 + 同文档短时间追改价拆分。
- 可执行价格结构：每份文档首轮按 cache miss 成本覆盖；30-60 分钟内同一文档连续修改按 cache hit 折扣或包含若干次追改；超时、换文档或大改 prompt 前缀时重新按冷启动估算。

## 2026-05-02 修复结果

已完成：

- 选区联动：左侧选中预览段落后，会弹出“已选中文稿内容”卡片，可把目标 ID、选中文本、对象原文、样式和批注上下文带入右侧输入框。
- 批注/审阅定位：批注项和审阅目标改为可点击项，会切回编辑视图并定位到对应段落或对象。
- 完成后重新预览修改稿：流式执行完成后自动切回编辑视图，并通过 session 下载地址重新 inspect 最新修改稿。
- 移动端修复：390px 宽度下工具栏、文稿页、工作区和 AI 面板不再越界。
- 可访问性修复：补充 `meta description`、favicon、编辑/预览区语义、状态/错误 ARIA、页面内 notice，替换阻塞式 `alert/confirm`。
- 本地化修复：把步骤状态 `OK/iter/hit/miss/style/to` 等残留英文改为中文展示。
- 生产可靠性护栏：关闭 `X-Powered-By`，增加 CSP、Referrer-Policy、X-Content-Type-Options、X-Frame-Options、Permissions-Policy、HTTPS HSTS；CORS 改为本地/显式 allowlist；session cache 增加总数量和总字节上限；自动执行完成事件改为直接交付 `Buffer` 给 session，保留旧 base64 兼容分支。
- 本地可打开性修复：HSTS 改为只在 `ENABLE_HSTS=1` 时启用，避免自签名证书下浏览器无法绕过；新增 `http://127.0.0.1:3001/` 作为日常网站调试入口。
- 品牌视觉修复：移除网页和项目说明里的第三方助手命名；侧栏、主按钮、favicon 和全局 accent 改为统一的深海鲸鱼蓝。

验证结果：

- `npm run typecheck` 通过。
- `node --check apps/web/app.js` 通过。
- `npm run test:ts -- tests/sessionCache.test.ts tests/apiSecurity.test.ts` 通过：56 个 TypeScript 测试全过。
- `npm run build` 通过。
- `npm test` 通过：56 个编译后测试全过。
- 系统 Chrome 回归通过：上传带批注 DOCX、选区带入、批注跳转、mock 流完成、下载按钮、完成后刷新修改稿预览、390px 移动端无越界、控制台和网络错误为 0。
- 系统 Chrome 已验证 `http://127.0.0.1:3001/` 可直接打开，标题、CSS、JS 均加载成功，控制台错误为 0。
- 系统 Chrome 已验证鲸鱼蓝版本：页面可见文本不含第三方助手名称，右侧侧栏背景为 `rgb(8, 47, 71)`，主按钮为 `rgb(74, 174, 224)`，控制台错误为 0。

剩余事项：

- 仍未跑真实 DeepSeek 成功链路。
- 左侧仍不是完整 Word WYSIWYG，只是更可信的 manifest preview + patch/result preview。
- `.doc/.ppt` 旧格式转换仍取决于 LibreOffice 可用性。
