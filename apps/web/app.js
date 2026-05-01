(() => {
  'use strict';

  // ---------- Element refs ----------
  const $ = (id) => document.getElementById(id);
  const dropzone = $('dropzone');
  const fileInput = $('file-input');
  const dzEmpty = $('dropzone-empty');
  const dzFilled = $('dropzone-filled');
  const fileNameEl = $('file-name');
  const fileMetaEl = $('file-meta');
  const fileClear = $('file-clear');
  const intentInput = $('intent-input');
  const refAdd = $('ref-add');
  const refInput = $('ref-input');
  const refList = $('ref-list');
  const btnRun = $('btn-run');
  const keyToggle = $('key-toggle');
  const keyPanel = $('key-panel');
  const keyInput = $('key-input');
  const keySave = $('key-save');
  const keyClear = $('key-clear');
  const keyRemember = $('key-remember');
  const keyState = $('key-state');
  const btnLabel = btnRun.querySelector('.btn-label');
  const btnSpinner = btnRun.querySelector('.btn-spinner');
  const progressCard = $('progress-card');
  const statusIcon = $('progress-status-icon');
  const statusText = $('progress-status-text');
  const rationaleEl = $('rationale');
  const stepsEl = $('steps');
  const metricsEl = $('metrics');
  const mCacheRatio = $('m-cache-ratio');
  const mCacheSub = $('m-cache-sub');
  const mInput = $('m-input');
  const mOutput = $('m-output');
  const mIters = $('m-iters');
  const whaleStage = $('whale-stage');
  const whaleCaption = $('whale-caption');
  const finalActions = $('final-actions');
  const btnDownload = $('btn-download');
  const btnReedit = $('btn-reedit');
  const reeditStatus = $('reedit-status');
  const btnRestart = $('btn-restart');
  const errorBox = $('error-box');
  const docState = $('doc-state');
  const documentPage = $('document-page');
  const aiLivePill = $('ai-live-pill');
  const chatScroll = $('chat-scroll');
  const compactUpload = $('compact-upload');
  const editorModeButtons = Array.from(document.querySelectorAll('[data-editor-mode]'));

  // ---------- State ----------
  let selectedFile = null;
  let referenceFiles = [];
  let modifiedBlob = null;
  let modifiedFilename = null;
  let modifiedDownloadUrl = null;
  let totalIterations = 0;
  let activeSession = null;
  let sessionTimer = null;
  let followUpMode = false;
  let isRunning = false;
  let editorMode = 'edit';
  let editorViewState = { kind: 'empty' };
  let documentInfo = null;
  let documentInfoKey = '';
  let inspectingDocument = false;
  let inspectError = '';
  let inspectRequestId = 0;
  let previewUpdatedIds = new Set();

  const REF_ACCEPTED = /\.(pdf|docx?|pptx?|md|markdown|txt)$/i;
  const MAX_REFS = 6;
  const KEY_STORAGE = 'office-agent.deepseek-key';
  const KEY_REMEMBER = 'office-agent.deepseek-key-remember';

  // ---------- API Key ----------
  let inMemoryKey = '';

  function getStoredKey() {
    try { return localStorage.getItem(KEY_STORAGE) || ''; } catch { return ''; }
  }

  function getStoredRemember() {
    try { return localStorage.getItem(KEY_REMEMBER) !== '0'; } catch { return true; }
  }

  function setStoredKey(key) {
    try {
      if (key) localStorage.setItem(KEY_STORAGE, key);
      else localStorage.removeItem(KEY_STORAGE);
    } catch {}
  }

  function setStoredRemember(remember) {
    try { localStorage.setItem(KEY_REMEMBER, remember ? '1' : '0'); } catch {}
  }

  function maskKey(key) {
    if (!key) return '';
    if (key.length <= 8) return key.slice(0, 3) + '...';
    return key.slice(0, 5) + '...' + key.slice(-4);
  }

  function currentKey() {
    return inMemoryKey || getStoredKey();
  }

  function refreshKeyState() {
    const k = currentKey();
    if (k) {
      keyState.textContent = '已设置 ' + maskKey(k);
      keyState.classList.add('set');
    } else {
      keyState.textContent = '未设置';
      keyState.classList.remove('set');
    }
  }

  inMemoryKey = getStoredKey();
  keyRemember.checked = getStoredRemember();
  refreshKeyState();

  keyToggle.addEventListener('click', () => {
    const open = !keyPanel.hidden;
    keyPanel.hidden = open;
    keyToggle.setAttribute('aria-expanded', String(!open));
    if (!open) keyInput.focus();
  });

  keySave.addEventListener('click', () => {
    const v = keyInput.value.trim();
    if (v && !v.startsWith('sk-')) {
      if (!confirm('看起来不是标准 sk- 开头的 DeepSeek key，仍然保存？')) return;
    }
    inMemoryKey = v;
    if (keyRemember.checked) {
      setStoredKey(v);
      setStoredRemember(true);
    } else {
      setStoredKey('');
      setStoredRemember(false);
    }
    keyInput.value = '';
    refreshKeyState();
    keyPanel.hidden = true;
    keyToggle.setAttribute('aria-expanded', 'false');
  });

  keyClear.addEventListener('click', () => {
    inMemoryKey = '';
    setStoredKey('');
    keyInput.value = '';
    refreshKeyState();
  });

  // ---------- UI helpers ----------
  function scrollChat() {
    requestAnimationFrame(() => {
      chatScroll.scrollTop = chatScroll.scrollHeight;
    });
  }

  function setLiveState(label, state = 'idle') {
    aiLivePill.textContent = label;
    aiLivePill.classList.remove('running', 'done', 'error');
    if (state !== 'idle') aiLivePill.classList.add(state);
  }

  function setDocState(label, loaded = false) {
    docState.textContent = label;
    docState.classList.toggle('loaded', loaded);
  }

  function setRunStatus(label, state = 'running') {
    statusText.textContent = label;
    statusIcon.className = `status-dot ${state}`;
    setLiveState(label, state === 'error' ? 'error' : state === 'done' ? 'done' : 'running');
    scrollChat();
  }

  function setEditorMode(mode) {
    editorMode = mode;
    renderDocumentWorkspace();
    if (mode !== 'edit') inspectCurrentDocument();
  }

  function updateToolButtons() {
    editorModeButtons.forEach((button) => {
      const active = button.dataset.editorMode === editorMode;
      button.classList.toggle('active', active);
      button.setAttribute('aria-pressed', String(active));
    });
  }

  function currentInspectableFile() {
    if (modifiedBlob) {
      const name = modifiedFilename || activeSession?.filename || selectedFile?.name || 'modified-document.docx';
      return new File([modifiedBlob], name, { type: modifiedBlob.type || 'application/octet-stream' });
    }
    return selectedFile;
  }

  function inspectionKey(file) {
    return `${file.name}|${file.size}|${file.lastModified || 0}`;
  }

  function clearDocumentInfo() {
    inspectRequestId++;
    documentInfo = null;
    documentInfoKey = '';
    inspectError = '';
    inspectingDocument = false;
    previewUpdatedIds = new Set();
  }

  function renderDocumentWorkspace() {
    updateToolButtons();
    documentPage.setAttribute('contenteditable', editorMode === 'edit' ? 'true' : 'false');

    if (editorMode === 'comments') {
      renderCommentsView();
      return;
    }
    if (editorMode === 'review') {
      renderReviewView();
      return;
    }
    renderEditView();
  }

  function renderEditView() {
    const source = currentInspectableFile();

    if (source && inspectingDocument) {
      documentPage.innerHTML = `
        <div class="doc-kicker">正在打开文稿</div>
        <h1>${escapeHtml(source.name)}</h1>
        <p>正在读取文档结构并生成左侧预览。</p>
        <div class="doc-skeleton"></div>
      `;
      return;
    }

    if (source && inspectError) {
      documentPage.innerHTML = `
        <div class="doc-kicker">预览读取失败</div>
        <h1>${escapeHtml(source.name)}</h1>
        <p class="doc-warning">${escapeHtml(inspectError)}</p>
        <div class="doc-actions">
          <button type="button" class="doc-action" data-inspect-refresh="1">重新读取</button>
        </div>
      `;
      return;
    }

    if (documentInfo?.preview) {
      renderLivePreview();
      return;
    }

    if (editorViewState.kind === 'file') {
      documentPage.innerHTML = `
        <div class="doc-kicker">${escapeHtml(editorViewState.ext.toUpperCase())} 文档</div>
        <h1>${escapeHtml(editorViewState.name)}</h1>
        <p>文档已载入，正在准备预览。你可以在右侧输入要修改的目标，左侧会在执行过程中同步显示可见变更。</p>
      `;
      return;
    }

    if (editorViewState.kind === 'result') {
      documentPage.innerHTML = `
        <div class="doc-kicker">${editorViewState.success ? '已生成修改稿' : '部分完成'}</div>
        <h1>${escapeHtml(editorViewState.name)}</h1>
        <p>${editorViewState.success ? '修改流程已完成。' : '部分步骤未完成。'}右侧可以下载结果，也可以继续发出下一条修改指令。</p>
        <p>继续修改会复用当前会话中的最新文档，避免重新上传。</p>
      `;
      return;
    }

    documentPage.innerHTML = `
      <div class="doc-kicker">Whale Editor Workspace</div>
      <h1>未打开文档</h1>
      <p>选择一个 Word 或 PowerPoint 文件后，这里会保留当前工作区状态。</p>
      <p>右侧对话框会显示规划、执行、校验和缓存命中情况。</p>
    `;
  }

  function renderLivePreview() {
    if (documentInfo.preview.type === 'pptx') {
      renderPptxPreview();
      return;
    }
    renderDocxPreview();
  }

  function firstPreviewTitle(paragraphs, fallback) {
    const firstText = paragraphs.find((p) => p.text && p.text.trim())?.text?.trim();
    return firstText || fallback || '未命名文档';
  }

  function paraClass(p, index) {
    const classes = ['doc-paragraph'];
    const style = `${p.styleId || ''} ${p.styleName || ''}`.toLowerCase();
    if (p.headingLevel === 1 || style.includes('title') || (index === 0 && (p.text || '').length <= 80)) {
      classes.push('doc-title-line');
    } else if (p.headingLevel || style.includes('heading')) {
      classes.push('doc-heading-line');
    }
    if (p.isListItem) classes.push('doc-list-line');
    if (p.commentIds?.length) classes.push('doc-commented-line');
    if (previewUpdatedIds.has(p.id)) classes.push('doc-updated-line');
    if (p.format?.alignment === 'center') classes.push('align-center');
    if (p.format?.alignment === 'right') classes.push('align-right-text');
    if (p.format?.alignment === 'both' || p.format?.alignment === 'justify') classes.push('align-justify');
    return classes.join(' ');
  }

  function renderDocxPreview() {
    const preview = documentInfo.preview;
    const paragraphs = (preview.paragraphs || []).filter((p) => (p.text || '').trim());
    const title = firstPreviewTitle(paragraphs, documentInfo.filename);
    const visibleParagraphs = paragraphs.slice(0, 90);
    const paragraphHtml = visibleParagraphs.length
      ? visibleParagraphs.map((p, index) => {
          const level = p.headingLevel || (index === 0 ? 1 : 0);
          const text = escapeHtml(p.text || '');
          const marker = p.isListItem ? '<span class="doc-list-marker">•</span>' : '';
          const tag = level === 1 ? 'h1' : level ? 'h2' : 'p';
          return `<${tag} class="${paraClass(p, index)}" data-preview-id="${escapeHtml(p.id)}">${marker}<span>${text}</span></${tag}>`;
        }).join('')
      : `<p class="doc-paragraph empty">未提取到可显示正文。文档结构仍可在“审阅”里查看。</p>`;

    const tablesHtml = (preview.tables || []).slice(0, 4).map((table) => `
      <table class="doc-table-preview" data-preview-id="${escapeHtml(table.id)}">
        <tbody>
          ${(table.rows || []).slice(0, 8).map((row) => `
            <tr>
              ${(row.cells || []).slice(0, 6).map((cell) => `
                <td class="${previewUpdatedIds.has(cell.id) ? 'doc-updated-line' : ''}" data-preview-id="${escapeHtml(cell.id)}">${escapeHtml(cell.text || '')}</td>
              `).join('')}
            </tr>
          `).join('')}
        </tbody>
      </table>
    `).join('');

    documentPage.innerHTML = `
      <div class="doc-preview-meta">
        <span>DOCX 文稿预览</span>
        <span>${escapeHtml(documentInfo.filename)}</span>
      </div>
      <div class="doc-live-note">${isRunning ? '修改中，左侧会随步骤更新' : '已载入，可在右侧发出修改指令'}</div>
      <section class="doc-preview-content" aria-label="文稿实时预览">
        ${paragraphHtml}
        ${tablesHtml}
      </section>
    `;
  }

  function renderPptxPreview() {
    const preview = documentInfo.preview;
    const slides = preview.slides || [];
    documentPage.innerHTML = `
      <div class="doc-preview-meta">
        <span>PPTX 预览</span>
        <span>${escapeHtml(documentInfo.filename)}</span>
      </div>
      <div class="slide-preview-stack">
        ${slides.slice(0, 18).map((slide) => `
          <section class="slide-preview" data-preview-id="${escapeHtml(slide.id)}">
            <div class="slide-index">第 ${slide.index} 页</div>
            <h1>${escapeHtml(slide.title || `幻灯片 ${slide.index}`)}</h1>
            ${(slide.shapes || []).filter((shape) => shape.text).slice(0, 8).map((shape) => `
              <p class="${previewUpdatedIds.has(shape.id) ? 'doc-updated-line' : ''}" data-preview-id="${escapeHtml(shape.id)}">${escapeHtml(shape.text)}</p>
            `).join('')}
          </section>
        `).join('')}
      </div>
    `;
  }

  function renderInspectShell(title, kicker, bodyHtml, actionsHtml = '') {
    documentPage.innerHTML = `
      <div class="doc-kicker">${escapeHtml(kicker)}</div>
      <h1>${escapeHtml(title)}</h1>
      ${bodyHtml}
      ${actionsHtml ? `<div class="doc-actions">${actionsHtml}</div>` : ''}
    `;
  }

  function renderInspectionState(title, kicker) {
    const source = currentInspectableFile();
    if (!source) {
      renderInspectShell(
        title,
        kicker,
        '<p>还没有可检查的文档。先在右侧上传 Word 或 PowerPoint 文件，再回到这里查看批注和审阅信息。</p>',
        quickIntentButton('按文档批注逐条修改，保持原格式、编号和表格结构。', '填入批注处理指令')
      );
      return true;
    }
    if (inspectingDocument) {
      renderInspectShell(
        title,
        kicker,
        `<p>正在读取 ${escapeHtml(source.name)} 的 OOXML 结构、批注和审阅上下文。</p><div class="doc-skeleton"></div>`
      );
      return true;
    }
    if (inspectError) {
      renderInspectShell(
        title,
        kicker,
        `<p class="doc-warning">读取失败：${escapeHtml(inspectError)}</p>`,
        '<button type="button" class="doc-action" data-inspect-refresh="1">重新读取</button>'
      );
      return true;
    }
    return false;
  }

  function renderCommentsView() {
    if (renderInspectionState('批注', '批注视图')) return;

    const comments = documentInfo?.comments || [];
    const isDocx = documentInfo?.fileType === 'docx';
    const summary = documentInfo?.summary || {};
    const command = '请按文档中的所有批注逐条修改，优先使用批注锚定段落，保持原格式、编号、表格结构和未受影响内容。';

    if (!isDocx) {
      renderInspectShell(
        '批注',
        '批注视图',
        '<p>当前演示面板已读取文件结构。PPTX 批注暂未下沉到批注列表，仍可让右侧 AI 按备注、标题和形状内容做审阅。</p>',
        quickIntentButton('请审阅当前演示文稿，优先处理备注、标题和明显格式问题，保持版式、母版和形状位置。', '填入演示审阅指令')
      );
      return;
    }

    if (comments.length === 0) {
      renderInspectShell(
        '批注',
        '批注视图',
        `<p>已读取 ${escapeHtml(documentInfo.filename)}。没有发现 Word 批注。</p>
         <div class="doc-mini-grid">
           <div><strong>${summary.paragraphs || 0}</strong><span>段落</span></div>
           <div><strong>${summary.tables || 0}</strong><span>表格</span></div>
           <div><strong>${summary.comments || 0}</strong><span>批注</span></div>
         </div>`,
        quickIntentButton('请审阅当前文档，修复明显不一致之处，保持原格式、编号和表格结构。', '填入审阅指令')
      );
      return;
    }

    const rows = comments.map((comment, index) => `
      <div class="review-item">
        <div class="review-item-head">
          <span>${escapeHtml(comment.commentId || `comment-${index + 1}`)}</span>
          <span>${escapeHtml(comment.author || '未知作者')}</span>
        </div>
        <p>${escapeHtml(comment.textPreview)}</p>
        <div class="review-item-meta">${escapeHtml((comment.anchoredParagraphIds || []).join(' · ') || '未定位锚点')}</div>
      </div>
    `).join('');

    renderInspectShell(
      '批注',
      '批注视图',
      `<p>已读取 ${comments.length} 条批注。右侧执行时会把这些批注作为 REVIEW_COMMENT 输入，并优先匹配锚定段落。</p>
       <div class="review-list">${rows}</div>`,
      quickIntentButton(command, '按全部批注修改')
    );
  }

  function renderReviewView() {
    if (renderInspectionState('审阅', '审阅上下文')) return;

    const summary = documentInfo?.summary || {};
    const warnings = documentInfo?.warnings || [];
    const errors = documentInfo?.errors || [];
    const targets = documentInfo?.reviewTargets || [];
    const isDocx = documentInfo?.fileType === 'docx';
    const command = isDocx
      ? '请审阅当前 Word 文档：处理可安全处理的批注和格式不一致，保持原样式、编号、缩进、表格几何和未受影响内容。'
      : '请审阅当前 PowerPoint：检查标题、形状文本、备注和明显版式问题，保持母版、主题、占位符和形状位置。';

    const metrics = isDocx
      ? [
          ['段落', summary.paragraphs || 0],
          ['表格', summary.tables || 0],
          ['批注', summary.comments || 0],
          ['编号', summary.hasNumbering ? '有' : '无'],
        ]
      : [
          ['幻灯片', summary.slides || 0],
          ['形状', summary.shapes || 0],
          ['备注页', summary.speakerNotes || 0],
          ['版式', summary.layouts || 0],
        ];

    const metricHtml = metrics.map(([label, value]) => `
      <div><strong>${escapeHtml(value)}</strong><span>${escapeHtml(label)}</span></div>
    `).join('');

    const targetHtml = targets.length > 0
      ? `<div class="review-list compact">${targets.map((target) => `
          <div class="review-item">
            <div class="review-item-head">
              <span>${escapeHtml(target.id)}</span>
              <span>${target.commentIds ? `${target.commentIds.length} 批注` : target.hasNotes ? '有备注' : ''}</span>
            </div>
            <p>${escapeHtml(target.textPreview || '无文本预览')}</p>
          </div>
        `).join('')}</div>`
      : '<p>没有发现需要特别列出的审阅目标。</p>';

    const issueHtml = [...errors, ...warnings].length > 0
      ? `<div class="doc-warning-list">${[...errors, ...warnings].slice(0, 6).map((item) => `<div>${escapeHtml(item)}</div>`).join('')}</div>`
      : '<p>包结构检查没有返回错误或警告。</p>';

    renderInspectShell(
      '审阅',
      '审阅上下文',
      `<p>已读取 ${escapeHtml(documentInfo?.filename || 'document')}。这里显示 agent 执行前可见的结构信息。</p>
       <div class="doc-mini-grid">${metricHtml}</div>
       <h2>审阅目标</h2>
       ${targetHtml}
       <h2>包检查</h2>
       ${issueHtml}`,
      `${quickIntentButton(command, '填入审阅修复指令')}
       <button type="button" class="doc-action secondary" data-inspect-refresh="1">重新读取</button>`
    );
  }

  function quickIntentButton(intent, label) {
    return `<button type="button" class="doc-action" data-quick-intent="${escapeHtml(intent)}">${escapeHtml(label)}</button>`;
  }

  function localizePlanText(value) {
    if (!value) return '';
    let text = String(value);

    const exactPatterns = [
      [
        /Ensure requested formatting is applied while preserving existing fonts and sizes\./gi,
        '按要求应用格式，同时保留现有字体和字号。'
      ],
      [
        /Ensure '([^']+)' heading \(([^)]+)\) is centered, bold, ([^,]+), matching template heading style\./gi,
        '确保“$1”标题（$2）居中、加粗，并使用 $3，匹配模板标题样式。'
      ],
      [
        /Ensure abstract body paragraphs \(([^)]+)\) use justified alignment with first-line indent\./gi,
        '确保摘要正文段落（$1）使用两端对齐，并设置首行缩进。'
      ],
      [
        /Ensure '([^']+)' line \(([^)]+)\) uses bold for label\./gi,
        '确保“$1”所在行（$2）的标签加粗。'
      ],
      [
        /For English Abstract, ensure 'Abstract' heading \(([^)]+)\) and 'Keywords:' line \(([^)]+)\) follow template formatting\./gi,
        '英文摘要部分：确保“Abstract”标题（$1）和“Keywords:”所在行（$2）遵循模板格式。'
      ],
      [
        /Preserve existing fonts and sizes\./gi,
        '保留现有字体和字号。'
      ],
    ];

    for (const [pattern, replacement] of exactPatterns) {
      text = text.replace(pattern, replacement);
    }

    const replacements = [
      [/\bFormat\b/gi, '格式调整'],
      [/\bUpdate\b/gi, '更新'],
      [/\bEdit\b/gi, '编辑'],
      [/\bReview\b/gi, '审阅'],
      [/\bEnsure\b/gi, '确保'],
      [/\bPreserve\b/gi, '保留'],
      [/\bmatching\b/gi, '匹配'],
      [/\btemplate\b/gi, '模板'],
      [/\bheading style\b/gi, '标题样式'],
      [/\bheading\b/gi, '标题'],
      [/\bbody paragraphs\b/gi, '正文段落'],
      [/\bparagraphs\b/gi, '段落'],
      [/\bparagraph\b/gi, '段落'],
      [/\bcentered\b/gi, '居中'],
      [/\bbold\b/gi, '加粗'],
      [/\bjustified alignment\b/gi, '两端对齐'],
      [/\bfirst-line indent\b/gi, '首行缩进'],
      [/\bline\b/gi, '行'],
      [/\blabel\b/gi, '标签'],
      [/\bEnglish Abstract\b/gi, '英文摘要'],
      [/\babstract body\b/gi, '摘要正文'],
      [/\bAbstract\b/g, '英文摘要'],
      [/\bKeywords\b/g, '关键词'],
      [/\bformatting\b/gi, '格式'],
      [/\bfonts\b/gi, '字体'],
      [/\bsizes\b/gi, '字号'],
      [/\bexisting\b/gi, '现有'],
      [/\bfollow\b/gi, '遵循'],
      [/\buse\b/gi, '使用'],
      [/\bwith\b/gi, '并带有'],
      [/\band\b/gi, '和'],
      [/\bis\b/gi, '为'],
    ];

    for (const [pattern, replacement] of replacements) {
      text = text.replace(pattern, replacement);
    }

    return text
      .replace(/\s*&\s*/g, ' 和 ')
      .replace(/ ,/g, '，')
      .replace(/\.($|\s)/g, '。$1')
      .replace(/\s{2,}/g, ' ')
      .trim();
  }

  function localizeStepTitle(value) {
    return localizePlanText(value)
      .replace(/^(\d+)\.\s*/, '')
      .slice(0, 80);
  }

  function resetDocumentPage() {
    editorViewState = { kind: 'empty' };
    clearDocumentInfo();
    renderDocumentWorkspace();
    setDocState('等待文档', false);
  }

  function updateDocumentPageForFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    editorViewState = { kind: 'file', name: file.name, ext };
    clearDocumentInfo();
    renderDocumentWorkspace();
    setDocState(`已载入 ${file.name}`, true);
    inspectCurrentDocument(true);
  }

  function updateDocumentPageForResult(success, refreshFromFile = false) {
    const name = modifiedFilename || activeSession?.filename || selectedFile?.name || 'document';
    editorViewState = { kind: 'result', name, success };
    renderDocumentWorkspace();
    setDocState(success ? '已生成修改稿' : '部分完成', true);
    if (refreshFromFile) inspectCurrentDocument(true);
  }

  function appendUserMessage(text, reuseSession) {
    const msg = document.createElement('div');
    msg.className = 'message user';
    msg.innerHTML = `
      <div class="message-label">${reuseSession ? '继续修改' : '用户指令'}</div>
      <div class="message-body"></div>
    `;
    msg.querySelector('.message-body').textContent = text || '按参考材料修改';
    chatScroll.insertBefore(msg, progressCard);
    scrollChat();
  }

  function clearProgress() {
    progressCard.hidden = false;
    stepsEl.innerHTML = '';
    rationaleEl.hidden = true;
    rationaleEl.textContent = '';
    metricsEl.hidden = true;
    finalActions.hidden = true;
    errorBox.hidden = true;
    errorBox.textContent = '';
    mCacheRatio.textContent = '--';
    mCacheSub.textContent = '--';
    mInput.textContent = '--';
    mOutput.textContent = '--';
    mIters.textContent = '--';
    modifiedBlob = null;
    modifiedFilename = null;
    modifiedDownloadUrl = null;
  }

  function canReuseSession() {
    return !!(activeSession && activeSession.expiresInMs > 0 && activeSession.reuseRemaining > 0);
  }

  async function inspectCurrentDocument(force = false) {
    const file = currentInspectableFile();
    if (!file) {
      clearDocumentInfo();
      renderDocumentWorkspace();
      return;
    }

    const key = inspectionKey(file);
    if (!force && documentInfo && documentInfoKey === key) return;

    const requestId = ++inspectRequestId;
    inspectingDocument = true;
    inspectError = '';
    renderDocumentWorkspace();

    const fd = new FormData();
    fd.append('file', file, file.name);

    try {
      const resp = await fetch('/api/document/inspect', { method: 'POST', body: fd });
      if (!resp.ok) {
        const text = await resp.text().catch(() => '');
        throw new Error(text || `HTTP ${resp.status}`);
      }
      const data = await resp.json();
      if (requestId !== inspectRequestId) return;
      documentInfo = data;
      documentInfoKey = key;
      inspectError = '';
    } catch (err) {
      if (requestId !== inspectRequestId) return;
      documentInfo = null;
      documentInfoKey = '';
      inspectError = err instanceof Error ? err.message : String(err);
    } finally {
      if (requestId === inspectRequestId) {
        inspectingDocument = false;
        renderDocumentWorkspace();
      }
    }
  }

  function applyQuickIntent(intent) {
    intentInput.value = intent;
    updateRunEnabled();
    intentInput.focus();
    setLiveState('等待执行', 'done');
    scrollChat();
  }

  function previewParagraphs() {
    return documentInfo?.preview?.paragraphs || [];
  }

  function previewTables() {
    return documentInfo?.preview?.tables || [];
  }

  function previewSlides() {
    return documentInfo?.preview?.slides || [];
  }

  function findPreviewParagraph(id) {
    return previewParagraphs().find((p) => p.id === id);
  }

  function markPreviewUpdated(id) {
    if (!id) return;
    previewUpdatedIds.add(id);
  }

  function parsePreviewId(targetId) {
    const parts = String(targetId || '').split('.');
    if (parts[0] === 'docx' && parts[1] === 'p') return { type: 'paragraph', id: `docx.p.${parts[2]}` };
    if (parts[0] === 'docx' && parts[1] === 'tbl') {
      return {
        type: 'cell',
        tableId: `docx.tbl.${parts[2]}`,
        cellId: `docx.tbl.${parts[2]}.r.${parts[4]}.c.${parts[6]}`,
      };
    }
    if (parts[0] === 'pptx' && parts[1] === 'slide') return { type: 'slide', id: `pptx.slide.${parts[2]}` };
    if (parts[0] === 'pptx' && parts[1] === 'shape') return { type: 'shape', id: targetId };
    return { type: 'unknown', id: targetId };
  }

  function applyOperationsToPreview(ops) {
    if (!documentInfo?.preview || !Array.isArray(ops) || ops.length === 0) return;

    for (const op of ops) {
      const target = parsePreviewId(op.targetId);
      const payload = op.payload || {};
      if (target.type === 'paragraph') {
        const para = findPreviewParagraph(target.id);
        if (!para) continue;
        if (op.op === 'replace_paragraph_text' && typeof payload.text === 'string') {
          para.text = payload.text;
          markPreviewUpdated(para.id);
        } else if (op.op === 'replace_text_in_paragraph' && typeof payload.find === 'string' && typeof payload.replace === 'string') {
          para.text = String(para.text || '').replace(payload.find, payload.replace);
          markPreviewUpdated(para.id);
        } else if (op.op === 'delete_paragraph') {
          const list = previewParagraphs();
          const idx = list.findIndex((p) => p.id === para.id);
          if (idx >= 0) list.splice(idx, 1);
        } else if (op.op === 'insert_paragraph_after' && typeof payload.text === 'string') {
          const list = previewParagraphs();
          const idx = list.findIndex((p) => p.id === para.id);
          if (idx >= 0) {
            const inserted = {
              id: `${para.id}.ins.${Date.now()}`,
              text: payload.text,
              styleId: payload.styleId || para.styleId,
              styleName: para.styleName,
              headingLevel: para.headingLevel,
              isListItem: false,
            };
            list.splice(idx + 1, 0, inserted);
            markPreviewUpdated(inserted.id);
          }
        } else if (op.op === 'apply_paragraph_style' && typeof payload.styleId === 'string') {
          para.styleId = payload.styleId;
          markPreviewUpdated(para.id);
        }
      }

      if (target.type === 'cell' && op.op === 'update_table_cell_text') {
        for (const table of previewTables()) {
          for (const row of table.rows || []) {
            const cell = (row.cells || []).find((c) => c.id === target.cellId);
            if (cell && typeof payload.text === 'string') {
              cell.text = payload.text;
              markPreviewUpdated(cell.id);
            }
          }
        }
      }

      if (target.type === 'slide' && op.op === 'replace_slide_title' && typeof payload.title === 'string') {
        const slide = previewSlides().find((s) => s.id === target.id);
        if (slide) {
          slide.title = payload.title;
          markPreviewUpdated(slide.id);
        }
      }

      if (target.type === 'shape' && op.op === 'replace_shape_text' && typeof payload.text === 'string') {
        for (const slide of previewSlides()) {
          const shape = (slide.shapes || []).find((s) => s.id === target.id);
          if (shape) {
            shape.text = payload.text;
            markPreviewUpdated(shape.id);
          }
        }
      }
    }

    if (editorMode === 'edit') renderDocumentWorkspace();
  }

  // ---------- File handling ----------
  const ACCEPTED = /\.(docx|pptx|doc|ppt)$/i;

  function setFile(file) {
    const emptyDrop = document.getElementById('empty-drop');
    const docPage = document.getElementById('document-page');
    if (!file) {
      selectedFile = null;
      followUpMode = false;
      modifiedBlob = null;
      modifiedFilename = null;
      modifiedDownloadUrl = null;
      clearDocumentInfo();
      dzEmpty.hidden = false;
      dzFilled.hidden = true;
      resetDocumentPage();
      if (emptyDrop) emptyDrop.hidden = false;
      if (docPage) docPage.hidden = true;
      updateRunEnabled();
      return;
    }
    if (!ACCEPTED.test(file.name)) {
      alert('只支持 .docx / .pptx / .doc / .ppt 文件');
      return;
    }
    if (file.size > 50 * 1024 * 1024) {
      alert('文件超过 50MB 上限');
      return;
    }
    clearSession();
    selectedFile = file;
    fileNameEl.textContent = file.name;
    fileMetaEl.textContent = `${formatSize(file.size)} · ${file.name.split('.').pop().toLowerCase()}`;
    dzEmpty.hidden = true;
    dzFilled.hidden = false;
    if (emptyDrop) emptyDrop.hidden = true;
    if (docPage) docPage.hidden = false;
    updateDocumentPageForFile(file);
    updateRunEnabled();
  }

  function formatSize(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
  }

  function updateRunEnabled() {
    const hasIntent = intentInput.value.trim() || referenceFiles.length > 0;
    const hasTarget = selectedFile || canReuseSession();
    btnRun.disabled = isRunning || !(hasTarget && hasIntent);
  }

  // ---------- References ----------
  function addReferences(fileList) {
    for (const f of fileList) {
      if (referenceFiles.length >= MAX_REFS) {
        alert(`最多只能添加 ${MAX_REFS} 份参考材料`);
        break;
      }
      if (!REF_ACCEPTED.test(f.name)) {
        alert(`不支持的格式：${f.name}`);
        continue;
      }
      if (f.size > 50 * 1024 * 1024) {
        alert(`参考材料 ${f.name} 超过 50MB`);
        continue;
      }
      const dup = referenceFiles.find((r) => r.name === f.name && r.size === f.size);
      if (dup) continue;
      referenceFiles.push(f);
    }
    renderRefList();
    updateRunEnabled();
  }

  function refKind(filename) {
    const m = filename.toLowerCase().match(/\.(pdf|docx?|pptx?|md|markdown|txt)$/);
    if (!m) return 'FILE';
    return m[1] === 'markdown' ? 'MD' : m[1].toUpperCase();
  }

  function renderRefList() {
    refList.innerHTML = '';
    referenceFiles.forEach((f, i) => {
      const li = document.createElement('li');
      li.className = 'ref-item';
      li.innerHTML = `
        <span class="ref-kind"></span>
        <span class="ref-name"></span>
        <span class="ref-size"></span>
        <button type="button" class="ref-remove" aria-label="移除">x</button>
      `;
      li.querySelector('.ref-kind').textContent = refKind(f.name);
      li.querySelector('.ref-name').textContent = f.name;
      li.querySelector('.ref-size').textContent = formatSize(f.size);
      li.querySelector('.ref-remove').addEventListener('click', () => {
        referenceFiles.splice(i, 1);
        renderRefList();
        updateRunEnabled();
      });
      refList.appendChild(li);
    });
  }

  // ---------- Event listeners ----------
  editorModeButtons.forEach((button) => {
    button.addEventListener('click', () => {
      setEditorMode(button.dataset.editorMode || 'edit');
    });
  });

  compactUpload.addEventListener('click', () => fileInput.click());

  documentPage.addEventListener('click', (e) => {
    const quickIntent = e.target.closest('[data-quick-intent]');
    if (quickIntent) {
      applyQuickIntent(quickIntent.dataset.quickIntent || '');
      return;
    }
    const refresh = e.target.closest('[data-inspect-refresh]');
    if (refresh) inspectCurrentDocument(true);
  });

  dropzone.addEventListener('click', (e) => {
    if (e.target === fileClear) return;
    fileInput.click();
  });

  dropzone.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      fileInput.click();
    }
  });

  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragging');
  });

  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragging'));

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragging');
    const f = e.dataTransfer.files[0];
    if (f) setFile(f);
  });

  // Make the entire left editor-stage a drop target (the new empty-state UI).
  const editorStage = document.getElementById('editor-stage');
  if (editorStage) {
    editorStage.addEventListener('dragover', (e) => {
      e.preventDefault();
      editorStage.classList.add('dragging');
    });
    editorStage.addEventListener('dragleave', (e) => {
      // Only clear when leaving the actual stage, not its children
      if (e.target === editorStage) editorStage.classList.remove('dragging');
    });
    editorStage.addEventListener('drop', (e) => {
      e.preventDefault();
      editorStage.classList.remove('dragging');
      const f = e.dataTransfer.files[0];
      if (f) setFile(f);
    });
    editorStage.addEventListener('click', (e) => {
      // Click on empty area opens file picker (but don't hijack contenteditable)
      const docPage = document.getElementById('document-page');
      if (!docPage || docPage.hidden) fileInput.click();
    });
  }

  fileInput.addEventListener('change', (e) => setFile(e.target.files[0]));

  fileClear.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.value = '';
    clearSession();
    setFile(null);
  });

  intentInput.addEventListener('input', updateRunEnabled);

  refAdd.addEventListener('click', () => refInput.click());

  refInput.addEventListener('change', (e) => {
    addReferences(e.target.files);
    refInput.value = '';
  });

  btnRun.addEventListener('click', () => {
    runAuto({ reuseSession: followUpMode && canReuseSession() });
  });

  intentInput.addEventListener('keydown', (e) => {
    if ((e.metaKey || e.ctrlKey) && e.key === 'Enter' && !btnRun.disabled) {
      e.preventDefault();
      btnRun.click();
    }
  });

  btnReedit.addEventListener('click', () => {
    if (!canReuseSession()) return;
    followUpMode = true;
    intentInput.value = '';
    referenceFiles = [];
    renderRefList();
    updateRunEnabled();
    intentInput.focus();
    setLiveState('等待指令', 'done');
    scrollChat();
  });

  btnRestart.addEventListener('click', () => {
    progressCard.hidden = true;
    stepsEl.innerHTML = '';
    rationaleEl.hidden = true;
    metricsEl.hidden = true;
    finalActions.hidden = true;
    errorBox.hidden = true;
    modifiedBlob = null;
    modifiedDownloadUrl = null;
    clearSession();
    fileInput.value = '';
    setFile(null);
    referenceFiles = [];
    renderRefList();
    intentInput.value = '';
    setLiveState('空闲');
    updateRunEnabled();
  });

  btnDownload.addEventListener('click', () => {
    if (modifiedDownloadUrl) {
      const a = document.createElement('a');
      a.href = modifiedDownloadUrl;
      a.download = modifiedFilename || 'modified-document';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      return;
    }

    if (!modifiedBlob) return;
    const url = URL.createObjectURL(modifiedBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = modifiedFilename || 'modified-document';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  function clearSession() {
    activeSession = null;
    followUpMode = false;
    if (sessionTimer) {
      clearInterval(sessionTimer);
      sessionTimer = null;
    }
    btnReedit.hidden = true;
    reeditStatus.textContent = '';
  }

  function startSessionCountdown() {
    if (sessionTimer) clearInterval(sessionTimer);
    if (!activeSession) return;
    const tick = () => {
      if (!activeSession) return;
      activeSession.expiresInMs = Math.max(0, activeSession.expiresInMs - 1000);
      const m = Math.floor(activeSession.expiresInMs / 60000);
      const s = Math.floor((activeSession.expiresInMs % 60000) / 1000);
      reeditStatus.textContent = `(${m}:${String(s).padStart(2, '0')} · 剩 ${activeSession.reuseRemaining} 次)`;
      if (activeSession.expiresInMs <= 0) {
        clearSession();
        updateRunEnabled();
      }
    };
    tick();
    sessionTimer = setInterval(tick, 1000);
  }

  // ---------- Run ----------
  async function runAuto({ reuseSession }) {
    const willReuse = reuseSession && canReuseSession();
    if (!willReuse && !selectedFile) return;

    const intent = intentInput.value.trim();
    if (!intent && referenceFiles.length === 0) return;

    isRunning = true;
    updateRunEnabled();
    appendUserMessage(intent, willReuse);
    clearProgress();
    btnLabel.hidden = true;
    btnSpinner.hidden = false;
    setRunStatus(willReuse ? '复用当前文档' : '上传文档', 'running');
    setDocState('修改中', true);
    totalIterations = 0;
    showWhale(willReuse ? '读取当前会话' : '上传中');

    const fd = new FormData();
    fd.append('intent', intent);
    if (willReuse) {
      fd.append('sessionId', activeSession.id);
    } else {
      fd.append('file', selectedFile);
      const ext = selectedFile.name.split('.').pop().toLowerCase();
      fd.append('fileType', ext === 'doc' ? 'docx' : ext === 'ppt' ? 'pptx' : ext);
    }
    for (const ref of referenceFiles) {
      fd.append('references', ref, ref.name);
    }

    intentInput.value = '';
    referenceFiles = [];
    renderRefList();
    updateRunEnabled();

    const headers = {};
    const userKey = currentKey();
    if (userKey) headers['X-DeepSeek-Key'] = userKey;

    let resp;
    try {
      resp = await fetch('/api/agent/auto', { method: 'POST', body: fd, headers });
    } catch (err) {
      showError(`请求失败：${err.message}`);
      resetRunButton();
      return;
    }

    if (!resp.ok || !resp.body) {
      const txt = await resp.text().catch(() => '');
      showError(`HTTP ${resp.status}\n${txt}`);
      resetRunButton();
      return;
    }

    setRunStatus('规划中', 'running');

    const reader = resp.body.getReader();
    const decoder = new TextDecoder('utf-8');
    let buf = '';
    let gotTerminalEvent = false;

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buf += decoder.decode(value, { stream: true });
        let nl;
        while ((nl = buf.indexOf('\n')) !== -1) {
          const line = buf.slice(0, nl).trim();
          buf = buf.slice(nl + 1);
          if (!line) continue;
          try {
            const ev = JSON.parse(line);
            if (ev.type === 'complete' || ev.type === 'error') gotTerminalEvent = true;
            handleEvent(ev);
          } catch (e) {
            console.warn('Bad event line:', line, e);
          }
        }
      }
      if (buf.trim()) {
        try {
          const ev = JSON.parse(buf.trim());
          if (ev.type === 'complete' || ev.type === 'error') gotTerminalEvent = true;
          handleEvent(ev);
        } catch (e) { console.warn(e); }
      }
    } catch (err) {
      if (!gotTerminalEvent) {
        showError(`流式读取失败：${err.message}`);
      }
    } finally {
      resetRunButton();
    }
  }

  function resetRunButton() {
    isRunning = false;
    btnLabel.hidden = false;
    btnSpinner.hidden = true;
    updateRunEnabled();
    hideWhale();
  }

  function showWhale(caption) {
    if (whaleCaption) whaleCaption.textContent = caption || 'DeepSeek 正在处理';
    whaleStage.classList.remove('idle');
    scrollChat();
  }

  function hideWhale() {
    whaleStage.classList.add('idle');
  }

  // ---------- Event handlers ----------
  function handleEvent(ev) {
    switch (ev.type) {
      case 'converting':
        setRunStatus(`正在把 ${ev.from.toUpperCase()} 转成 OOXML`, 'running');
        showWhale('转换格式中');
        break;
      case 'session':
        activeSession = ev.session;
        startSessionCountdown();
        break;
      case 'manifest_summary':
        setRunStatus(`已扫描 ${ev.objectCount} 个对象`, 'running');
        showWhale('扫描文档结构');
        if (editorMode === 'edit') renderDocumentWorkspace();
        break;
      case 'plan':
        renderPlan(ev.plan);
        setRunStatus(`开始执行 ${ev.plan.steps.length} 个步骤`, 'running');
        showWhale(`执行 ${ev.plan.steps.length} 个步骤`);
        break;
      case 'step_start':
        markStep(ev.index, 'running');
        setRunStatus(`执行中 (${ev.index + 1}/${ev.total})：${localizeStepTitle(ev.step.title)}`, 'running');
        showWhale(`第 ${ev.index + 1} 步 · ${localizeStepTitle(ev.step.title)}`);
        break;
      case 'step_done':
        markStep(ev.index, ev.success ? 'done' : 'error', {
          iters: ev.iterations,
          cacheRatio: ev.usage?.cacheHitRatio,
        });
        renderStepOps(ev.index, ev.operations);
        applyOperationsToPreview(ev.operations);
        totalIterations += ev.iterations;
        if (!ev.success) appendStepError(ev.index, ev.summary);
        scrollChat();
        break;
      case 'step_error':
        markStep(ev.index, 'error');
        appendStepError(ev.index, ev.error);
        setRunStatus('步骤失败', 'error');
        break;
      case 'complete':
        finishRun(ev);
        break;
      case 'error':
        showError(ev.error);
        break;
      default:
        console.log('Unknown event:', ev);
    }
  }

  function renderPlan(plan) {
    if (plan.rationale) {
      rationaleEl.textContent = localizePlanText(plan.rationale);
      rationaleEl.hidden = false;
    }
    stepsEl.innerHTML = '';
    plan.steps.forEach((step, i) => {
      const li = document.createElement('li');
      li.className = 'step pending';
      li.dataset.index = String(i);
      li.innerHTML = `
        <div class="step-icon">-</div>
        <div class="step-body">
          <div class="step-title"></div>
          <div class="step-task"></div>
        </div>
        <div class="step-meta"></div>
      `;
      li.querySelector('.step-title').textContent = `${i + 1}. ${localizeStepTitle(step.title)}`;
      li.querySelector('.step-task').textContent = localizePlanText(step.task);
      stepsEl.appendChild(li);
    });
    scrollChat();
  }

  function markStep(index, state, meta) {
    const li = stepsEl.querySelector(`.step[data-index="${index}"]`);
    if (!li) return;
    li.classList.remove('pending', 'running', 'done', 'error');
    li.classList.add(state);
    const icon = li.querySelector('.step-icon');
    icon.textContent = state === 'running' ? '...' : state === 'done' ? 'OK' : state === 'error' ? '!' : '-';
    if (meta) {
      const metaEl = li.querySelector('.step-meta');
      const parts = [];
      if (typeof meta.iters === 'number') parts.push(`${meta.iters} iter`);
      if (typeof meta.cacheRatio === 'number') parts.push(`${(meta.cacheRatio * 100).toFixed(0)}% 命中`);
      metaEl.textContent = parts.join(' · ');
    }
  }

  const OP_LABELS = {
    replace_paragraph_text: '替换段落',
    replace_text_in_paragraph: '段内查找替换',
    update_table_cell_text: '改表格单元格',
    insert_paragraph_after: '插入段落',
    delete_paragraph: '删除段落',
    apply_paragraph_style: '应用段落样式',
    update_header_text: '改页眉',
    update_footer_text: '改页脚',
    replace_shape_text: '替换形状文本',
    replace_slide_title: '改幻灯片标题',
    update_speaker_notes: '改备注',
    insert_slide_from_layout: '新增幻灯片',
    delete_slide: '删除幻灯片',
    move_shape: '移动形状',
    resize_shape: '调整大小',
    apply_text_style: '应用文本样式',
    replace_image: '替换图片',
    fit_text_to_shape: '自适应文本',
  };

  function truncate(s, n) {
    if (typeof s !== 'string') return s;
    return s.length > n ? s.slice(0, n) + '...' : s;
  }

  function renderOpDetail(op) {
    const p = op.payload || {};
    if ('find' in p && 'replace' in p) {
      return `<span class="from">${escapeHtml(truncate(p.find, 60))}</span><span class="arrow">-&gt;</span><span class="to">${escapeHtml(truncate(p.replace, 60))}</span>`;
    }
    if ('text' in p) return `<span class="to">${escapeHtml(truncate(p.text, 100))}</span>`;
    if ('title' in p) return `<span class="to">${escapeHtml(truncate(p.title, 100))}</span>`;
    if ('styleId' in p) return `<span class="to">style: ${escapeHtml(p.styleId)}</span>`;
    if ('x' in p && 'y' in p) return `<span class="to">to (${p.x}, ${p.y})</span>`;
    if ('width' in p && 'height' in p) return `<span class="to">${p.width} x ${p.height}</span>`;
    if (typeof p === 'object') {
      const summary = Object.entries(p)
        .filter(([_, v]) => v !== undefined && v !== null)
        .map(([k, v]) => `${k}=${typeof v === 'string' ? truncate(v, 30) : v}`)
        .join(' · ');
      return `<span class="to">${escapeHtml(summary)}</span>`;
    }
    return '';
  }

  function escapeHtml(s) {
    if (s === null || s === undefined) return '';
    return String(s).replace(/[&<>"']/g, (c) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
    }[c]));
  }

  function renderStepOps(index, ops) {
    if (!ops || ops.length === 0) return;
    const li = stepsEl.querySelector(`.step[data-index="${index}"]`);
    if (!li) return;
    li.querySelector('.step-ops')?.remove();
    const wrap = document.createElement('div');
    wrap.className = 'step-ops';
    for (const op of ops) {
      const row = document.createElement('div');
      row.className = 'step-op';
      const kind = OP_LABELS[op.op] || op.op;
      row.innerHTML = `
        <span class="step-op-kind">${escapeHtml(kind)}</span>
        <span class="step-op-target">${escapeHtml(op.targetId || '')}</span>
        <span class="step-op-detail">${renderOpDetail(op)}</span>
      `;
      wrap.appendChild(row);
    }
    li.querySelector('.step-body').appendChild(wrap);
  }

  function appendStepError(index, msg) {
    const li = stepsEl.querySelector(`.step[data-index="${index}"]`);
    if (!li) return;
    const taskEl = li.querySelector('.step-task');
    if (msg && msg !== taskEl.textContent) {
      taskEl.textContent += `  · 失败：${msg}`;
    }
  }

  function finishRun(ev) {
    hideWhale();
    if (ev.success) {
      setRunStatus('全部步骤已完成', 'done');
    } else {
      setRunStatus('部分步骤失败', 'warn');
    }

    const u = ev.totalUsage || {};
    if (typeof u.cacheHitRatio === 'number') {
      mCacheRatio.textContent = `${(u.cacheHitRatio * 100).toFixed(1)}%`;
      mCacheSub.textContent = `${u.cacheHitTokens || 0} hit · ${u.cacheMissTokens || 0} miss`;
    }
    mInput.textContent = (u.inputTokens || 0).toLocaleString();
    mOutput.textContent = (u.outputTokens || 0).toLocaleString();
    mIters.textContent = totalIterations.toString();
    metricsEl.hidden = false;

    const baseName = selectedFile?.name || activeSession?.filename || 'document';
    const baseExt = baseName.split('.').pop().toLowerCase();
    const ooxmlExt = baseExt === 'doc' ? 'docx' : baseExt === 'ppt' ? 'pptx' : baseExt;
    const renamedBase = (baseExt === 'doc' || baseExt === 'ppt')
      ? baseName.replace(/\.(doc|ppt)$/i, '.' + ooxmlExt)
      : baseName;

    if (ev.modifiedFileBase64) {
      const bytes = Uint8Array.from(atob(ev.modifiedFileBase64), (c) => c.charCodeAt(0));
      const mime = ooxmlExt === 'docx'
        ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        : 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
      modifiedBlob = new Blob([bytes], { type: mime });
      modifiedFilename = `modified_${renamedBase}`;
      modifiedDownloadUrl = null;
      btnDownload.hidden = false;
    } else if (ev.downloadUrl) {
      modifiedBlob = null;
      modifiedDownloadUrl = ev.downloadUrl;
      modifiedFilename = ev.modifiedFilename || `modified_${renamedBase}`;
      btnDownload.hidden = false;
    } else {
      modifiedDownloadUrl = null;
      btnDownload.hidden = true;
    }

    finalActions.hidden = false;
    followUpMode = canReuseSession();
    btnReedit.hidden = !canReuseSession();
    updateDocumentPageForResult(ev.success, !!ev.modifiedFileBase64);
    updateRunEnabled();
    scrollChat();
  }

  function showError(msg) {
    errorBox.hidden = false;
    errorBox.textContent = msg;
    setRunStatus('失败', 'error');
    setDocState(selectedFile ? '修改失败' : '等待文档', !!selectedFile);
    scrollChat();
  }

  resetDocumentPage();
  setLiveState('空闲');
})();
