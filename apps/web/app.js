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
  const editorStage = document.querySelector('.editor-stage');
  const aiLivePill = $('ai-live-pill');
  const chatScroll = $('chat-scroll');
  const compactUpload = $('compact-upload');
  const noticeRegion = $('notice-region');
  const selectionCard = $('selection-card');
  const selectionText = $('selection-text');
  const selectionMeta = $('selection-meta');
  const selectionUse = $('selection-use');
  const pricingCard = $('pricing-card');
  const pricingTier = $('pricing-tier');
  const pricingPrice = $('pricing-price');
  const pricingScope = $('pricing-scope');
  const pricingWindow = $('pricing-window');
  const pricingRounds = $('pricing-rounds');
  const pricingDownload = $('pricing-download');
  const pricingCountdown = $('pricing-countdown');
  const pricingNote = $('pricing-note');
  const themeButtons = Array.from(document.querySelectorAll('[data-theme-choice]'));
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
  let selectedPreviewContext = null;
  let pricingStartedAt = 0;
  let pricingTimer = null;

  const REF_ACCEPTED = /\.(pdf|docx?|pptx?|md|markdown|txt)$/i;
  const MAX_REFS = 6;
  const KEY_STORAGE = 'office-agent.deepseek-key';
  const KEY_REMEMBER = 'office-agent.deepseek-key-remember';
  const THEME_STORAGE = 'office-agent.theme';
  const CHECKOUT_PENDING_STORAGE = 'whale-editor.pending-checkout';
  const QUOTE_PREVIEW_MS = 10 * 60 * 1000;

  // ---------- Theme ----------
  const themeMedia = window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : null;

  function normalizeThemeChoice(choice) {
    return ['system', 'light', 'dark'].includes(choice) ? choice : 'system';
  }

  function storedThemeChoice() {
    try {
      return normalizeThemeChoice(localStorage.getItem(THEME_STORAGE) || 'system');
    } catch {
      return 'system';
    }
  }

  function resolveTheme(choice) {
    return choice === 'dark' || (choice === 'system' && themeMedia?.matches) ? 'dark' : 'light';
  }

  function applyThemeChoice(choice, persist = false) {
    const normalized = normalizeThemeChoice(choice);
    const resolved = resolveTheme(normalized);
    document.documentElement.dataset.themePreference = normalized;
    document.documentElement.dataset.themeResolved = resolved;
    if (persist) {
      try { localStorage.setItem(THEME_STORAGE, normalized); } catch {}
    }
    themeButtons.forEach((button) => {
      const active = button.dataset.themeChoice === normalized;
      button.classList.toggle('active', active);
      button.setAttribute('aria-pressed', String(active));
    });
  }

  themeButtons.forEach((button) => {
    button.addEventListener('click', () => {
      applyThemeChoice(button.dataset.themeChoice || 'system', true);
    });
  });

  if (themeMedia) {
    const onThemeMediaChange = () => {
      if (storedThemeChoice() === 'system') applyThemeChoice('system');
    };
    if (themeMedia.addEventListener) themeMedia.addEventListener('change', onThemeMediaChange);
    else if (themeMedia.addListener) themeMedia.addListener(onThemeMediaChange);
  }

  applyThemeChoice(storedThemeChoice());

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
      setNotice('DeepSeek API Key 通常以 sk- 开头，请检查后再保存。', 'error', true);
      return;
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
    setNotice(v ? 'API Key 已保存。' : '已切换为使用服务器默认 key。', 'info');
    keyPanel.hidden = true;
    keyToggle.setAttribute('aria-expanded', 'false');
  });

  keyClear.addEventListener('click', () => {
    inMemoryKey = '';
    setStoredKey('');
    keyInput.value = '';
    refreshKeyState();
    setNotice('API Key 已清除。', 'info');
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
    progressCard.setAttribute('aria-busy', state === 'running' ? 'true' : 'false');
    setLiveState(label, state === 'error' ? 'error' : state === 'done' ? 'done' : 'running');
    scrollChat();
  }

  function setNotice(message, tone = 'info', focus = false) {
    if (!noticeRegion) return;
    noticeRegion.textContent = message;
    noticeRegion.hidden = false;
    noticeRegion.className = `notice-region ${tone}`;
    noticeRegion.setAttribute('role', tone === 'error' ? 'alert' : 'status');
    if (focus) noticeRegion.focus({ preventScroll: false });
  }

  function clearNotice() {
    if (!noticeRegion) return;
    noticeRegion.textContent = '';
    noticeRegion.hidden = true;
    noticeRegion.className = 'notice-region';
    noticeRegion.setAttribute('role', 'status');
  }

  function updateDownloadButtonState() {
    if (!btnDownload) return;
    const locked = !!(activeSession && !activeSession.paid);
    btnDownload.textContent = locked ? '解锁下载' : '下载修改后的文档';
    btnDownload.dataset.locked = locked ? 'true' : 'false';
  }

  function safeJsonParse(value) {
    try {
      return JSON.parse(value);
    } catch {
      return null;
    }
  }

  function checkoutPending() {
    try {
      return safeJsonParse(localStorage.getItem(CHECKOUT_PENDING_STORAGE) || '');
    } catch {
      return null;
    }
  }

  function saveCheckoutPending() {
    if (!activeSession?.id) return;
    try {
      localStorage.setItem(CHECKOUT_PENDING_STORAGE, JSON.stringify({
        sessionId: activeSession.id,
        filename: modifiedFilename || activeSession.filename || 'modified-document',
        downloadUrl: modifiedDownloadUrl,
        at: Date.now(),
      }));
    } catch {
      // Checkout can still continue without local restore metadata.
    }
  }

  function clearCheckoutPending() {
    try {
      localStorage.removeItem(CHECKOUT_PENDING_STORAGE);
    } catch {
      // Ignore storage failures.
    }
  }

  async function readJsonResponse(resp) {
    const text = await resp.text().catch(() => '');
    if (!text) return {};
    return safeJsonParse(text) || { error: text };
  }

  function triggerBlobDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename || 'modified-document';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
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
    const isDraftEditor = editorMode === 'edit' && !currentInspectableFile();
    if (editorMode !== 'edit') hideSelectionCard();
    documentPage.setAttribute('contenteditable', isDraftEditor ? 'true' : 'false');
    documentPage.setAttribute('role', isDraftEditor ? 'textbox' : 'document');
    documentPage.setAttribute('aria-label', isDraftEditor ? '文稿编辑区' : '文稿预览区');
    documentPage.setAttribute('aria-multiline', isDraftEditor ? 'true' : 'false');

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
    hideSelectionCard();
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
      <div class="doc-kicker">Whale Editor</div>
      <h1>拖入文档开始修改</h1>
      <p>DOCX / PPTX 会在这里展开预览。</p>
      <p>右侧输入修改要求，左侧同步显示可见变更。</p>
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
          return `<${tag} class="${paraClass(p, index)}" data-preview-id="${escapeHtml(p.id)}" tabindex="0">${marker}<span>${text}</span></${tag}>`;
        }).join('')
      : `<p class="doc-paragraph empty">未提取到可显示正文。文档结构仍可在“审阅”里查看。</p>`;

    const tablesHtml = (preview.tables || []).slice(0, 4).map((table) => `
      <table class="doc-table-preview" data-preview-id="${escapeHtml(table.id)}">
        <tbody>
          ${(table.rows || []).slice(0, 8).map((row) => `
            <tr>
              ${(row.cells || []).slice(0, 6).map((cell) => `
                <td class="${previewUpdatedIds.has(cell.id) ? 'doc-updated-line' : ''}" data-preview-id="${escapeHtml(cell.id)}" tabindex="0">${escapeHtml(cell.text || '')}</td>
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
          <section class="slide-preview" data-preview-id="${escapeHtml(slide.id)}" tabindex="0">
            <div class="slide-index">第 ${slide.index} 页</div>
            <h1>${escapeHtml(slide.title || `幻灯片 ${slide.index}`)}</h1>
            ${(slide.shapes || []).filter((shape) => shape.text).slice(0, 8).map((shape) => `
              <p class="${previewUpdatedIds.has(shape.id) ? 'doc-updated-line' : ''}" data-preview-id="${escapeHtml(shape.id)}" tabindex="0">${escapeHtml(shape.text)}</p>
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

    const rows = comments.map((comment, index) => {
      const anchor = (comment.anchoredParagraphIds || [])[0] || '';
      return `
      <button type="button" class="review-item review-jump" data-jump-preview-id="${escapeHtml(anchor)}">
        <div class="review-item-head">
          <span>${escapeHtml(comment.commentId || `comment-${index + 1}`)}</span>
          <span>${escapeHtml(comment.author || '未知作者')}</span>
        </div>
        <p>${escapeHtml(comment.textPreview)}</p>
        <div class="review-item-meta">${escapeHtml((comment.anchoredParagraphIds || []).join(' · ') || '未定位锚点')}</div>
      </button>
    `;
    }).join('');

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
          <button type="button" class="review-item review-jump" data-jump-preview-id="${escapeHtml(target.id)}">
            <div class="review-item-head">
              <span>${escapeHtml(target.id)}</span>
              <span>${target.commentIds ? `${target.commentIds.length} 批注` : target.hasNotes ? '有备注' : ''}</span>
            </div>
            <p>${escapeHtml(target.textPreview || '无文本预览')}</p>
          </button>
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
    clearNotice();
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
        renderPricingCard();
      }
    }
  }

  function applyQuickIntent(intent) {
    intentInput.value = intent;
    updateRunEnabled();
    intentInput.focus();
    setLiveState('等待执行', 'done');
    setNotice('已把指令填入右侧输入框。', 'info');
    scrollChat();
  }

  function findPreviewObject(id) {
    if (!id) return null;
    const para = previewParagraphs().find((p) => p.id === id);
    if (para) return { type: '段落', id, text: para.text || '', styleId: para.styleId, commentIds: para.commentIds || [] };
    for (const table of previewTables()) {
      for (const row of table.rows || []) {
        const cell = (row.cells || []).find((c) => c.id === id);
        if (cell) return { type: '表格单元格', id, text: cell.text || '', tableId: table.id };
      }
    }
    const slide = previewSlides().find((s) => s.id === id);
    if (slide) return { type: '幻灯片', id, text: slide.title || '', slideIndex: slide.index };
    for (const slideItem of previewSlides()) {
      const shape = (slideItem.shapes || []).find((s) => s.id === id);
      if (shape) return { type: '形状文本', id, text: shape.text || '', slideId: slideItem.id };
    }
    return { type: '文稿对象', id, text: '' };
  }

  function describePreviewContext(ctx) {
    const parts = [ctx.type, ctx.id];
    if (ctx.styleId) parts.push(`样式 ${ctx.styleId}`);
    if (ctx.commentIds?.length) parts.push(`${ctx.commentIds.length} 条批注`);
    if (ctx.tableId) parts.push(ctx.tableId);
    if (ctx.slideIndex) parts.push(`第 ${ctx.slideIndex} 页`);
    return parts.join(' · ');
  }

  function hideSelectionCard() {
    selectedPreviewContext = null;
    documentPage.querySelectorAll('.doc-selected-line').forEach((el) => el.classList.remove('doc-selected-line'));
    if (selectionCard) selectionCard.hidden = true;
  }

  function updateSelectionCard() {
    if (editorMode !== 'edit' || !documentInfo?.preview) {
      hideSelectionCard();
      return;
    }

    const selection = window.getSelection();
    if (!selection || selection.isCollapsed || selection.rangeCount === 0) {
      hideSelectionCard();
      return;
    }

    const range = selection.getRangeAt(0);
    const container = range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
      ? range.commonAncestorContainer
      : range.commonAncestorContainer.parentElement;
    if (!container || !documentPage.contains(container)) {
      hideSelectionCard();
      return;
    }

    const targetEl = container.closest?.('[data-preview-id]');
    const selectedText = selection.toString().trim();
    if (!targetEl || selectedText.length < 1) {
      hideSelectionCard();
      return;
    }

    const ctx = findPreviewObject(targetEl.dataset.previewId);
    if (!ctx) {
      hideSelectionCard();
      return;
    }

    documentPage.querySelectorAll('.doc-selected-line').forEach((el) => el.classList.remove('doc-selected-line'));
    targetEl.classList.add('doc-selected-line');
    selectedPreviewContext = { ...ctx, selectedText };
    selectionText.textContent = truncate(selectedText, 140);
    selectionMeta.textContent = describePreviewContext(selectedPreviewContext);
    selectionCard.hidden = false;
  }

  function useSelectionInIntent() {
    if (!selectedPreviewContext) return;
    const original = selectedPreviewContext.text || selectedPreviewContext.selectedText;
    const instruction = [
      `请只修改左侧选中的${selectedPreviewContext.type}。`,
      `目标 ID：${selectedPreviewContext.id}`,
      selectedPreviewContext.styleId ? `样式：${selectedPreviewContext.styleId}` : '',
      selectedPreviewContext.commentIds?.length ? `关联批注：${selectedPreviewContext.commentIds.join(', ')}` : '',
      `选中文本：${selectedPreviewContext.selectedText}`,
      original && original !== selectedPreviewContext.selectedText ? `对象原文：${original}` : '',
      '修改要求：',
    ].filter(Boolean).join('\n');
    intentInput.value = intentInput.value.trim()
      ? `${intentInput.value.trim()}\n\n${instruction}`
      : instruction;
    updateRunEnabled();
    intentInput.focus();
    setLiveState('已带入选区', 'done');
    setNotice('已把左侧选区上下文带入右侧输入框。', 'info');
  }

  function cssEscape(value) {
    return window.CSS?.escape ? window.CSS.escape(value) : String(value).replace(/["\\]/g, '\\$&');
  }

  function jumpToPreviewId(id) {
    if (!id) return;
    setEditorMode('edit');
    requestAnimationFrame(() => {
      const el = documentPage.querySelector(`[data-preview-id="${cssEscape(id)}"]`);
      if (!el) return;
      documentPage.querySelectorAll('.doc-selected-line').forEach((item) => item.classList.remove('doc-selected-line'));
      el.classList.add('doc-selected-line');
      el.scrollIntoView({ block: 'center', behavior: 'smooth' });
      el.focus({ preventScroll: true });
      setNotice(`已定位到 ${id}`, 'info');
    });
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
    if (!file) {
      selectedFile = null;
      followUpMode = false;
      modifiedBlob = null;
      modifiedFilename = null;
      modifiedDownloadUrl = null;
      clearPricingCard();
      clearDocumentInfo();
      dropzone.classList.remove('has-file');
      dzEmpty.hidden = false;
      dzFilled.hidden = true;
      resetDocumentPage();
      updateRunEnabled();
      return;
    }
    if (!ACCEPTED.test(file.name)) {
      setNotice('只支持 .docx / .pptx / .doc / .ppt 文件。', 'error', true);
      return;
    }
    if (file.size > 50 * 1024 * 1024) {
      setNotice('文件超过 50MB 上限。', 'error', true);
      return;
    }
    clearSession();
    selectedFile = file;
    fileNameEl.textContent = file.name;
    fileMetaEl.textContent = `${formatSize(file.size)} · ${file.name.split('.').pop().toLowerCase()}`;
    dropzone.classList.add('has-file');
    dzEmpty.hidden = true;
    dzFilled.hidden = false;
    pricingStartedAt = Date.now();
    startPricingTimer();
    updateDocumentPageForFile(file);
    setNotice(`已载入 ${file.name}。`, 'info');
    updateRunEnabled();
  }

  function hasDraggedFiles(e) {
    return Array.from(e.dataTransfer?.types || []).includes('Files');
  }

  function setStageDragging(active) {
    if (editorStage) editorStage.classList.toggle('document-dragging', active);
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

  function formatCountdown(ms) {
    const safe = Math.max(0, ms);
    const minutes = Math.floor(safe / 60000);
    const seconds = Math.floor((safe % 60000) / 1000);
    return `${minutes}:${String(seconds).padStart(2, '0')}`;
  }

  function documentStatsForPricing() {
    const source = currentInspectableFile();
    const summary = documentInfo?.summary || {};
    const fileType = documentInfo?.fileType || source?.name?.split('.').pop()?.toLowerCase() || 'docx';
    const paragraphs = Number(summary.paragraphs || 0);
    const tables = Number(summary.tables || 0);
    const comments = Number(summary.comments || 0);
    const slides = Number(summary.slides || 0);
    const fileSizeMb = source ? source.size / 1024 / 1024 : 0;
    const estimatedPages = fileType.includes('ppt')
      ? slides
      : Math.max(1, Math.ceil((paragraphs || 18) / 18 + tables * 0.6 + comments * 0.12));

    return {
      fileType,
      paragraphs,
      tables,
      comments,
      slides,
      fileSizeMb,
      estimatedPages,
      referenceCount: referenceFiles.length,
      ready: !!documentInfo?.summary,
    };
  }

  function choosePricingPlan(stats) {
    const unitLabel = stats.fileType.includes('ppt') ? '页演示' : '页文稿';
    const count = stats.fileType.includes('ppt') ? stats.slides || stats.estimatedPages : stats.estimatedPages;
    const custom = count > 120 || stats.slides > 80 || stats.fileSizeMb > 30;
    const deep = custom || count > 60 || stats.tables > 20 || stats.comments > 30 || stats.referenceCount > 3;
    const quick = !deep && count <= 20 && stats.tables <= 8 && stats.comments <= 10 && stats.referenceCount <= 1;

    if (custom) {
      return {
        tier: '超长复杂文档',
        price: '¥69 起',
        window: '上传后报价',
        rounds: '按报价',
        download: '24 小时',
        scope: `${count || '多'} ${unitLabel} · ${stats.tables} 表格 · ${stats.comments} 批注`,
        note: '文档规模较大，建议先生成结构报告，再确认最终价格。',
      };
    }

    if (deep) {
      return {
        tier: '深度修改',
        price: '¥39.9',
        window: '24 小时',
        rounds: '15 轮',
        download: '24 小时',
        scope: `${count || '多'} ${unitLabel} · ${stats.tables} 表格 · ${stats.comments} 批注`,
        note: '适合长文档、批注多、表格多或多参考材料的格式修订。',
      };
    }

    if (quick) {
      return {
        tier: '快速修改',
        price: '¥9.9',
        window: '1 小时',
        rounds: '3 轮',
        download: '24 小时',
        scope: `${count || '少量'} ${unitLabel} · ${stats.tables} 表格 · ${stats.comments} 批注`,
        note: '适合小文档快速改字、统一轻量格式或处理少量批注。',
      };
    }

    return {
      tier: '标准修改',
      price: '¥19.9',
      window: '2 小时',
      rounds: '8 轮',
      download: '24 小时',
      scope: `${count || '中等'} ${unitLabel} · ${stats.tables} 表格 · ${stats.comments} 批注`,
      note: '默认主力套餐，适合多数 Word / PPT 文档连续追改。',
    };
  }

  function renderPricingCard() {
    if (!pricingCard) return;
    const source = currentInspectableFile();
    if (!source) {
      pricingCard.hidden = true;
      return;
    }

    if (!pricingStartedAt) pricingStartedAt = Date.now();
    const stats = documentStatsForPricing();
    const plan = choosePricingPlan(stats);
    const remaining = pricingStartedAt + QUOTE_PREVIEW_MS - Date.now();

    pricingCard.hidden = false;
    pricingTier.textContent = plan.tier;
    pricingPrice.textContent = plan.price;
    pricingScope.textContent = stats.ready ? plan.scope : `${source.name} · 正在读取文档结构`;
    pricingWindow.textContent = plan.window;
    pricingRounds.textContent = plan.rounds;
    pricingDownload.textContent = plan.download;
    pricingCountdown.textContent = remaining > 0 ? formatCountdown(remaining) : '需重算';
    pricingNote.textContent = remaining > 0
      ? `${plan.note} 付款后开始计算修改窗口，底层缓存不作为用户权益保证。`
      : '报价预览已超过 10 分钟，建议重新读取文档后再确认价格。';
  }

  function startPricingTimer() {
    if (pricingTimer) clearInterval(pricingTimer);
    renderPricingCard();
    pricingTimer = setInterval(renderPricingCard, 1000);
  }

  function clearPricingCard() {
    pricingStartedAt = 0;
    if (pricingTimer) {
      clearInterval(pricingTimer);
      pricingTimer = null;
    }
    if (pricingCard) pricingCard.hidden = true;
  }

  // ---------- References ----------
  function addReferences(fileList) {
    for (const f of fileList) {
      if (referenceFiles.length >= MAX_REFS) {
        setNotice(`最多只能添加 ${MAX_REFS} 份参考材料。`, 'error', true);
        break;
      }
      if (!REF_ACCEPTED.test(f.name)) {
        setNotice(`不支持的参考材料格式：${f.name}`, 'error', true);
        continue;
      }
      if (f.size > 50 * 1024 * 1024) {
        setNotice(`参考材料 ${f.name} 超过 50MB。`, 'error', true);
        continue;
      }
      const dup = referenceFiles.find((r) => r.name === f.name && r.size === f.size);
      if (dup) continue;
      referenceFiles.push(f);
    }
    renderRefList();
    renderPricingCard();
    updateRunEnabled();
  }

  function refKind(filename) {
    const m = filename.toLowerCase().match(/\.(pdf|docx?|pptx?|md|markdown|txt)$/);
    if (!m) return '文件';
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
        renderPricingCard();
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
    const jump = e.target.closest('[data-jump-preview-id]');
    if (jump) {
      jumpToPreviewId(jump.dataset.jumpPreviewId || '');
      return;
    }
    const refresh = e.target.closest('[data-inspect-refresh]');
    if (refresh) inspectCurrentDocument(true);
  });

  documentPage.addEventListener('keyup', (e) => {
    if (e.key === 'Enter') {
      const jump = e.target.closest?.('[data-jump-preview-id]');
      if (jump) jumpToPreviewId(jump.dataset.jumpPreviewId || '');
    }
  });

  document.addEventListener('selectionchange', () => {
    window.requestAnimationFrame(updateSelectionCard);
  });

  selectionUse.addEventListener('mousedown', (e) => e.preventDefault());
  selectionUse.addEventListener('click', useSelectionInIntent);

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

  if (editorStage) {
    editorStage.addEventListener('dragenter', (e) => {
      if (!hasDraggedFiles(e)) return;
      e.preventDefault();
      setStageDragging(true);
    });

    editorStage.addEventListener('dragover', (e) => {
      if (!hasDraggedFiles(e)) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = 'copy';
      setStageDragging(true);
    });

    editorStage.addEventListener('dragleave', (e) => {
      if (!editorStage.contains(e.relatedTarget)) setStageDragging(false);
    });

    editorStage.addEventListener('drop', (e) => {
      if (!hasDraggedFiles(e)) return;
      e.preventDefault();
      setStageDragging(false);
      const f = e.dataTransfer.files[0];
      if (f) setFile(f);
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
    downloadUnlockedDocument();
  });

  async function downloadUnlockedDocument() {
    if (activeSession && !activeSession.paid) {
      await startCheckout();
      return;
    }

    if (modifiedDownloadUrl) {
      btnDownload.disabled = true;
      try {
        const resp = await fetch(modifiedDownloadUrl, { cache: 'no-store' });
        if (resp.status === 402) {
          activeSession = activeSession ? { ...activeSession, paid: false } : activeSession;
          updateDownloadButtonState();
          await startCheckout();
          return;
        }
        if (!resp.ok) {
          const data = await readJsonResponse(resp);
          throw new Error(data.error || `HTTP ${resp.status}`);
        }
        const blob = await resp.blob();
        triggerBlobDownload(blob, modifiedFilename || 'modified-document');
        setNotice('已开始下载修改后的文档。', 'info');
      } catch (err) {
        setNotice(`下载失败：${err.message || String(err)}`, 'error', true);
      } finally {
        btnDownload.disabled = false;
        updateDownloadButtonState();
      }
      return;
    }

    if (modifiedBlob) {
      triggerBlobDownload(modifiedBlob, modifiedFilename || 'modified-document');
      setNotice('已开始下载修改后的文档。', 'info');
      return;
    }

    setNotice('还没有可下载的修改稿。', 'error', true);
  }

  async function startCheckout() {
    if (!activeSession?.id) {
      setNotice('请先完成一次文档修改，再解锁下载。', 'error', true);
      return;
    }

    btnDownload.disabled = true;
    btnDownload.textContent = '打开支付...';

    try {
      const resp = await fetch('/api/checkout/session', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sessionId: activeSession.id }),
      });
      const data = await readJsonResponse(resp);
      if (!resp.ok) throw new Error(data.error || `HTTP ${resp.status}`);

      if (data.paid) {
        activeSession = data.session || { ...activeSession, paid: true };
        modifiedDownloadUrl = data.downloadUrl || modifiedDownloadUrl;
        modifiedFilename = data.modifiedFilename || modifiedFilename;
        startSessionCountdown();
        updateDownloadButtonState();
        await downloadUnlockedDocument();
        return;
      }

      if (data.url) {
        saveCheckoutPending();
        window.location.assign(data.url);
        return;
      }

      throw new Error('支付服务未返回可打开的结算页面。');
    } catch (err) {
      setNotice(`无法解锁下载：${err.message || String(err)}`, 'error', true);
    } finally {
      btnDownload.disabled = false;
      updateDownloadButtonState();
    }
  }

  function cleanCheckoutQuery() {
    const params = new URLSearchParams(window.location.search);
    params.delete('checkout_session_id');
    params.delete('document_session_id');
    params.delete('payment');
    const next = `${window.location.pathname}${params.toString() ? `?${params}` : ''}${window.location.hash}`;
    window.history.replaceState({}, document.title, next);
  }

  async function confirmCheckoutFromUrl() {
    const params = new URLSearchParams(window.location.search);
    const cancelled = params.get('payment') === 'cancelled';
    const checkoutSessionId = params.get('checkout_session_id');
    const pending = checkoutPending();

    if (cancelled) {
      clearCheckoutPending();
      cleanCheckoutQuery();
      setNotice('已取消付款，修改稿仍处于锁定状态。', 'info');
      return;
    }

    if (!checkoutSessionId) return;

    const documentSessionId = params.get('document_session_id') || pending?.sessionId || '';
    progressCard.hidden = false;
    finalActions.hidden = false;
    btnDownload.hidden = false;
    btnDownload.disabled = true;
    btnDownload.textContent = '确认支付...';
    setRunStatus('确认支付状态', 'running');
    setDocState('正在解锁下载', true);

    try {
      const query = new URLSearchParams({ checkout_session_id: checkoutSessionId });
      if (documentSessionId) query.set('document_session_id', documentSessionId);
      const resp = await fetch(`/api/checkout/confirm?${query}`, { cache: 'no-store' });
      const data = await readJsonResponse(resp);
      if (!resp.ok) throw new Error(data.error || `HTTP ${resp.status}`);

      if (!data.paid) {
        activeSession = data.session || activeSession;
        updateDownloadButtonState();
        setRunStatus('支付尚未完成', 'warn');
        setNotice('还没有收到付款完成状态，修改稿暂时不能下载。', 'error', true);
        return;
      }

      activeSession = data.session || activeSession;
      modifiedDownloadUrl = data.downloadUrl || pending?.downloadUrl || modifiedDownloadUrl;
      modifiedFilename = data.modifiedFilename || pending?.filename || modifiedFilename || 'modified-document';
      modifiedBlob = null;
      followUpMode = canReuseSession();
      startSessionCountdown();
      btnReedit.hidden = !canReuseSession();
      btnDownload.hidden = false;
      updateDownloadButtonState();
      editorMode = 'edit';
      updateDocumentPageForResult(true, false);
      await refreshModifiedPreviewFromSession();
      setRunStatus('已解锁下载', 'done');
      setNotice('支付完成，修改稿已解锁，可以下载。', 'info');
      clearCheckoutPending();
    } catch (err) {
      setRunStatus('支付确认失败', 'error');
      setNotice(`支付确认失败：${err.message || String(err)}`, 'error', true);
    } finally {
      btnDownload.disabled = false;
      updateDownloadButtonState();
      cleanCheckoutQuery();
      updateRunEnabled();
    }
  }

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
    icon.textContent = state === 'running' ? '进行' : state === 'done' ? '完成' : state === 'error' ? '错' : '待';
    icon.setAttribute('aria-label', state === 'running' ? '进行中' : state === 'done' ? '已完成' : state === 'error' ? '失败' : '待执行');
    if (meta) {
      const metaEl = li.querySelector('.step-meta');
      const parts = [];
      if (typeof meta.iters === 'number') parts.push(`${meta.iters} 轮`);
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
    if ('styleId' in p) return `<span class="to">样式：${escapeHtml(p.styleId)}</span>`;
    if ('x' in p && 'y' in p) return `<span class="to">移动到 (${p.x}, ${p.y})</span>`;
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
      mCacheSub.textContent = `${u.cacheHitTokens || 0} 命中 · ${u.cacheMissTokens || 0} 未命中`;
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

    updateDownloadButtonState();
    finalActions.hidden = false;
    followUpMode = canReuseSession();
    btnReedit.hidden = !canReuseSession();
    if (ev.success) editorMode = 'edit';
    updateDocumentPageForResult(ev.success, !!ev.modifiedFileBase64);
    if (ev.success && activeSession?.id) {
      refreshModifiedPreviewFromSession();
    } else if (ev.success && ev.modifiedFileBase64) {
      inspectCurrentDocument(true);
    }
    updateRunEnabled();
    scrollChat();
  }

  function showError(msg) {
    progressCard.hidden = false;
    errorBox.hidden = false;
    errorBox.textContent = msg;
    setNotice(msg, 'error');
    setRunStatus('失败', 'error');
    setDocState(selectedFile ? '修改失败' : '等待文档', !!selectedFile);
    scrollChat();
  }

  async function refreshModifiedPreviewFromSession() {
    if (!activeSession?.id) return;
    setDocState('正在刷新修改稿预览', true);
    const requestId = ++inspectRequestId;
    inspectingDocument = true;
    inspectError = '';
    renderDocumentWorkspace();
    try {
      const resp = await fetch(`/api/session/${encodeURIComponent(activeSession.id)}/inspect`, { cache: 'no-store' });
      if (!resp.ok) {
        const data = await readJsonResponse(resp);
        throw new Error(data.error || `HTTP ${resp.status}`);
      }
      const data = await resp.json();
      if (requestId !== inspectRequestId) return;
      documentInfo = data;
      documentInfoKey = `session|${activeSession.id}|${Date.now()}`;
      inspectError = '';
      previewUpdatedIds = new Set();
      editorViewState = {
        kind: 'result',
        name: modifiedFilename || data.filename || activeSession.filename || 'document',
        success: true,
      };
      editorMode = 'edit';
      setDocState('已刷新修改稿预览', true);
    } catch (err) {
      if (requestId !== inspectRequestId) return;
      console.warn('Unable to refresh modified preview:', err);
      setDocState('已生成修改稿', true);
      inspectError = err instanceof Error ? err.message : String(err);
    } finally {
      if (requestId === inspectRequestId) {
        inspectingDocument = false;
        renderDocumentWorkspace();
        renderPricingCard();
      }
    }
  }

  resetDocumentPage();
  setLiveState('空闲');
  confirmCheckoutFromUrl();
})();
