/* ========== Constants ========== */
const API_BASE = 'https://generativelanguage.googleapis.com';
const MODEL_CHAT = 'gemini-2.5-flash';
const MODEL_ADVANCED = 'gemini-2.5-flash';
const STORAGE_KEY_API = 'gemini_api_key';
const STORAGE_KEY_SEARCH = 'gemini_use_search';
const STORAGE_KEY_TRACK = 'gemini_track_changes';

/* ========== State ========== */
let conversationHistory = []; // { role: 'user'|'model', parts: [{text}] }[]
let currentPreviewText = '';
let currentPreviewMode = 'insert'; // 'insert' | 'replace-body'
let isProcessing = false;
let pendingDocFile = null;
let docUpdateMode = null; // 'full' | 'partial'

/* ========== Office.js Init ========== */
Office.onReady(function (info) {
  const loadingScreen = document.getElementById('loading-screen');
  const mainContent = document.getElementById('main-content');

  if (info.host === Office.HostType.Word) {
    loadingScreen.style.display = 'none';
    mainContent.classList.remove('hidden');
    initApp();
  } else {
    loadingScreen.innerHTML =
      '<p style="color:#a4262c;padding:20px;text-align:center">이 Add-in은 Microsoft Word 전용입니다.</p>';
  }
});

/* ========== App Initialization ========== */
function initApp() {
  loadSettings();
  checkApiKey();
  bindEvents();
}

function loadSettings() {
  const key = localStorage.getItem(STORAGE_KEY_API) || '';
  const search = localStorage.getItem(STORAGE_KEY_SEARCH) === 'true';
  const track = localStorage.getItem(STORAGE_KEY_TRACK) === 'true';
  document.getElementById('api-key-input').value = key;
  document.getElementById('google-search-toggle').checked = search;
  document.getElementById('track-changes-toggle').checked = track;
}

function saveSettings() {
  const key = document.getElementById('api-key-input').value.trim();
  const search = document.getElementById('google-search-toggle').checked;
  const track = document.getElementById('track-changes-toggle').checked;
  localStorage.setItem(STORAGE_KEY_API, key);
  localStorage.setItem(STORAGE_KEY_SEARCH, search);
  localStorage.setItem(STORAGE_KEY_TRACK, track);
  closeSettings();
  checkApiKey();
  showToast('설정이 저장되었습니다.', 'success');
}

function getApiKey() { return localStorage.getItem(STORAGE_KEY_API) || ''; }
function useSearch() { return localStorage.getItem(STORAGE_KEY_SEARCH) === 'true'; }
function useTrackChanges() { return localStorage.getItem(STORAGE_KEY_TRACK) === 'true'; }

function checkApiKey() {
  const warning = document.getElementById('api-warning');
  if (getApiKey()) {
    warning.classList.add('hidden');
  } else {
    warning.classList.remove('hidden');
  }
}

/* ========== Event Binding ========== */
function bindEvents() {
  // Header
  document.getElementById('settings-btn').addEventListener('click', toggleSettings);
  document.getElementById('clear-btn').addEventListener('click', clearConversation);

  // Settings
  document.getElementById('save-settings-btn').addEventListener('click', saveSettings);
  document.getElementById('close-settings-btn').addEventListener('click', closeSettings);
  document.getElementById('open-settings-from-warning').addEventListener('click', openSettings);
  document.getElementById('toggle-api-visibility').addEventListener('click', toggleApiVisibility);

  // Chat
  document.getElementById('send-btn').addEventListener('click', sendMessage);
  document.getElementById('message-input').addEventListener('keydown', onInputKeydown);
  document.getElementById('get-selection-btn').addEventListener('click', getWordSelection);

  // Action buttons
  document.getElementById('book-format-btn').addEventListener('click', formatAsBook);
  document.getElementById('voice-upload-btn').addEventListener('click', () => {
    document.getElementById('audio-file-input').click();
  });
  document.getElementById('audio-file-input').addEventListener('change', onAudioFileSelected);
  document.getElementById('doc-upload-btn').addEventListener('click', () => {
    document.getElementById('doc-file-input').click();
  });
  document.getElementById('doc-file-input').addEventListener('change', onDocumentFileSelected);

  // Document update dialog
  document.getElementById('doc-full-replace-btn').addEventListener('click', () => selectDocUpdateMode('full'));
  document.getElementById('doc-partial-update-btn').addEventListener('click', () => selectDocUpdateMode('partial'));
  document.getElementById('doc-confirm-btn').addEventListener('click', confirmDocUpdate);
  document.getElementById('doc-cancel-btn').addEventListener('click', cancelDocUpdate);
  document.getElementById('doc-cancel-close-btn').addEventListener('click', cancelDocUpdate);

  // Preview panel
  document.getElementById('close-preview-btn').addEventListener('click', closePreview);
  document.getElementById('cancel-preview-btn').addEventListener('click', closePreview);
  document.getElementById('insert-preview-btn').addEventListener('click', insertPreviewToDocument);
}

/* ========== Settings UI ========== */
function toggleSettings() {
  const panel = document.getElementById('settings-panel');
  panel.classList.toggle('hidden');
}
function openSettings() {
  document.getElementById('settings-panel').classList.remove('hidden');
}
function closeSettings() {
  document.getElementById('settings-panel').classList.add('hidden');
}
function toggleApiVisibility() {
  const input = document.getElementById('api-key-input');
  input.type = input.type === 'password' ? 'text' : 'password';
}

/* ========== Chat ========== */
function onInputKeydown(e) {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    sendMessage();
  }
}

async function sendMessage() {
  if (isProcessing) return;
  const input = document.getElementById('message-input');
  const text = input.value.trim();
  if (!text) return;
  if (!getApiKey()) {
    openSettings();
    showToast('API 키를 먼저 설정해주세요.', 'warning');
    return;
  }

  input.value = '';
  autoResizeTextarea(input);
  isProcessing = true;
  setSendButtonState(false);

  addMessageToUI('user', text);
  conversationHistory.push({ role: 'user', parts: [{ text }] });

  showTypingIndicator();

  try {
    const responseText = await callGeminiAPI(conversationHistory, MODEL_CHAT, useSearch());
    hideTypingIndicator();
    addMessageToUI('assistant', responseText, true);
    conversationHistory.push({ role: 'model', parts: [{ text: responseText }] });
  } catch (err) {
    hideTypingIndicator();
    addErrorMessage(err.message);
  } finally {
    isProcessing = false;
    setSendButtonState(true);
    scrollChatToBottom();
  }
}

function addMessageToUI(role, text, showActions = false) {
  const messages = document.getElementById('messages');

  // Remove welcome message on first real message
  const welcome = messages.querySelector('.welcome-msg');
  if (welcome) welcome.remove();

  const wrapper = document.createElement('div');
  wrapper.className = `message ${role}`;

  const bubble = document.createElement('div');
  bubble.className = 'bubble';
  bubble.innerHTML = renderMarkdown(text);
  wrapper.appendChild(bubble);

  const time = document.createElement('div');
  time.className = 'msg-time';
  time.textContent = formatTime(new Date());
  wrapper.appendChild(time);

  if (role === 'assistant' && showActions) {
    const actions = document.createElement('div');
    actions.className = 'msg-actions';

    const copyBtn = makeActionBtn('📋 복사', () => copyToClipboard(text));
    const insertBtn = makeActionBtn('📄 삽입', () => insertTextToDocument(text));
    const replaceBtn = makeActionBtn('✏️ 선택 교체', () => replaceSelectionInDocument(text));

    actions.append(copyBtn, insertBtn, replaceBtn);
    wrapper.appendChild(actions);
  }

  messages.appendChild(wrapper);
  scrollChatToBottom();
}

function addErrorMessage(msg) {
  const messages = document.getElementById('messages');
  const wrapper = document.createElement('div');
  wrapper.className = 'message error';
  const bubble = document.createElement('div');
  bubble.className = 'bubble';
  bubble.textContent = `⚠️ 오류: ${msg}`;
  wrapper.appendChild(bubble);
  messages.appendChild(wrapper);
  scrollChatToBottom();
}

function makeActionBtn(label, onClick) {
  const btn = document.createElement('button');
  btn.className = 'msg-action-btn';
  btn.textContent = label;
  btn.addEventListener('click', onClick);
  return btn;
}

function clearConversation() {
  if (!confirm('대화 내용을 모두 지우시겠습니까?')) return;
  conversationHistory = [];
  const messages = document.getElementById('messages');
  messages.innerHTML = `
    <div class="welcome-msg">
      <div class="welcome-icon">✨</div>
      <p><strong>Gemini AI Assistant</strong></p>
      <p>대화가 초기화되었습니다. 새로운 대화를 시작해보세요.</p>
    </div>`;
}

/* ========== Word Selection ========== */
async function getWordSelection() {
  try {
    const text = await getSelectedText();
    if (!text || !text.trim()) {
      showToast('Word에서 텍스트를 선택해주세요.', 'info');
      return;
    }
    const input = document.getElementById('message-input');
    const cursor = input.selectionStart;
    const before = input.value.substring(0, cursor);
    const after = input.value.substring(cursor);
    const insertion = (before ? '\n\n' : '') + `[선택된 텍스트]\n${text.trim()}\n\n`;
    input.value = before + insertion + after;
    input.focus();
    autoResizeTextarea(input);
    showToast('선택한 텍스트가 입력창에 추가되었습니다.', 'info');
  } catch (err) {
    showToast('텍스트를 가져오지 못했습니다.', 'error');
  }
}

function getSelectedText() {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Text,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(result.error.message));
        }
      }
    );
  });
}

/* ========== Document Insertion ========== */
function supportsWordApi(version) {
  return !!(
    typeof Office !== 'undefined' &&
    Office.context.requirements &&
    Office.context.requirements.isSetSupported('WordApi', version || '1.1')
  );
}

async function insertTextToDocument(text) {
  try {
    await insertMarkdownAsWordFormatting(text, false, useTrackChanges());
    showToast('Word에 삽입되었습니다.', 'success');
  } catch (err) {
    showToast(`삽입 실패: ${err.message}`, 'error', 5000);
    throw err;
  }
}

async function replaceSelectionInDocument(text) {
  try {
    await insertMarkdownAsWordFormatting(text, true, useTrackChanges());
    showToast('선택 텍스트가 교체되었습니다.', 'success');
  } catch (err) {
    showToast(`교체 실패: ${err.message}`, 'error', 5000);
    throw err;
  }
}

// Word Online 호환: body.insertText로 문서 끝에 삽입
async function insertToWordDocument(text) {
  if (supportsWordApi('1.1')) {
    await Word.run(async (context) => {
      context.document.body.insertText(text, Word.InsertLocation.end);
      await context.sync();
    });
  } else {
    await setSelectedData(text);
  }
}

// Word Online 호환: 선택 영역을 replace로 교체
async function replaceWordSelection(text) {
  if (supportsWordApi('1.1')) {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(text, Word.InsertLocation.replace);
      await context.sync();
    });
  } else {
    await setSelectedData(text);
  }
}

async function insertMarkdownAsWordFormatting(markdownText, replaceSelection = false, trackChanges = false) {
  const paragraphs = parseMarkdownToWordParagraphs(markdownText);
  console.log('[Insert] 단락 파싱 완료:', paragraphs.length, '개', paragraphs.map(p => p.type));
  if (paragraphs.length === 0) return;

  await Word.run(async (context) => {

    // 1. 트랙 변경 모드 (삽입 후 원래 상태로 복구)
    let originalTrackingMode = null;
    if (trackChanges) {
      try {
        context.document.load('changeTrackingMode');
        await context.sync();
        originalTrackingMode = context.document.changeTrackingMode;
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await context.sync();
        console.log('[Insert] 트랙 변경 모드 활성화 (원래 모드:', originalTrackingMode, ')');
      } catch (e) {
        console.error('[Insert] 트랙 변경 모드 오류:', e.message, e.code);
      }
    }

    // 2. 삽입 기준점 결정
    //    [수정] Range(collapsed) 대신 Paragraph를 anchor로 사용
    //    → Word Online에서 Range.insertParagraph('', after)는 InvalidArgument 유발
    let ref; // Paragraph object — insertParagraph의 안정적 anchor

    if (replaceSelection) {
      // [수정] selection.delete() 대신 insertText('', replace) 사용
      //        → 삭제 후 cursor가 더 명확한 Paragraph 위치에 남음
      const sel = context.document.getSelection();
      sel.insertText('', Word.InsertLocation.replace);
      await context.sync();
      console.log('[Insert] 선택 영역 교체(삭제) 성공');
      // 삭제 후 커서가 위치한 단락을 Paragraph로 가져옴
      ref = context.document.getSelection().paragraphs.getFirst();
    } else {
      // 비교체 모드: 커서가 있는 단락을 Paragraph로 가져옴
      ref = context.document.getSelection().paragraphs.getFirst();
    }

    // 3. 스타일 맵
    //    [수정] 'List Bullet' / 'List Number' 제거
    //           → 해당 스타일이 문서에 없으면 sync 시점에 InvalidArgument 발생
    //           → 'Normal' + 수동 프리픽스(•, 1.) 방식으로 대체
    const styleMap = {
      h1: 'Heading 1',
      h2: 'Heading 2',
      h3: 'Heading 3',
      h4: 'Heading 4',
      bullet:     'Normal',  // 'List Bullet' 제거
      number:     'Normal',  // 'List Number' 제거
      normal:     'Normal',
      code:       'Normal',
      blockquote: 'Normal',
      empty:      'Normal',
    };

    let listCounter = 0; // 번호 목록 카운터
    let prevType = null;

    for (let i = 0; i < paragraphs.length; i++) {
      const para = paragraphs[i];
      console.log(`[Insert] 단락 ${i} 처리 중 (type: ${para.type})`);

      // 번호 목록 카운터 관리 (연속된 number 타입 추적)
      if (para.type === 'number') {
        if (prevType !== 'number') listCounter = 0;
        listCounter++;
      } else {
        listCounter = 0;
      }
      prevType = para.type;

      // 4. 단락 삽입
      //    [수정] ref가 Paragraph이므로 Paragraph.insertParagraph 호출
      //           → Range.insertParagraph보다 Word Online에서 안정적
      let newPara;
      try {
        newPara = ref.insertParagraph('', Word.InsertLocation.after);
        console.log(`[Insert] 단락 ${i} insertParagraph 성공`);
      } catch (e) {
        console.error(`[Insert] 단락 ${i} insertParagraph 오류:`, e.message, e.code, e.debugInfo);
        throw e;
      }

      // 5. 스타일 적용 (Heading 1~4, Normal만 사용 — 안전한 built-in 스타일)
      newPara.style = styleMap[para.type] || 'Normal';

      // 6. 내용 삽입
      try {
        if (para.type === 'hr') {
          const hrRange = newPara.insertText('─'.repeat(40), Word.InsertLocation.end);
          try { hrRange.font.color = '#AAAAAA'; } catch (_) {}

        } else if (para.type === 'bullet') {
          // [수정] 'List Bullet' 스타일 대신 수동 불릿 프리픽스 삽입
          newPara.insertText('• ', Word.InsertLocation.end);
          for (const run of para.runs) {
            if (!run.text) continue;
            const r = newPara.insertText(run.text, Word.InsertLocation.end);
            try {
              if (run.bold)  r.font.bold = true;
              if (run.italic) r.font.italic = true;
              if (run.code) { r.font.name = 'Courier New'; r.font.size = 10; }
            } catch (_) {}
          }

        } else if (para.type === 'number') {
          // [수정] 'List Number' 스타일 대신 수동 번호 프리픽스 삽입
          newPara.insertText(`${listCounter}. `, Word.InsertLocation.end);
          for (const run of para.runs) {
            if (!run.text) continue;
            const r = newPara.insertText(run.text, Word.InsertLocation.end);
            try {
              if (run.bold)  r.font.bold = true;
              if (run.italic) r.font.italic = true;
              if (run.code) { r.font.name = 'Courier New'; r.font.size = 10; }
            } catch (_) {}
          }

        } else if (para.type === 'code') {
          try { newPara.font.name = 'Courier New'; newPara.font.size = 10; } catch (_) {}
          for (const run of para.runs) {
            if (run.text) newPara.insertText(run.text, Word.InsertLocation.end);
          }

        } else if (para.type === 'blockquote') {
          try { newPara.leftIndent = 36; } catch (_) {}
          for (const run of para.runs) {
            if (!run.text) continue;
            const r = newPara.insertText(run.text, Word.InsertLocation.end);
            try { r.font.italic = true; } catch (_) {}
            if (run.bold) { try { r.font.bold = true; } catch (_) {} }
          }

        } else if (para.type !== 'empty') {
          for (const run of para.runs) {
            if (!run.text) continue;
            const r = newPara.insertText(run.text, Word.InsertLocation.end);
            try {
              if (run.bold)  r.font.bold = true;
              if (run.italic) r.font.italic = true;
              if (run.code) { r.font.name = 'Courier New'; r.font.size = 10; }
            } catch (e) {
              console.warn(`[Insert] 단락 ${i} font 설정 오류:`, e.message, 'run:', run);
            }
          }
        }
        console.log(`[Insert] 단락 ${i} 내용 삽입 성공`);
      } catch (e) {
        console.error(`[Insert] 단락 ${i} 내용 삽입 오류:`, e.message, e.code, 'para:', para);
        throw e;
      }

      ref = newPara;
    }

    // 7. 트랙 변경 모드 원래 상태로 복구
    if (trackChanges && originalTrackingMode !== null) {
      try {
        context.document.changeTrackingMode = originalTrackingMode;
        console.log('[Insert] 트랙 변경 모드 복구:', originalTrackingMode);
      } catch (e) {
        console.warn('[Insert] 트랙 변경 모드 복구 오류:', e.message);
      }
    }

    // 8. 최종 sync
    try {
      await context.sync();
      console.log('[Insert] 최종 context.sync() 성공');
    } catch (e) {
      console.error('[Insert] 최종 context.sync() 오류:', e.message, e.code, e.debugInfo);
      throw e;
    }
  });
}

function parseMarkdownToWordParagraphs(text) {
  const lines = text.split('\n');
  const result = [];
  let inCodeBlock = false;
  let codeLines = [];

  for (const line of lines) {
    if (line.trim().startsWith('```')) {
      if (inCodeBlock) {
        for (const cl of codeLines) {
          result.push({ type: 'code', runs: [{ text: cl, bold: false, italic: false, code: false }] });
        }
        codeLines = [];
        inCodeBlock = false;
      } else {
        inCodeBlock = true;
      }
      continue;
    }

    if (inCodeBlock) { codeLines.push(line); continue; }

    if (!line.trim()) {
      if (result.length > 0 && result[result.length - 1].type !== 'empty') {
        result.push({ type: 'empty', runs: [] });
      }
      continue;
    }

    if (/^(---+|===+|\*\*\*+)$/.test(line.trim())) {
      result.push({ type: 'hr', runs: [] });
      continue;
    }

    const h4 = line.match(/^#### (.+)$/);
    if (h4) { result.push({ type: 'h4', runs: parseInlineMarkdown(h4[1]) }); continue; }

    const h3 = line.match(/^### (.+)$/);
    if (h3) { result.push({ type: 'h3', runs: parseInlineMarkdown(h3[1]) }); continue; }

    const h2 = line.match(/^## (.+)$/);
    if (h2) { result.push({ type: 'h2', runs: parseInlineMarkdown(h2[1]) }); continue; }

    const h1 = line.match(/^# (.+)$/);
    if (h1) { result.push({ type: 'h1', runs: parseInlineMarkdown(h1[1]) }); continue; }

    const taskUnchecked = line.match(/^[-*+] \[ \] (.+)$/);
    if (taskUnchecked) {
      result.push({ type: 'bullet', runs: [{ text: '☐ ', bold: false, italic: false, code: false }, ...parseInlineMarkdown(taskUnchecked[1])] });
      continue;
    }

    const taskChecked = line.match(/^[-*+] \[x\] (.+)$/i);
    if (taskChecked) {
      result.push({ type: 'bullet', runs: [{ text: '☑ ', bold: false, italic: false, code: false }, ...parseInlineMarkdown(taskChecked[1])] });
      continue;
    }

    const bullet = line.match(/^[-*+] (.+)$/);
    if (bullet) { result.push({ type: 'bullet', runs: parseInlineMarkdown(bullet[1]) }); continue; }

    const ordered = line.match(/^\d+\. (.+)$/);
    if (ordered) { result.push({ type: 'number', runs: parseInlineMarkdown(ordered[1]) }); continue; }

    const blockquote = line.match(/^> (.+)$/);
    if (blockquote) { result.push({ type: 'blockquote', runs: parseInlineMarkdown(blockquote[1]) }); continue; }

    result.push({ type: 'normal', runs: parseInlineMarkdown(line) });
  }

  if (inCodeBlock) {
    for (const cl of codeLines) {
      result.push({ type: 'code', runs: [{ text: cl, bold: false, italic: false, code: false }] });
    }
  }

  while (result.length > 0 && result[result.length - 1].type === 'empty') result.pop();

  return result;
}

function parseInlineMarkdown(text) {
  const runs = [];
  let rem = text;

  while (rem.length > 0) {
    const bim = rem.match(/^\*\*\*(.+?)\*\*\*/);
    if (bim) { runs.push({ text: bim[1], bold: true, italic: true, code: false }); rem = rem.slice(bim[0].length); continue; }

    const bm = rem.match(/^\*\*(.+?)\*\*/);
    if (bm) { runs.push({ text: bm[1], bold: true, italic: false, code: false }); rem = rem.slice(bm[0].length); continue; }

    const im = rem.match(/^\*(.+?)\*/);
    if (im) { runs.push({ text: im[1], bold: false, italic: true, code: false }); rem = rem.slice(im[0].length); continue; }

    const cm = rem.match(/^`([^`]+)`/);
    if (cm) { runs.push({ text: cm[1], bold: false, italic: false, code: true }); rem = rem.slice(cm[0].length); continue; }

    const next = rem.search(/\*\*\*|\*\*|\*|`/);
    if (next === -1) {
      runs.push({ text: rem, bold: false, italic: false, code: false });
      rem = '';
    } else if (next > 0) {
      runs.push({ text: rem.slice(0, next), bold: false, italic: false, code: false });
      rem = rem.slice(next);
    } else {
      runs.push({ text: rem[0], bold: false, italic: false, code: false });
      rem = rem.slice(1);
    }
  }

  return runs.length > 0 ? runs : [{ text: text, bold: false, italic: false, code: false }];
}

// 삽입 에러의 debugInfo를 미리보기 패널에 표시
function showPreviewInsertError(err) {
  const lines = [`⚠️ 삽입 실패: ${err.message}`];
  if (err.debugInfo) {
    lines.push('', '[디버그 정보]');
    if (err.debugInfo.message) lines.push(`message: ${err.debugInfo.message}`);
    if (err.debugInfo.errorLocation) lines.push(`location: ${err.debugInfo.errorLocation}`);
    if (err.debugInfo.innerError) {
      lines.push(`innerError: ${JSON.stringify(err.debugInfo.innerError)}`);
    }
  }
  document.getElementById('preview-title').textContent = '⚠️ 삽입 오류';
  document.getElementById('preview-content').textContent = lines.join('\n');
  showToast(`삽입 실패: ${err.message}`, 'error', 5000);
}

/* ========== 단행본 형식으로 정리 ========== */
async function formatAsBook() {
  if (conversationHistory.length === 0) {
    showToast('정리할 대화 내용이 없습니다.', 'warning');
    return;
  }
  if (!getApiKey()) {
    openSettings();
    showToast('API 키를 먼저 설정해주세요.', 'warning');
    return;
  }

  const convText = conversationHistory.map(msg => {
    const role = msg.role === 'user' ? '사용자' : 'AI';
    const text = msg.parts.map(p => p.text || '').join('');
    return `${role}: ${text}`;
  }).join('\n\n');

  const prompt =
    `아래는 특정 주제에 대한 대화 내용입니다.\n` +
    `이 대화를 단행본 형식의 글로 정리해주세요.\n` +
    `챕터와 소제목 구조로 구성하고, 자연스러운 문어체로 작성해주세요.\n\n` +
    `[대화 내용]\n${convText}`;

  showLoadingOverlay('단행본 형식으로 정리 중...');

  try {
    const result = await callGeminiAPI(
      [{ role: 'user', parts: [{ text: prompt }] }],
      MODEL_ADVANCED,
      false
    );
    hideLoadingOverlay();
    showPreview(result, '📖 단행본 형식 정리');
  } catch (err) {
    hideLoadingOverlay();
    showToast(`오류: ${err.message}`, 'error');
  }
}

/* ========== 음성 파일 업로드 → 회의록 ========== */
function onAudioFileSelected(e) {
  const file = e.target.files[0];
  e.target.value = '';
  if (file) processAudioFile(file);
}

async function processAudioFile(file) {
  if (!getApiKey()) {
    openSettings();
    showToast('API 키를 먼저 설정해주세요.', 'warning');
    return;
  }

  const SUPPORTED = ['audio/mpeg', 'audio/wav', 'audio/mp4', 'audio/flac', 'audio/x-m4a'];
  const ext = file.name.split('.').pop().toLowerCase();
  const supportedExt = ['mp3', 'wav', 'm4a', 'flac'];
  if (!SUPPORTED.includes(file.type) && !supportedExt.includes(ext)) {
    showToast('지원 형식: MP3, WAV, M4A, FLAC', 'warning');
    return;
  }

  const sizeMB = (file.size / 1024 / 1024).toFixed(1);
  showLoadingOverlay(`파일 업로드 중... (${sizeMB}MB)`);

  try {
    // Step 1: Upload to Gemini Files API
    const formData = new FormData();
    const metadata = { file: { displayName: file.name } };
    formData.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
    formData.append('file', file);

    const uploadResp = await fetch(
      `${API_BASE}/upload/v1beta/files?key=${getApiKey()}`,
      {
        method: 'POST',
        headers: { 'X-Goog-Upload-Protocol': 'multipart' },
        body: formData
      }
    );

    if (!uploadResp.ok) {
      const errData = await uploadResp.json().catch(() => ({}));
      throw new Error(errData.error?.message || `업로드 실패 (${uploadResp.status})`);
    }

    const { file: uploadedFile } = await uploadResp.json();

    // Step 2: Wait for file to be ACTIVE (polling)
    updateLoadingText('파일 처리 대기 중...');
    await waitForFileActive(uploadedFile.name);

    updateLoadingText('회의록 생성 중...');

    // Step 3: Generate meeting minutes
    const prompt =
      `다음 음성 파일은 회의 녹음입니다.\n` +
      `아래 형식으로 회의록을 작성해주세요:\n\n` +
      `# 회의록\n` +
      `- 일시: (파악 가능한 경우)\n` +
      `- 참석자: (파악 가능한 경우)\n\n` +
      `## 주요 논의사항\n` +
      `(내용 정리)\n\n` +
      `## 결정사항\n` +
      `(결정된 내용)\n\n` +
      `## 액션아이템\n` +
      `- [ ] 담당자: 내용 (기한이 언급된 경우 포함)`;

    const analysisResp = await fetch(
      `${API_BASE}/v1beta/models/${MODEL_ADVANCED}:generateContent?key=${getApiKey()}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{
            parts: [
              { fileData: { mimeType: uploadedFile.mimeType, fileUri: uploadedFile.uri } },
              { text: prompt }
            ]
          }],
          generationConfig: { temperature: 0.3, maxOutputTokens: 8192 }
        })
      }
    );

    if (!analysisResp.ok) {
      const errData = await analysisResp.json().catch(() => ({}));
      throw new Error(errData.error?.message || `분석 실패 (${analysisResp.status})`);
    }

    const data = await analysisResp.json();
    const content = data.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!content) throw new Error('AI 응답을 받지 못했습니다.');

    hideLoadingOverlay();
    showPreview(content, '🎙️ 회의록');

  } catch (err) {
    hideLoadingOverlay();
    showToast(`오류: ${err.message}`, 'error');
  }
}

async function waitForFileActive(fileName, maxWaitMs = 120000) {
  const interval = 3000;
  const start = Date.now();
  while (Date.now() - start < maxWaitMs) {
    await sleep(interval);
    const resp = await fetch(
      `${API_BASE}/v1beta/${fileName}?key=${getApiKey()}`
    );
    if (!resp.ok) break;
    const data = await resp.json();
    if (data.state === 'ACTIVE') return;
    if (data.state === 'FAILED') throw new Error('파일 처리 실패');
  }
}

/* ========== 문서 파일 업로드 → 업데이트 ========== */
function onDocumentFileSelected(e) {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;

  const supportedExt = ['pdf', 'docx', 'txt'];
  const ext = file.name.split('.').pop().toLowerCase();
  if (!supportedExt.includes(ext)) {
    showToast('지원 형식: PDF, DOCX, TXT', 'warning');
    return;
  }
  if (!getApiKey()) { openSettings(); showToast('API 키를 먼저 설정해주세요.', 'warning'); return; }

  pendingDocFile = file;
  docUpdateMode = null;
  document.getElementById('doc-update-filename').textContent = `📎 ${file.name}`;
  document.getElementById('doc-partial-instruction').classList.add('hidden');
  document.getElementById('doc-confirm-btn').classList.add('hidden');
  document.querySelectorAll('.doc-mode-btn').forEach(b => b.classList.remove('selected'));
  document.getElementById('doc-update-dialog').classList.remove('hidden');
}

function selectDocUpdateMode(mode) {
  docUpdateMode = mode;
  document.querySelectorAll('.doc-mode-btn').forEach(b => b.classList.remove('selected'));
  document.getElementById(mode === 'full' ? 'doc-full-replace-btn' : 'doc-partial-update-btn').classList.add('selected');
  document.getElementById('doc-partial-instruction').classList.toggle('hidden', mode !== 'partial');
  document.getElementById('doc-confirm-btn').classList.remove('hidden');
}

function cancelDocUpdate() {
  pendingDocFile = null;
  docUpdateMode = null;
  document.getElementById('doc-update-dialog').classList.add('hidden');
}

async function confirmDocUpdate() {
  if (!pendingDocFile || !docUpdateMode) return;
  const instruction = docUpdateMode === 'partial'
    ? document.getElementById('doc-instruction-input').value.trim()
    : '';
  if (docUpdateMode === 'partial' && !instruction) {
    showToast('업데이트할 내용을 설명해주세요.', 'warning');
    return;
  }
  document.getElementById('doc-update-dialog').classList.add('hidden');
  const file = pendingDocFile;
  const mode = docUpdateMode;
  pendingDocFile = null;
  docUpdateMode = null;
  await processDocumentUpdate(file, mode, instruction);
}

async function processDocumentUpdate(file, mode, instruction) {
  showLoadingOverlay('현재 문서 읽는 중...');
  try {
    const currentText = await getCurrentDocumentText();
    if (!currentText.trim()) throw new Error('현재 문서가 비어 있습니다.');

    updateLoadingText(`파일 업로드 중... (${(file.size / 1024 / 1024).toFixed(1)}MB)`);
    const formData = new FormData();
    formData.append('metadata', new Blob([JSON.stringify({ file: { displayName: file.name } })], { type: 'application/json' }));
    formData.append('file', file);

    const uploadResp = await fetch(
      `${API_BASE}/upload/v1beta/files?key=${getApiKey()}`,
      { method: 'POST', headers: { 'X-Goog-Upload-Protocol': 'multipart' }, body: formData }
    );
    if (!uploadResp.ok) {
      const e = await uploadResp.json().catch(() => ({}));
      throw new Error(e.error?.message || `업로드 실패 (${uploadResp.status})`);
    }
    const { file: uploaded } = await uploadResp.json();

    updateLoadingText('파일 처리 대기 중...');
    await waitForFileActive(uploaded.name);

    updateLoadingText('AI 분석 중...');
    const prompt = mode === 'full'
      ? `새로 업로드된 문서의 내용으로 현재 문서를 완전히 교체해 주세요.\n결과를 마크다운 형식으로 작성해 주세요.\n\n[현재 문서]\n${currentText}`
      : `다음 지시에 따라 현재 문서의 특정 부분을 업데이트해 주세요:\n"${instruction}"\n나머지 부분은 그대로 유지하고, 업데이트된 전체 문서를 마크다운 형식으로 작성해 주세요.\n\n[현재 문서]\n${currentText}`;

    const analysisResp = await fetch(
      `${API_BASE}/v1beta/models/${MODEL_ADVANCED}:generateContent?key=${getApiKey()}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [
            { fileData: { mimeType: uploaded.mimeType, fileUri: uploaded.uri } },
            { text: prompt }
          ]}],
          generationConfig: { temperature: 0.3, maxOutputTokens: 8192 }
        })
      }
    );
    if (!analysisResp.ok) {
      const e = await analysisResp.json().catch(() => ({}));
      throw new Error(e.error?.message || `분석 실패 (${analysisResp.status})`);
    }
    const data = await analysisResp.json();
    const content = data.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!content) throw new Error('AI 응답을 받지 못했습니다.');

    hideLoadingOverlay();
    const title = mode === 'full' ? '🔄 전체 교체 미리보기' : '✏️ 부분 업데이트 미리보기';
    showPreview(content, title, 'replace-body');

  } catch (err) {
    hideLoadingOverlay();
    showToast(`오류: ${err.message}`, 'error');
  }
}

async function getCurrentDocumentText() {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load('text');
    await context.sync();
    return body.text;
  });
}

async function replaceDocumentBody(markdownText) {
  const paragraphs = parseMarkdownToWordParagraphs(markdownText);
  if (paragraphs.length === 0) return;

  await Word.run(async (context) => {
    context.document.body.clear();
    await context.sync();

    // [수정] 'List Bullet' / 'List Number' 제거 → Word Online InvalidArgument 방지
    const styleMap = {
      h1: 'Heading 1', h2: 'Heading 2', h3: 'Heading 3', h4: 'Heading 4',
      bullet: 'Normal', number: 'Normal',
      normal: 'Normal', code: 'Normal', blockquote: 'Normal', empty: 'Normal',
    };

    let ref = context.document.body;
    let isFirst = true;
    let listCounter = 0;
    let prevType = null;

    for (const para of paragraphs) {
      // 번호 목록 카운터 관리
      if (para.type === 'number') {
        if (prevType !== 'number') listCounter = 0;
        listCounter++;
      } else {
        listCounter = 0;
      }
      prevType = para.type;

      const loc = isFirst ? Word.InsertLocation.start : Word.InsertLocation.after;
      const newPara = ref.insertParagraph('', loc);
      newPara.style = styleMap[para.type] || 'Normal';

      if (para.type === 'hr') {
        const r = newPara.insertText('─'.repeat(40), Word.InsertLocation.end);
        try { r.font.color = '#AAAAAA'; } catch (_) {}

      } else if (para.type === 'bullet') {
        // [수정] 수동 불릿 프리픽스
        newPara.insertText('• ', Word.InsertLocation.end);
        for (const run of para.runs) {
          if (!run.text) continue;
          const r = newPara.insertText(run.text, Word.InsertLocation.end);
          try {
            if (run.bold) r.font.bold = true;
            if (run.italic) r.font.italic = true;
            if (run.code) { r.font.name = 'Courier New'; r.font.size = 10; }
          } catch (_) {}
        }

      } else if (para.type === 'number') {
        // [수정] 수동 번호 프리픽스
        newPara.insertText(`${listCounter}. `, Word.InsertLocation.end);
        for (const run of para.runs) {
          if (!run.text) continue;
          const r = newPara.insertText(run.text, Word.InsertLocation.end);
          try {
            if (run.bold) r.font.bold = true;
            if (run.italic) r.font.italic = true;
            if (run.code) { r.font.name = 'Courier New'; r.font.size = 10; }
          } catch (_) {}
        }

      } else if (para.type === 'code') {
        try { newPara.font.name = 'Courier New'; newPara.font.size = 10; } catch (_) {}
        for (const run of para.runs) {
          if (run.text) newPara.insertText(run.text, Word.InsertLocation.end);
        }

      } else if (para.type === 'blockquote') {
        try { newPara.leftIndent = 36; } catch (_) {}
        for (const run of para.runs) {
          if (!run.text) continue;
          const r = newPara.insertText(run.text, Word.InsertLocation.end);
          try { r.font.italic = true; if (run.bold) r.font.bold = true; } catch (_) {}
        }

      } else if (para.type !== 'empty') {
        for (const run of para.runs) {
          if (!run.text) continue;
          const r = newPara.insertText(run.text, Word.InsertLocation.end);
          try {
            if (run.bold) r.font.bold = true;
            if (run.italic) r.font.italic = true;
            if (run.code) { r.font.name = 'Courier New'; r.font.size = 10; }
          } catch (_) {}
        }
      }

      ref = newPara;
      isFirst = false;
    }

    await context.sync();
  });
}

/* ========== Gemini API ========== */
async function callGeminiAPI(messages, model = MODEL_CHAT, withSearch = false) {
  const key = getApiKey();
  if (!key) throw new Error('API 키가 설정되지 않았습니다.');

  const body = {
    contents: messages,
    generationConfig: {
      temperature: 0.7,
      topP: 0.95,
      maxOutputTokens: 8192
    }
  };

  if (withSearch) {
    body.tools = [{ google_search: {} }];
  }

  const resp = await fetch(
    `${API_BASE}/v1beta/models/${model}:generateContent?key=${key}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    }
  );

  if (!resp.ok) {
    const errData = await resp.json().catch(() => ({}));
    const msg = errData.error?.message || `API 오류 (${resp.status})`;
    throw new Error(msg);
  }

  const data = await resp.json();

  // Handle grounding metadata (search results)
  const candidate = data.candidates?.[0];
  if (!candidate?.content?.parts?.[0]?.text) {
    // Try to get text from any part
    const allText = candidate?.content?.parts
      ?.filter(p => p.text)
      ?.map(p => p.text)
      ?.join('');
    if (allText) return allText;
    throw new Error('AI 응답이 비어있습니다.');
  }

  return candidate.content.parts[0].text;
}

/* ========== Preview Panel ========== */
function showPreview(text, title = '미리보기', mode = 'insert') {
  currentPreviewText = text;
  currentPreviewMode = mode;
  document.getElementById('preview-title').textContent = title;
  document.getElementById('preview-content').innerHTML = renderMarkdown(text);
  const insertBtn = document.getElementById('insert-preview-btn');
  insertBtn.textContent = mode === 'replace-body' ? '🔄 문서 전체 교체' : '📄 Word에 삽입';
  document.getElementById('preview-panel').classList.remove('hidden');
}

function closePreview() {
  document.getElementById('preview-panel').classList.add('hidden');
  currentPreviewText = '';
  currentPreviewMode = 'insert';
}

async function insertPreviewToDocument() {
  if (!currentPreviewText) return;
  try {
    if (currentPreviewMode === 'replace-body') {
      await replaceDocumentBody(currentPreviewText);
    } else {
      await insertMarkdownAsWordFormatting(currentPreviewText, false, useTrackChanges());
    }
    showToast('Word에 삽입되었습니다.', 'success');
    closePreview();
  } catch (err) {
    showPreviewInsertError(err);
  }
}

/* ========== Loading Overlay ========== */
function showLoadingOverlay(text = '처리 중...') {
  document.getElementById('loading-text').textContent = text;
  document.getElementById('loading-overlay').classList.remove('hidden');
}

function updateLoadingText(text) {
  document.getElementById('loading-text').textContent = text;
}

function hideLoadingOverlay() {
  document.getElementById('loading-overlay').classList.add('hidden');
}

/* ========== Typing Indicator ========== */
function showTypingIndicator() {
  document.getElementById('typing-indicator').classList.remove('hidden');
}
function hideTypingIndicator() {
  document.getElementById('typing-indicator').classList.add('hidden');
}

/* ========== Toast ========== */
let toastTimer = null;
function showToast(msg, type = 'info', durationMs = 3000) {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.className = type;
  toast.classList.remove('hidden');
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.add('hidden'), durationMs);
}

/* ========== Send Button State ========== */
function setSendButtonState(enabled) {
  const btn = document.getElementById('send-btn');
  btn.disabled = !enabled;
}

/* ========== Markdown Renderer ========== */
function renderMarkdown(text) {
  if (!text) return '';

  let html = escapeHtml(text);

  // Code blocks (must come before inline code)
  html = html.replace(/```[\w]*\n?([\s\S]*?)```/g, (_, code) =>
    `<pre><code>${code.trim()}</code></pre>`
  );

  // Inline code
  html = html.replace(/`([^`]+)`/g, '<code>$1</code>');

  // Headings
  html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
  html = html.replace(/^## (.+)$/gm, '<h2>$1</h2>');
  html = html.replace(/^# (.+)$/gm, '<h1>$1</h1>');

  // Horizontal rule
  html = html.replace(/^---+$/gm, '<hr>');

  // Bold and italic
  html = html.replace(/\*\*\*(.+?)\*\*\*/g, '<strong><em>$1</em></strong>');
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/\*(.+?)\*/g, '<em>$1</em>');

  // Blockquote
  html = html.replace(/^&gt; (.+)$/gm, '<blockquote>$1</blockquote>');

  // Task list items
  html = html.replace(/^- \[ \] (.+)$/gm,
    '<div class="task-item"><input type="checkbox" disabled> <span>$1</span></div>');
  html = html.replace(/^- \[x\] (.+)$/gim,
    '<div class="task-item"><input type="checkbox" checked disabled> <span>$1</span></div>');

  // Unordered list items
  html = html.replace(/^[-*] (.+)$/gm, '<li>$1</li>');
  html = html.replace(/(<li>.*<\/li>(\n|$))+/g, m => `<ul>${m}</ul>`);

  // Ordered list items
  html = html.replace(/^\d+\. (.+)$/gm, '<li>$1</li>');

  // Links
  html = html.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" target="_blank">$1</a>');

  // Paragraphs: wrap consecutive non-tag lines
  const lines = html.split('\n');
  const result = [];
  let inParagraph = false;

  for (let line of lines) {
    const isBlock = /^<(h[1-6]|ul|ol|li|pre|blockquote|hr|div)/.test(line) || line === '';
    if (!isBlock) {
      if (!inParagraph) { result.push('<p>'); inParagraph = true; }
      result.push(line + ' ');
    } else {
      if (inParagraph) { result.push('</p>'); inParagraph = false; }
      result.push(line);
    }
  }
  if (inParagraph) result.push('</p>');

  return result.join('\n');
}

function escapeHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/* ========== Utilities ========== */
function scrollChatToBottom() {
  const chatArea = document.getElementById('chat-area');
  requestAnimationFrame(() => {
    chatArea.scrollTop = chatArea.scrollHeight;
  });
}

function autoResizeTextarea(el) {
  el.style.height = 'auto';
  el.style.height = Math.min(el.scrollHeight, 120) + 'px';
}

document.addEventListener('DOMContentLoaded', () => {
  const input = document.getElementById('message-input');
  if (input) input.addEventListener('input', () => autoResizeTextarea(input));
});

function formatTime(date) {
  return date.toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' });
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function copyToClipboard(text) {
  navigator.clipboard.writeText(text)
    .then(() => showToast('클립보드에 복사되었습니다.', 'success'))
    .catch(() => {
      const ta = document.createElement('textarea');
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
      showToast('클립보드에 복사되었습니다.', 'success');
    });
}
