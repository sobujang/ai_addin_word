/* ========== Config & State ========== */
const API_BASE = 'https://generativelanguage.googleapis.com';
const MODEL_DEFAULT = 'gemini-2.5-flash';
const STORAGE_KEYS = { API: 'gemini_api_key_v2', MODEL: 'gemini_model_v2', LANG: 'gemini_lang_v2', TRACK: 'gemini_track_changes_v2' };

const TEMPLATES = {
  normal: "전문 비서로서 친절하게 답변해주세요.",
  report: "보고서 양식으로 각 항목을 섹션별로 요약해 주세요.",
  memo: "공문서 성격의 정중하고 공식적인 문체로 작성해 주세요.",
  plan: "기획서 양식입니다. 문제정의와 기대효과를 포함해 주세요.",
  meeting: "시간, 장소, 참석자, 주요 결정사항 위주의 회의록 형식입니다."
};

let currentTemplate = 'normal';
let conversationHistory = [];
let isProcessing = false;
let styleGuide = "";

/* ========== Initialization ========== */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    loadSettings();
    initEventListeners();
  }
});

function initEventListeners() {
  document.getElementById('send-btn').addEventListener('click', handleSendMessage);
  document.getElementById('user-input').addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSendMessage(); }
    autoResizeTextarea(e.target);
  });

  document.getElementById('attach-btn').addEventListener('click', (e) => {
    e.stopPropagation();
    document.getElementById('attach-menu').classList.toggle('show');
  });
  document.addEventListener('click', () => document.getElementById('attach-menu').classList.remove('show'));

  document.getElementById('style-select').addEventListener('change', (e) => {
    currentTemplate = e.target.value;
    addMessage('assistant', `✏️ [${e.target.options[e.target.selectedIndex].text}] 모드로 전환되었습니다.`);
  });

  document.getElementById('summarize-book-btn').addEventListener('click', handleSummarizeBook);
  document.getElementById('upload-audio-btn').addEventListener('click', () => document.getElementById('audio-file-input').click());
  document.getElementById('upload-doc-btn').addEventListener('click', () => document.getElementById('doc-file-input').click());
  document.getElementById('fetch-selection-btn').addEventListener('click', handleFetchSelection);

  document.getElementById('audio-file-input').addEventListener('change', (e) => onFileSelected(e, 'audio'));
  document.getElementById('doc-file-input').addEventListener('change', (e) => onFileSelected(e, 'doc'));

  document.querySelectorAll('.quick-btn').forEach(btn => {
    btn.addEventListener('click', () => handleQuickAction(btn.dataset.action));
  });

  document.getElementById('settings-btn').addEventListener('click', () => toggleModal('settings-overlay', true));
  document.getElementById('cancel-settings-btn').addEventListener('click', () => toggleModal('settings-overlay', false));
  document.getElementById('save-settings-btn').addEventListener('click', saveSettings);
  document.getElementById('clear-chat-btn').addEventListener('click', clearChat);
}

/* ========== API & Chat Core ========== */
async function callGeminiAPI(contents, specializedModel = null) {
  const apiKey = localStorage.getItem(STORAGE_KEYS.API);
  if (!apiKey) throw new Error("API 키 설정 필요");

  const model = specializedModel || localStorage.getItem(STORAGE_KEYS.MODEL) || MODEL_DEFAULT;
  const lang = localStorage.getItem(STORAGE_KEYS.LANG) === 'en' ? 'English' : 'Korean';
  const templateGuide = TEMPLATES[currentTemplate] || TEMPLATES.normal;

  const systemInstruction = {
    role: "user",
    parts: [{ text: `Answer in ${lang}. ${templateGuide} ${styleGuide ? '문체 가이드: ' + styleGuide : ''} Maintain a professional tone.` }]
  };

  const response = await fetch(`${API_BASE}/v1beta/models/${model}:generateContent?key=${apiKey}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ contents: [systemInstruction, ...contents], generationConfig: { temperature: 0.7, maxOutputTokens: 8192 } })
  });

  if (!response.ok) throw new Error("API 요청 실패");
  const data = await response.json();
  return data.candidates?.[0]?.content?.parts?.[0]?.text;
}

async function handleSendMessage() {
  if (isProcessing) return;
  const inputEl = document.getElementById('user-input');
  const text = inputEl.value.trim();
  if (!text) return;

  inputEl.value = '';
  autoResizeTextarea(inputEl);
  addMessage('user', text);
  conversationHistory.push({ role: 'user', parts: [{ text }] });

  showThinking(true);
  try {
    const response = await callGeminiAPI(conversationHistory);
    addMessage('assistant', response);
    conversationHistory.push({ role: 'model', parts: [{ text: response }] });
    if (document.getElementById('auto-apply-toggle').checked) await replaceMarkdownInWord(response);
  } catch (err) {
    addMessage('assistant', `⚠️ 오류: ${err.message}`);
  } finally {
    showThinking(false);
  }
}

/* ========== Actions ========== */
async function handleQuickAction(action) {
  if (isProcessing) return;
  const text = await getSelection();
  if (!text) return showToast("텍스트를 선택해 주세요.");

  let prompt = "";
  if (action === 'polish') prompt = `교조/비즈니스 문체로 교정해줘: \n${text}`;
  else if (action === 'shorten') prompt = `간결하게 요약해줘: \n${text}`;
  else if (action === 'formal') prompt = `격식 있는 말투로 바꿔줘: \n${text}`;
  else if (action === 'scan') { styleGuide = text.substring(0, 500); return showToast("스타일 학습 완료"); }

  showThinking(true);
  try {
    const res = await callGeminiAPI([{ role: 'user', parts: [{ text: prompt }] }]);
    addMessage('assistant', res);
    if (document.getElementById('auto-apply-toggle').checked) await replaceMarkdownInWord(res);
  } catch (err) {
    showToast("처리 중 오류 발생");
  } finally { showThinking(false); }
}

async function handleSummarizeBook() {
  if (conversationHistory.length === 0) return showToast("대화 내용이 없습니다.");
  showThinking(true);
  try {
    const context = conversationHistory.map(m => `${m.role === 'user' ? '사용자' : 'AI'}: ${m.parts[0].text}`).join('\n\n');
    const prompt = `이 대화 내용을 단행본 형식의 글로 정리해줘.\n\n${context}`;
    const res = await callGeminiAPI([{ role: 'user', parts: [{ text: prompt }] }]);
    addMessage('assistant', "📖 단행본 정리가 완료되었습니다.");
    await insertMarkdownToWord(res);
  } catch (err) { showToast("정리 실패"); }
  finally { showThinking(false); }
}

/* ========== File Processing ========== */
async function onFileSelected(e, type) {
  const file = e.target.files[0];
  if (!file) return;
  e.target.value = '';
  const apiKey = localStorage.getItem(STORAGE_KEYS.API);
  if (!apiKey) return showToast("API 키 설정을 확인하세요.");

  setProcessing(true, `${file.name} 분석 중...`);
  try {
    const formData = new FormData();
    formData.append('metadata', new Blob([JSON.stringify({ file: { displayName: file.name } })], { type: 'application/json' }));
    formData.append('file', file);

    const uploadResp = await fetch(`${API_BASE}/upload/v1beta/files?key=${apiKey}`, {
      method: "POST", headers: { "X-Goog-Upload-Protocol": "multipart" }, body: formData
    });
    const { file: uploadedFile } = await uploadResp.json();
    await waitForFileActive(uploadedFile.name, apiKey);

    const prompt = type === 'audio' ? "세부 회의록을 작성해줘." : "문서 내용을 요약해줘.";
    const response = await callGeminiAPI([{
      role: 'user',
      parts: [
        { fileData: { mimeType: uploadedFile.mimeType, fileUri: uploadedFile.uri } },
        { text: prompt }
      ]
    }]);

    addMessage('assistant', `✅ ${file.name} 처리 완료.`);
    await insertMarkdownToWord(response);
  } catch (err) { addMessage('assistant', `⚠️ 파일 처리 오류: ${err.message}`); }
  finally { setProcessing(false); }
}

/* ========== Word Interaction ========== */
async function getSelection() {
  return new Promise(r => Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, res => r(res.value)));
}

async function replaceMarkdownInWord(md) {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.clear();
    const lines = md.split('\n');
    let anchor = sel.paragraphs.getFirst();
    for (const [idx, line] of lines.entries()) {
      if (!line.trim() && idx > 0) continue;
      let para = idx === 0 ? anchor : anchor.insertParagraph("", "After");
      para.clear();
      if (line.startsWith('# ')) { para.styleBuiltIn = "Heading1"; para.insertText(line.substring(2), "Replace"); }
      else if (line.startsWith('## ')) { para.styleBuiltIn = "Heading2"; para.insertText(line.substring(3), "Replace"); }
      else { para.insertText(line, "Replace"); }
      anchor = para;
    }
    await ctx.sync();
    showToast("반영 완료");
  });
}

function insertMarkdownToWord(md) { replaceMarkdownInWord(md); }

/* ========== UI Helpers ========== */
function addMessage(role, text) {
  const container = document.getElementById('chat-container');
  const wrapper = document.createElement('div');
  wrapper.className = `message ${role}`;
  const bubble = document.createElement('div');
  bubble.className = 'bubble';
  bubble.innerHTML = text.replace(/\n/g, '<br>');
  wrapper.appendChild(bubble);

  if (role === 'assistant') {
    const actions = document.createElement('div');
    actions.className = 'msg-actions';
    actions.innerHTML = `<button class="action-mini-btn" onclick="insertMarkdownToWord(\`${text}\`)">📄 삽입</button>
                         <button class="action-mini-btn" style="color:#1a73e8" onclick="replaceMarkdownInWord(\`${text}\`)">🪄 반영</button>
                         <button class="action-mini-btn" onclick="navigator.clipboard.writeText(\`${text}\`); showToast('복사됨')">📋 복사</button>`;
    wrapper.appendChild(actions);
  }
  container.appendChild(wrapper);
  container.scrollTop = container.scrollHeight;
}

function showThinking(show) {
  isProcessing = show;
  const existing = document.getElementById('ai-thinking');
  if (show && !existing) {
    const div = document.createElement('div');
    div.id = 'ai-thinking'; div.className = 'thinking-bubble';
    div.innerHTML = `<div class="typing-dots"><span></span><span></span><span></span></div> Reasoning...`;
    document.getElementById('chat-container').appendChild(div);
    document.getElementById('chat-container').scrollTop = document.getElementById('chat-container').scrollHeight;
  } else if (!show && existing) existing.remove();
}

function showToast(msg) {
  const t = document.getElementById('toast');
  if (!t) return;
  t.textContent = msg; t.className = 'show';
  setTimeout(() => t.className = '', 2000);
}

function setProcessing(loading, text = "처리 중...") {
  const overlay = document.getElementById('loading-overlay');
  if (overlay) {
    overlay.style.display = loading ? 'flex' : 'none';
    document.getElementById('loading-text').textContent = text;
  }
}

async function waitForFileActive(fileName, apiKey) {
  while (true) {
    const resp = await fetch(`${API_BASE}/v1beta/${fileName}?key=${apiKey}`);
    const data = await resp.json();
    if (data.state === 'ACTIVE') return;
    await new Promise(r => setTimeout(r, 2000));
  }
}

function toggleModal(id, show) { document.getElementById(id).style.display = show ? 'flex' : 'none'; }
function autoResizeTextarea(el) { el.style.height = 'auto'; el.style.height = el.scrollHeight + 'px'; }
function clearChat() { if (confirm("대화 삭제?")) { conversationHistory = []; document.getElementById('chat-container').innerHTML = ''; styleGuide = ""; } }
async function handleFetchSelection() {
  const t = await getSelection();
  if (t) { document.getElementById('user-input').value += ` [선택: ${t}] `; autoResizeTextarea(document.getElementById('user-input')); }
}
function loadSettings() {
  document.getElementById('api-key-input').value = localStorage.getItem(STORAGE_KEYS.API) || '';
}
function saveSettings() {
  localStorage.setItem(STORAGE_KEYS.API, document.getElementById('api-key-input').value.trim());
  localStorage.setItem(STORAGE_KEYS.MODEL, document.getElementById('model-select').value);
  localStorage.setItem(STORAGE_KEYS.LANG, document.getElementById('lang-select').value);
  toggleModal('settings-overlay', false);
  showToast("저장됨");
}
