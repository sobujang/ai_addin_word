/* ========== Constants & Config ========== */
const API_BASE = 'https://generativelanguage.googleapis.com';
const MODEL_CHAT = 'gemini-2.5-flash';
const STORAGE_KEY_API = 'gemini_api_key_v2';
const STORAGE_KEY_MODEL = 'gemini_model_v2';
const STORAGE_KEY_LANG = 'gemini_lang_v2';
const STORAGE_KEY_TRACK = 'gemini_track_changes_v2';

const TEMPLATES = {
  normal: "전문적인 비서로서 친절하게 답변해주세요.",
  report: "보고서 양식입니다. [개요-주요내용-기대효과-결론] 순서로 구성하고, 개조식(#, ##, -)을 적극적으로 사용하세요.",
  memo: "공문서 양식입니다. 정중하고 공식적인 문체를 사용하며, [1. 목적, 2. 근거, 3. 협조사항] 형식을 준수하세요.",
  plan: "기획서 양식입니다. 문제 정의, 해결 방안, 추진 일정, 예상 소요 예산을 포함하여 논리적으로 작성하세요.",
  meeting: "회의록 양식입니다. 일시, 장소, 참석자, 안건, 논의 내용, 결정사항, 향후 일정을 명확히 구분하여 작성하세요."
};

let currentTemplate = 'normal';
let conversationHistory = [];
let isProcessing = false;
let styleGuide = ""; // 스타일 미러링을 위한 가이드

/* ========== Office Initialization ========== */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    loadSettings();
    initEventListeners();
    addMessage('assistant', "세팅이 완료되었습니다. 무엇을 도와드릴까요?");
  }
});

/* ========== Event Listeners ========== */
function initEventListeners() {
  document.getElementById('send-btn').addEventListener('click', handleSendMessage);
  document.getElementById('summarize-book-btn').addEventListener('click', handleSummarizeBook);
  document.getElementById('upload-audio-btn').addEventListener('click', () => document.getElementById('audio-file-input').click());
  document.getElementById('upload-doc-btn').addEventListener('click', () => document.getElementById('doc-file-input').click());
  document.getElementById('fetch-selection-btn').addEventListener('click', handleFetchSelection);

  document.getElementById('audio-file-input').addEventListener('change', (e) => onFileSelected(e, 'audio'));
  document.getElementById('doc-file-input').addEventListener('change', (e) => onFileSelected(e, 'doc'));

  document.getElementById('settings-btn').addEventListener('click', () => toggleModal('settings-overlay', true));
  document.getElementById('cancel-settings-btn').addEventListener('click', () => toggleModal('settings-overlay', false));
  document.getElementById('save-settings-btn').addEventListener('click', saveSettings);
  document.getElementById('clear-chat-btn').addEventListener('click', clearChat);

  // Template Selector
  document.querySelectorAll('.template-chip').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.template-chip').forEach(b => b.classList.remove('active'));
      e.target.classList.add('active');
      currentTemplate = e.target.dataset.template;
      addMessage('assistant', `✏️ 워크플로우가 [${e.target.textContent}] 양식으로 변경되었습니다.`);
    });
  });

  // Quick Actions (New)
  document.querySelectorAll('.quick-btn').forEach(btn => {
    btn.addEventListener('click', (e) => handleQuickAction(e.target.dataset.action));
  });

  document.getElementById('user-input').addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
    autoResizeTextarea(e.target);
  });
}

/* ========== API Handling ========== */
async function callGeminiAPI(contents, specializedModel = null) {
  const apiKey = localStorage.getItem(STORAGE_KEY_API);
  if (!apiKey) throw new Error("API 키가 없습니다. 설정에서 입력해주세요.");

  let modelName = specializedModel || localStorage.getItem(STORAGE_KEY_MODEL) || 'gemini-2.5-flash';
  const displayLang = localStorage.getItem(STORAGE_KEY_LANG) === 'en' ? 'English' : 'Korean';

  const templateGuide = TEMPLATES[currentTemplate] || TEMPLATES.normal;
  const systemInstruction = {
    role: "user",
    parts: [{ text: `Answer in ${displayLang}. ${templateGuide} ${styleGuide ? '문체 가이드: ' + styleGuide : ''} Maintain a professional tone.` }]
  };

  const body = {
    contents: [systemInstruction, ...contents],
    generationConfig: { temperature: 0.7, maxOutputTokens: 8192 }
  };

  try {
    return await executeFetch(modelName, apiKey, body);
  } catch (err) {
    return await executeFetch('gemini-2.5-flash', apiKey, body);
  }
}

async function executeFetch(model, key, body) {
  const response = await fetch(`${API_BASE}/v1beta/models/${model}:generateContent?key=${key}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  if (!response.ok) {
    const errorBody = await response.json().catch(() => ({}));
    throw new Error(errorBody.error?.message || `API 실패 (${response.status})`);
  }
  const data = await response.json();
  return data.candidates?.[0]?.content?.parts?.[0]?.text;
}

/* ========== Chat Core ========== */
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
    const responseText = await callGeminiAPI(conversationHistory);
    addMessage('assistant', responseText);
    conversationHistory.push({ role: 'model', parts: [{ text: responseText }] });
  } catch (err) {
    addMessage('assistant', `⚠️ ${err.message}`);
  } finally {
    showThinking(false);
  }
}

/* ========== Quick Actions & Style Mirroring ========== */
async function handleQuickAction(action) {
  if (isProcessing) return;

  // 선택 영역 텍스트 가져오기
  try {
    const text = await getWordSelection();
    if (!text) {
      showToast("먼저 워드에서 텍스트를 선택해 주세요.");
      return;
    }

    let prompt = "";
    switch (action) {
      case 'polish': prompt = `아래 텍스트의 맞춤법을 교정하고 비즈니스 업무에 적합한 세련된 문체로 다듬어줘: \n\n${text}`; break;
      case 'shorten': prompt = `아래 내용을 핵심 위주로 세 줄 이내로 요약해줘: \n\n${text}`; break;
      case 'formal': prompt = `아래 내용을 격식 있는 공문서체로 변환해줘: \n\n${text}`; break;
      case 'scan':
        styleGuide = text.substring(0, 500);
        addMessage('assistant', "✨ 현재 선택된 영역의 문체를 AI가 학습했습니다. 앞으로의 답변은 이 스타일을 따릅니다.");
        return;
    }

    showThinking(true);
    const response = await callGeminiAPI([{ role: 'user', parts: [{ text: prompt }] }]);
    addMessage('assistant', response);
  } catch (err) {
    showToast("텍스트 처리에 실패했습니다.");
  } finally {
    showThinking(false);
  }
}

async function getWordSelection() {
  return new Promise((resolve) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : null);
    });
  });
}

/* ========== File Processing ========== */
async function onFileSelected(e, type) {
  const file = e.target.files[0];
  if (!file) return;
  e.target.value = '';

  const apiKey = localStorage.getItem(STORAGE_KEY_API);
  if (!apiKey) { toggleModal('settings-overlay', true); return; }

  setProcessing(true, `${file.name} 업로드 중...`); // 파일은 여전히 오버레이 사용 (전체 주도권 필요)
  try {
    const supportedTypes = ['application/pdf', 'text/plain', 'audio/'];
    if (!supportedTypes.some(t => file.type.startsWith(t)) && !file.name.endsWith('.pdf')) {
      throw new Error("PDF 또는 텍스트 파일만 지원됩니다.");
    }

    const formData = new FormData();
    formData.append('metadata', new Blob([JSON.stringify({ file: { displayName: file.name } })], { type: 'application/json' }));
    formData.append('file', file);

    const uploadResp = await fetch(`${API_BASE}/upload/v1beta/files?key=${apiKey}`, {
      method: "POST",
      headers: { "X-Goog-Upload-Protocol": "multipart" },
      body: formData
    });
    const { file: uploadedFile } = await uploadResp.json();
    await waitForFileActive(uploadedFile.name, apiKey);

    const prompt = type === 'audio' ? "세부 회의록을 작성해줘." : "핵심 요약본을 만들어줘.";
    const response = await callGeminiAPI([{
      role: 'user',
      parts: [
        { fileData: { mimeType: uploadedFile.mimeType, fileUri: uploadedFile.uri } },
        { text: prompt }
      ]
    }]);

    addMessage('assistant', `✅ ${file.name} 분석 완료.`);
    await insertMarkdownToWord(response);
  } catch (err) {
    addMessage('assistant', `⚠️ ${err.message}`);
  } finally {
    setProcessing(false);
  }
}

/* ========== Word Interaction ========== */
async function insertMarkdownToWord(markdown) {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    const lines = markdown.split('\n');
    let anchorPara = selection.paragraphs.getFirst();
    let isFirst = true;

    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed && !isFirst) continue;
      let targetPara = isFirst ? anchorPara : anchorPara.insertParagraph("", "After");
      targetPara.clear();

      let cleanLine = line;
      if (trimmed.startsWith('# ')) { targetPara.styleBuiltIn = "Heading1"; cleanLine = trimmed.replace('# ', ''); }
      else if (trimmed.startsWith('## ')) { targetPara.styleBuiltIn = "Heading2"; cleanLine = trimmed.replace('## ', ''); }
      else { targetPara.styleBuiltIn = "Normal"; }

      parseAndInsertInline(targetPara, cleanLine);
      anchorPara = targetPara;
      isFirst = false;
    }
    await context.sync();
  });
}

function parseAndInsertInline(paragraph, text) {
  const regex = /(\*\*|__)(.*?)\1|(\*|_)(.*?)\3|(`)(.*?)\5|([^*`_]+|[*`_])/g;
  let match;
  while ((match = regex.exec(text)) !== null) {
    let content, isBold = false, isItalic = false;
    if (match[1]) { content = match[2]; isBold = true; }
    else if (match[3]) { content = match[4]; isItalic = true; }
    else { content = match[7]; }

    if (content) {
      const range = paragraph.insertText(content, "End");
      if (isBold) range.font.bold = true;
      if (isItalic) range.font.italic = true;
    }
  }
}

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
    const insBtn = document.createElement('button'); insBtn.className = 'action-mini-btn'; insBtn.innerHTML = '📄 삽입';
    insBtn.onclick = () => insertMarkdownToWord(text);
    const cpBtn = document.createElement('button'); cpBtn.className = 'action-mini-btn'; cpBtn.innerHTML = '📋 복사';
    cpBtn.onclick = () => { navigator.clipboard.writeText(text); showToast("복사되었습니다."); };
    actions.append(insBtn, cpBtn);
    wrapper.appendChild(actions);
  }
  container.appendChild(wrapper);
  container.scrollTop = container.scrollHeight;
}

function showThinking(show) {
  isProcessing = show;
  const container = document.getElementById('chat-container');
  const existing = document.getElementById('ai-thinking');
  if (show) {
    if (existing) return;
    const thinking = document.createElement('div');
    thinking.id = 'ai-thinking';
    thinking.className = 'thinking-bubble';
    thinking.innerHTML = `<div class="typing-dots"><span></span><span></span><span></span></div><div class="thinking-text">Gemini가 생각 중...</div>`;
    container.appendChild(thinking);
    container.scrollTop = container.scrollHeight;
  } else {
    if (existing) existing.remove();
  }
}

function setProcessing(loading, text = "처리 중...") {
  document.getElementById('loading-overlay').style.display = loading ? 'flex' : 'none';
  document.getElementById('loading-text').textContent = text;
}

function showToast(msg) {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.className = 'show';
  setTimeout(() => { toast.className = ''; }, 3000);
}

function loadSettings() {
  document.getElementById('api-key-input').value = localStorage.getItem(STORAGE_KEY_API) || '';
  document.getElementById('model-select').value = localStorage.getItem(STORAGE_KEY_MODEL) || 'gemini-2.5-flash';
  document.getElementById('lang-select').value = localStorage.getItem(STORAGE_KEY_LANG) || 'ko';
  document.getElementById('track-changes-toggle').checked = localStorage.getItem(STORAGE_KEY_TRACK) === 'true';
}

function saveSettings() {
  localStorage.setItem(STORAGE_KEY_API, document.getElementById('api-key-input').value.trim());
  localStorage.setItem(STORAGE_KEY_MODEL, document.getElementById('model-select').value);
  localStorage.setItem(STORAGE_KEY_LANG, document.getElementById('lang-select').value);
  localStorage.setItem(STORAGE_KEY_TRACK, document.getElementById('track-changes-toggle').checked);
  toggleModal('settings-overlay', false);
  showToast("설정이 저장되었습니다.");
}

async function handleSummarizeBook() {
  if (conversationHistory.length === 0) return;
  showThinking(true);
  try {
    const context = conversationHistory.map(m => `${m.role === 'user' ? '사용자' : 'AI'}: ${m.parts[0].text}`).join('\n\n');
    const prompt = `대화 내용을 단행본 형식으로 정리해줘.\n\n${context}`;
    const result = await callGeminiAPI([{ role: 'user', parts: [{ text: prompt }] }]);
    await insertMarkdownToWord(result);
  } catch (err) { addMessage('assistant', `⚠️ ${err.message}`); }
  finally { showThinking(false); }
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
function autoResizeTextarea(el) { el.style.height = 'auto'; el.style.height = (el.scrollHeight) + 'px'; }
async function handleFetchSelection() {
  const text = await getWordSelection();
  if (text) {
    document.getElementById('user-input').value += `\n[선택한 내용]: ${text}\n`;
    autoResizeTextarea(document.getElementById('user-input'));
  }
}
function clearChat() { if (confirm("삭제하시겠습니까?")) { conversationHistory = []; document.getElementById('chat-container').innerHTML = ''; styleGuide = ""; } }
