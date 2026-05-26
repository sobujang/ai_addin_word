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

  // Input
  document.getElementById('user-input').addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
    autoResizeTextarea(e.target);
  });
}

/* ========== API Handling (With Fallback) ========== */
async function callGeminiAPI(contents, specializedModel = null) {
  const apiKey = localStorage.getItem(STORAGE_KEY_API);
  if (!apiKey) throw new Error("API 키가 없습니다. 설정에서 입력해주세요.");

  let modelName = specializedModel || localStorage.getItem(STORAGE_KEY_MODEL) || 'gemini-2.5-flash';
  const displayLang = localStorage.getItem(STORAGE_KEY_LANG) === 'en' ? 'English' : 'Korean';

  // System instruction for language and template
  const templateGuide = TEMPLATES[currentTemplate] || TEMPLATES.normal;
  const systemInstruction = {
    role: "user",
    parts: [{ text: `Answer in ${displayLang}. ${templateGuide} Maintain a professional and helpful tone.` }]
  };

  const body = {
    contents: [systemInstruction, ...contents],
    generationConfig: { temperature: 0.7, maxOutputTokens: 8192 }
  };

  try {
    return await executeFetch(modelName, apiKey, body);
  } catch (err) {
    console.warn(`${modelName} 호출 실패, 2.5-flash로 폴백 시도...`, err.message);
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

/* ========== File Upload & Processing ========== */
async function onFileSelected(e, type) {
  const file = e.target.files[0];
  if (!file) return;
  e.target.value = '';

  const apiKey = localStorage.getItem(STORAGE_KEY_API);
  if (!apiKey) { toggleModal('settings-overlay', true); return; }

  setProcessing(true, `${file.name} 분석 중...`);
  try {
    // 1. Upload
    const formData = new FormData();
    formData.append('metadata', new Blob([JSON.stringify({ file: { displayName: file.name } })], { type: 'application/json' }));
    formData.append('file', file);

    const uploadResp = await fetch(`${API_BASE}/upload/v1beta/files?key=${apiKey}`, {
      method: "POST",
      headers: { "X-Goog-Upload-Protocol": "multipart" },
      body: formData
    });
    const { file: uploadedFile } = await uploadResp.json();

    // 2. Poll
    await waitForFileActive(uploadedFile.name, apiKey);

    // 3. Analyze based on type
    let prompt = "";
    if (type === 'audio') {
      prompt = `
        다음 음성 파일은 중요한 회의 녹음 파일입니다. 
        전체 내용을 처음부터 끝까지 꼼꼼하게 듣고, 누락되는 내용 없이 상세한 회의록을 작성해 주세요.

        작성 가이드:
        1. 메타 데이터: 파악 가능한 일시와 참석자를 적어주세요.
        2. 상세 논의 내용: 회의에서 오간 대화들을 주요 주제별(Section)로 나누어 아주 상세하게 요약해 주세요. 각 주제 내에서 누가 어떤 의견을 냈는지도 포함하면 좋습니다.
        3. 결정 사항 및 액션 아이템: 확정된 결론과 향후 실행할 과제(담당자 포함)를 명확히 정리해 주세요.
        
        응답 형식:
        # 🎙️ 상세 회의록
        - 일시: 
        - 참석자: 

        ## 📝 주요 주제별 논의 내용
        (주제 1: ...)
        (주제 2: ...)

        ## ✅ 결정 사항 및 향후 계획
        - 결정된 내용 1...
        
        ## 🚀 액션 아이템
        - [ ] 담당자: 할 일 (기한)
      `;
    } else {
      prompt = "이 문서 내용을 꼼꼼히 읽고 전체 내용을 파악해 주세요. 그리고 현재 작성 중인 워드 문서의 문체와 일치하도록 내용을 논리적으로 다듬어서 삽입 가능한 형태로 정리해 주세요.";
    }

    const response = await callGeminiAPI([{
      role: 'user',
      parts: [
        { fileData: { mimeType: uploadedFile.mimeType, fileUri: uploadedFile.uri } },
        { text: prompt }
      ]
    }]);

    addMessage('assistant', `${type === 'audio' ? '🎙️ 회의록' : '📄 문서 분석'}이 완료되었습니다.`);
    await insertMarkdownToWord(response);
  } catch (err) {
    addMessage('assistant', `⚠️ 처리 오류: ${err.message}`);
  } finally {
    setProcessing(false);
  }
}

/* ========== Features ========== */
async function handleSendMessage() {
  if (isProcessing) return;
  const inputEl = document.getElementById('user-input');
  const text = inputEl.value.trim();
  if (!text) return;

  inputEl.value = '';
  autoResizeTextarea(inputEl);
  addMessage('user', text);
  conversationHistory.push({ role: 'user', parts: [{ text }] });

  setProcessing(true);
  try {
    const responseText = await callGeminiAPI(conversationHistory);
    addMessage('assistant', responseText);
    conversationHistory.push({ role: 'model', parts: [{ text: responseText }] });
  } catch (err) {
    addMessage('assistant', `⚠️ ${err.message}`);
  } finally {
    setProcessing(false);
  }
}

async function handleSummarizeBook() {
  if (conversationHistory.length === 0) return;
  setProcessing(true, "단행본 형식 정리 중...");
  try {
    const context = conversationHistory.map(m => `${m.role === 'user' ? '사용자' : 'AI'}: ${m.parts[0].text}`).join('\n\n');
    const prompt = `다음 대화를 단행본 형식으로 정리해주세요. 챕터와 소제목을 나누고 문어체로 작성하세요.\n\n${context}`;
    const result = await callGeminiAPI([{ role: 'user', parts: [{ text: prompt }] }]);
    await insertMarkdownToWord(result);
  } catch (err) {
    addMessage('assistant', `⚠️ 오류: ${err.message}`);
  } finally {
    setProcessing(false);
  }
}

/* ========== Word & Utilities ========== */
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
      targetPara.clear(); // 기존 내용 삭제

      // 1. 단락 스타일(제목 등) 결정
      let cleanLine = line;
      if (trimmed.startsWith('# ')) {
        targetPara.styleBuiltIn = "Heading1";
        cleanLine = trimmed.replace('# ', '');
      } else if (trimmed.startsWith('## ')) {
        targetPara.styleBuiltIn = "Heading2";
        cleanLine = trimmed.replace('## ', '');
      } else if (trimmed.startsWith('### ')) {
        targetPara.styleBuiltIn = "Heading3";
        cleanLine = trimmed.replace('### ', '');
      } else {
        targetPara.styleBuiltIn = "Normal";
      }

      // 2. 인라인 서식(Bold, Italic) 파싱 및 삽입
      parseAndInsertInline(targetPara, cleanLine);

      anchorPara = targetPara;
      isFirst = false;
    }
    await context.sync();
  });
}

/**
 * 인라인 마크다운(**, *)을 파싱하여 워드 단락에 삽입합니다.
 */
function parseAndInsertInline(paragraph, text) {
  // 간단한 정규식 파서: **bold**, *italic* 등을 찾음
  const regex = /(\*\*|__)(.*?)\1|(\*|_)(.*?)\3|(`)(.*?)\5|([^*`_]+|[*`_])/g;
  let match;

  while ((match = regex.exec(text)) !== null) {
    let content, isBold = false, isItalic = false, isCode = false;

    if (match[1]) { // Bold (**text**)
      content = match[2];
      isBold = true;
    } else if (match[3]) { // Italic (*text*)
      content = match[4];
      isItalic = true;
    } else if (match[5]) { // Code (`text`)
      content = match[6];
      isCode = true;
    } else { // Plain text
      content = match[7];
    }

    if (content) {
      const range = paragraph.insertText(content, "End");
      if (isBold) range.font.bold = true;
      if (isItalic) range.font.italic = true;
      if (isCode) {
        range.font.name = "Courier New";
        range.font.size = 10;
      }
    }
  }
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
}

function addMessage(role, text) {
  const container = document.getElementById('chat-container');
  const wrapper = document.createElement('div');
  wrapper.className = `message ${role}`;

  const bubble = document.createElement('div');
  bubble.className = 'bubble';
  bubble.innerHTML = text.replace(/\n/g, '<br>');
  wrapper.appendChild(bubble);

  // AI 응답일 경우 액션 버튼 추가
  if (role === 'assistant') {
    const actions = document.createElement('div');
    actions.className = 'msg-actions';

    // 삽입 버튼
    const insertBtn = document.createElement('button');
    insertBtn.className = 'action-mini-btn';
    insertBtn.innerHTML = '📄 삽입';
    insertBtn.onclick = () => insertMarkdownToWord(text);

    // 복사 버튼
    const copyBtn = document.createElement('button');
    copyBtn.className = 'action-mini-btn';
    copyBtn.innerHTML = '📋 복사';
    copyBtn.onclick = () => {
      navigator.clipboard.writeText(text);
      showToast("클립보드에 복사되었습니다.");
    };

    actions.append(insertBtn, copyBtn);
    wrapper.appendChild(actions);
  }

  container.appendChild(wrapper);
  container.scrollTop = container.scrollHeight;
}

function showToast(msg) {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.className = 'show';
  setTimeout(() => {
    toast.className = '';
  }, 3000);
}

function setProcessing(loading, text = "처리 중...") {
  isProcessing = loading;
  document.getElementById('loading-overlay').style.display = loading ? 'flex' : 'none';
  document.getElementById('loading-text').textContent = text;
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
  Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
      document.getElementById('user-input').value += `\n[선택한 내용]: ${result.value}\n`;
      autoResizeTextarea(document.getElementById('user-input'));
    }
  });
}
function clearChat() { if (confirm("삭제하시겠습니까?")) { conversationHistory = []; document.getElementById('chat-container').innerHTML = ''; } }
