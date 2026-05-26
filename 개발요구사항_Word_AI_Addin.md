# Word AI Add-in 개발 요구사항 문서

## 1. 프로젝트 개요

**프로젝트명:** Word AI Add-in (Gemini 연동)  
**베이스 오픈소스:** [AnsonLai/Gemini-AI-for-Office](https://github.com/AnsonLai/Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting)  
**목적:** Microsoft Word 우측 패널에서 Gemini AI와 대화하고, 대화 내용을 문서로 정리하거나 음성 파일을 회의록으로 변환하여 Word에 자동 삽입, 다른 문서파일을 업로드해서 원본파일 업데이트 기능 추가  
**사용자:** 사무국 내부 4명  
**플랫폼:** Windows, Microsoft Word (Office Add-in)

---

## 2. 기술 스택

| 항목 | 내용 |
|------|------|
| 개발 방식 | Microsoft Office Add-in (Task Pane) |
| 언어 | HTML, CSS, JavaScript |
| Office 연동 | Office.js API |
| AI 모델 | Google Gemini API (gemini-2.5-flash) |
| 인증 방식 | API 키 방식 (각 사용자가 개별 키 입력) |
| 배포 방식 | 폴더 공유 (내부 sideload 방식) |

---

## 3. AnsonLai 베이스에서 추가/수정할 기능

AnsonLai 원본 코드를 그대로 유지하되, 아래 기능을 추가한다.

### 3-1. 기존 유지 기능 (AnsonLai 원본)
- Gemini API 키 입력 필드 (⚙️ 설정 아이콘)
- 우측 패널 채팅 UI
- 채팅 내용 기반 문서 편집 및 Word 삽입
- Track Changes(수정 추적) 지원
- Google 검색 연동

### 3-2. 추가할 기능

#### ① 대화 내용 → 단행본 형식 정리 버튼
- 채팅 패널 하단에 **"단행본 형식으로 정리"** 버튼 추가
- 버튼 클릭 시 지금까지의 전체 대화 내용을 Gemini에 전달
- Gemini에 전달할 프롬프트:
  ```
  아래는 특정 주제에 대한 대화 내용입니다.
  이 대화를 단행본 형식의 글로 정리해주세요.
  챕터와 소제목 구조로 구성하고, 자연스러운 문어체로 작성해주세요.
  
  [대화 내용]
  {conversation_history}
  ```
- 생성된 글을 Word 문서의 현재 커서 위치에 삽입
- 삽입 전 패널에서 미리보기 제공

#### ② 음성 파일 업로드 → 회의록 자동 작성
- 채팅 패널 내 **"음성 파일 업로드"** 버튼 추가
- 지원 포맷: MP3, WAV, M4A, FLAC
- 업로드 흐름:
  1. 사용자가 파일 선택
  2. Gemini Files API로 파일 업로드
  3. 아래 프롬프트로 분석 요청:
     ```
     다음 음성 파일은 회의 녹음입니다.
     아래 형식으로 회의록을 작성해주세요:
     
     # 회의록
     - 일시: (파악 가능한 경우)
     - 참석자: (파악 가능한 경우)
     
     ## 주요 논의사항
     (내용 정리)
     
     ## 결정사항
     (결정된 내용)
     
     ## 액션아이템
     - [ ] 담당자: 내용 (기한이 언급된 경우 포함)
     ```
  4. 생성된 회의록을 Word 문서에 삽입
- 처리 중 로딩 표시 필요 (파일 크기에 따라 시간 소요)

---

## 4. Gemini API 사용 방식

### 모델 선택 기준
- **일반 채팅 및 문서 편집:** `gemini-2.0-flash-lite` (무료 1,000 RPD)
- **단행본 정리 및 회의록 작성:** `gemini-2.0-flash` (무료 200 RPD, 품질 우선)

### 음성 파일 업로드 API 흐름
```javascript
// Step 1: 파일 업로드 (Files API)
const uploadResponse = await fetch(
  `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${API_KEY}`,
  {
    method: "POST",
    headers: { "X-Goog-Upload-Protocol": "multipart" },
    body: formData  // 음성 파일 포함
  }
);
const { file } = await uploadResponse.json();

// Step 2: 분석 요청
const response = await fetch(
  `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${API_KEY}`,
  {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      contents: [{
        parts: [
          { fileData: { mimeType: file.mimeType, fileUri: file.uri } },
          { text: "위 음성 파일을 회의록 형식으로 정리해주세요." }
        ]
      }]
    })
  }
);
```

### Word 문서 삽입 방법
```javascript
// 현재 커서 위치에 텍스트 삽입
Office.context.document.setSelectedDataAsync(
  generatedText,
  { coercionType: Office.CoercionType.Text },
  (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error("삽입 실패:", result.error.message);
    }
  }
);
```

---

## 5. UI 구성 (태스크 패널)

```
┌─────────────────────────────┐
│  ⚙️ Gemini AI Assistant      │
├─────────────────────────────┤
│                             │
│  [채팅 대화 영역]            │
│  User: ...                  │
│  AI: ...                    │
│                             │
├─────────────────────────────┤
│  [텍스트 입력창]             │
│  [전송 버튼]                 │
├─────────────────────────────┤
│  [단행본 형식으로 정리] 버튼  │
│  [🎙️ 음성 파일 업로드] 버튼  │
└─────────────────────────────┘
```

---

## 6. 설치 및 실행 방법 (개발 환경)

### 사전 준비
- Node.js LTS 버전 설치 (https://nodejs.org)
- VS Code 설치 (https://code.visualstudio.com)

### 설치 순서
```bash
# 1. AnsonLai 저장소 클론
git clone https://github.com/AnsonLai/Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting.git
cd Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting

# 2. 의존성 설치
npm install

# 3. 개발 서버 실행
npm start
```

---

## 7. 동료 배포 방법 (Sideload)

### 배포 파일 준비
1. 개발 완료 후 폴더 전체를 압축
2. 구글 드라이브 또는 공유 폴더에 업로드

### 동료 설치 순서
1. Node.js 설치
2. 폴더 압축 해제 후 `npm install` 실행
3. `npm start` 실행 (로컬 서버 구동)
4. Word 실행 → 삽입 → 추가 기능 → 내 추가 기능 → `manifest.xml` 파일 로드
5. ⚙️ 아이콘 클릭 → 본인 Gemini API 키 입력

### Gemini API 키 발급 방법 (동료 안내용)
1. https://aistudio.google.com 접속 (Google 계정 로그인)
2. 좌측 메뉴 "Get API key" 클릭
3. "Create API key" 클릭
4. 생성된 키 복사 후 Add-in ⚙️ 설정에 붙여넣기

---


---

## 9. 추가 및 고도화 완료 기능 (2026-05-26)

### 9-1. AI 모델 유연화
- 사용자가 `Gemini 2.5 Flash` 및 `Gemini 2.5 Lite` 모델을 선택 가능.
- API 오류 시 상용 모델(`gemini-2.0-flash`)로 자동 전환되는 Fallback 로직 적용.

### 9-2. 다국어 지원 (Localization)
- 설정 메뉴를 통해 응답 언어(한국어/영어) 선택 가능.
- 시스템 인스트럭션을 통해 선택된 언어에 최적화된 문체 유지.

### 9-3. 멀티 모달 문서 분석
- PDF, Docx, Txt 파일 업로드 지원.
- 업로드된 문서의 스타일을 분석하여 현재 작성 중인 문서에 자연스럽게 녹여내는 요약 기능.

---

## 10. 향후 고도화 제안 (Next Steps)

1. **AI 실무 보정 (Smart Polish)**
   - "격식 있게", "부드럽게", "간결하게" 등 클릭 한 번으로 문장을 다듬는 칩 버튼 추가.
2. **출처 인용 자동화 (Auto-Citation)**
   - 업로드된 참고 문서의 어느 부분에서 해당 내용이 파생되었는지 주석(Comment) 또는 각주로 표기.
3. **사무국 전용 템플릿 라이브러리**
   - 자주 사용하는 기획서, 공문서, 보고서 양식(포맷)을 AI가 학습하여 형식을 100% 맞춰주는 기능.
