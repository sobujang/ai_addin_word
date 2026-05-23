const express = require('express');
const https = require('https');
const path = require('path');

const PORT = 3000;

async function startServer() {
  const app = express();
  app.use(express.static(path.join(__dirname, 'public')));

  let httpsOptions;
  try {
    httpsOptions = await require('office-addin-dev-certs').getHttpsServerOptions();
  } catch (e) {
    console.error('\n❌ HTTPS 인증서 오류. 아래 명령어를 실행한 후 다시 시도하세요:');
    console.error('   npx office-addin-dev-certs install\n');
    console.error('(Windows의 경우 관리자 권한으로 실행 필요)\n');
    process.exit(1);
  }

  const server = https.createServer(httpsOptions, app);

  server.listen(PORT, () => {
    console.log('\n✅ Gemini AI Word Add-in 서버 실행 중');
    console.log(`📍 주소: https://localhost:${PORT}/taskpane.html`);
    console.log('\n📌 Word에서 사용하는 방법:');
    console.log('   1. Microsoft Word 실행');
    console.log('   2. 삽입 → 추가 기능 → 내 추가 기능');
    console.log('   3. manifest.xml 파일 로드');
    console.log('   4. ⚙️ 아이콘 클릭 → Gemini API 키 입력');
    console.log('\nCtrl+C 로 서버를 종료할 수 있습니다.\n');
  });
}

startServer().catch(console.error);
