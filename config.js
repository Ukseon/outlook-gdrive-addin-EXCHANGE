// ===== 설정 =====
window.ADDIN_CONFIG = {
  // Google Cloud Console에서 만든 OAuth 클라이언트 ID
  GOOGLE_CLIENT_ID: 'YOUR_CLIENT_ID.apps.googleusercontent.com',

  // 회사 Google Workspace 도메인. 이 도메인 계정으로만 로그인할 수 있고,
  // 외부 공유가 막혀 있을 때 '사내 공유'로 대신 설정합니다.
  HOSTED_DOMAIN: 'daimtechnology.com',

  // 업로드 위치. '' = 내 드라이브 최상위.
  FOLDER_NAME: 'Outlook 대용량 첨부',

  // 분할 업로드 단위(MB)
  CHUNK_MB: 8
};
