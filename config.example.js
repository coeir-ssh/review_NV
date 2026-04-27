/**
 * 설정 파일 예시
 * 이 파일을 config.js로 복사한 후 실제 값을 채워 넣으세요.
 *
 * ANTHROPIC_API_KEY: https://console.anthropic.com 에서 발급
 * NAVER_CLIENT_ID / NAVER_CLIENT_SECRET: 스마트스토어 개발자센터 앱 등록 후 발급
 * SLACK_BOT_TOKEN: https://api.slack.com/apps 에서 발급 (Bot Token Scopes에 chat:write 추가)
 * DATA_GO_KR_KEY: https://www.data.go.kr/data/15012690/openapi.do 에서 발급
 */
module.exports = {
  ANTHROPIC_API_KEY: process.env.ANTHROPIC_API_KEY || 'sk-ant-...',
  NAVER_CLIENT_ID:     process.env.NAVER_CLIENT_ID     || 'YOUR_NAVER_CLIENT_ID',
  NAVER_CLIENT_SECRET: process.env.NAVER_CLIENT_SECRET || 'YOUR_NAVER_CLIENT_SECRET',
  SELLER_ID: process.env.SELLER_ID || 'your_seller_id@example.com',
  SELLER_PW: process.env.SELLER_PW || 'YOUR_SELLER_PW',
  SLACK_BOT_TOKEN:   process.env.SLACK_BOT_TOKEN   || 'xoxb-...',
  SLACK_USER_ID:     process.env.SLACK_USER_ID     || 'U........',
  SLACK_CHANNEL_ID:  process.env.SLACK_CHANNEL_ID  || 'C........',
  DATA_GO_KR_KEY:    process.env.DATA_GO_KR_KEY    || 'YOUR_DATA_GO_KR_KEY',
};
