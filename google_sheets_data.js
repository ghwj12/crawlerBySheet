const { google } = require("googleapis");
const path = require("path");

async function readSheet() {
  // 구글 인증: 서비스 계정 키 파일로 인증
  const auth = new google.auth.GoogleAuth({
    keyFile: path.join(__dirname, "package-google-key.json"), // 구글 시트 API 계정 키 파일
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  // 클라이언트 객체
  const client = await auth.getClient();

  // 구글 시트 API 클라이언트 생성
  const sheets = google.sheets({ version: "v4", auth: client });

  // 읽어올 시트ID와 범위 지정
  const spreadsheetId = "1Yqd9vsAg2LMR7IveQ1xUJMjaU2tW0NH9tT6-yy330yY"; // 시트 url에서 d/와 /edit 사이 문자열
  const range = "오집!G7:H"; // '오집' 시트의 G7행부터 H열까지 모든 행

  // 데이터 요청
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,         
  });

  // 받은 데이터(2차원 배열 형태)
  const rows = res.data.values;
  if (rows && rows.length) {
    rows.forEach(([keyword, mid], i) => {
      console.log(`row ${i + 7} → keyword: ${keyword}, mid: ${mid}`);
    });
  } else {
    console.log("데이터가 없습니다.");
  }
}

readSheet().catch(console.error);
