const { google } = require("googleapis");
const path = require("path");
const puppeteer = require("puppeteer");
let browser, page;

// 구글 인증: 서비스 계정 키 파일로 인증
const keyFile = path.join(__dirname, "package-google-key.json"); // 구글 시트 API 계정 키 파일
const scopes = ["https://www.googleapis.com/auth/spreadsheets"];

// 읽어올 시트ID와 범위 지정
const spreadsheetId = "1Yqd9vsAg2LMR7IveQ1xUJMjaU2tW0NH9tT6-yy330yY"; // 시트 url에서 d/와 /edit 사이 문자열
const sheetName = "오집";
const range = `${sheetName}!G7:H`; // 시트의 G7행부터 H열까지 모든 행

// 시트에 있는 데이터 가져오는 함수
async function getRowsFromSheet(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
  });

  return res.data.values || [];
}

// 시트의 sheetId 가져오는 함수
async function getSheetIdByName(sheets, sheetName) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const found = meta.data.sheets.find((s) => s.properties.title === sheetName);
  if (!found) throw new Error(`시트 "${sheetName}"를 찾을 수 없습니다`);
  return found.properties.sheetId;
}

// 오늘의 집-쇼핑 페이지 여는 함수
async function pageOpen() {
  try {
    browser = await puppeteer.launch({ headless: false });
    page = await browser.newPage();

    // headless: true일 때 모바일 뷰로 판단되어 오류나는 것을 방지하기 위해
    // 데스크탑 환경처럼 설정
    await page.setViewport({ width: 1280, height: 800 });
    await page.setUserAgent(
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
    );
    await page.goto("https://store.ohou.se/", { waitUntil: "networkidle2" });
    console.log("오늘의 집-쇼핑 페이지 열림");
    return true;
  } catch (e) {
    if (browser) await browser.close();
    return false;
  }
}

// 오늘의 집에서 순위 조회하는 함수
async function getRankFromOhouse(keyword, mid) {
  try {
    let rank = "";

    // keyword가 빈 값인지 확인
    if (keyword !== undefined && keyword !== null && keyword !== "") {
      const inputSelector =
        "input[placeholder='쇼핑 검색'].css-1pneado.e1rynmtb2";
      await page.waitForSelector(inputSelector);
      await page.type(inputSelector, keyword);
      await page.keyboard.press("Enter");

      console.log("키워드 검색 됨");

      const totalUrls = new Set();
      let found = false;

      let repeatCount = 0;
      const MAX_REPEAT = 100;
      let prevLastFour = "";

      while (!found) {
        // 최소 하나의 상품은 존재할 때까지 대기
        await page.waitForFunction(() => {
          return (
            document.querySelectorAll(
              ".production-feed__item-wrap.col-6.col-md-4.col-lg-3"
            ).length > 0
          );
        });

        // 새로 추가된 상품 url만 수집
        const newUrls = await page.evaluate(() => {
          const elements = Array.from(
            document.querySelectorAll(
              ".production-feed__item-wrap.col-6.col-md-4.col-lg-3"
            )
          );
          return elements
            .map((el) => el.querySelector("a")?.getAttribute("href"))
            .filter(Boolean);
        });

        // newUrls 마지막 4개 추출
        const lastFour = newUrls.slice(-4).join(",");

        // 직전과 같으면 카운트, 다르면 초기화
        if (lastFour === prevLastFour) {
          repeatCount++;
        } else {
          repeatCount = 1;
          prevLastFour = lastFour;
        }

        // 페이지 끝까지 도달 했는지 확인
        if (repeatCount >= MAX_REPEAT) {
          console.log(`keyword: ${keyword}, mid: ${mid} 해당 상품 없음`);
          break;
        }

        // 중복 제거하며 순서대로 저장
        for (const url of newUrls) {
          totalUrls.add(url);
        }

        // 순위 계산
        for (let url of totalUrls) {
          const match = url.match(/productions\/(\d+).*affect_id=(\d+)/);
          if (match && match[1] === mid) {
            rank = match[2];
            found = true;
            break;
          }
        }

        if (found) break;

        // 스크롤 아래로
        await page.evaluate(() => {
          window.scrollBy(0, window.innerHeight);
        });
      }

      // 스크롤 맨 위로
      await page.evaluate(() => {
        window.scrollTo(0, 0);
      });

      // 검색어 초기화
      const clearBtnSelector = "button.css-ytyqhb.e1rynmtb1";

      // 버튼 존재 여부 확인
      const isBtnVisible = await page.$(clearBtnSelector);
      if (isBtnVisible) {
        await page.click(clearBtnSelector);
      }
    }

    return rank || "";
  } catch (e) {
    return ""; // 오류시 빈 값
  }
}

// 구글 시트에 순위 데이터 업데이트 하는 함수
async function sendDataToSheet(sheets, ranks) {
  const sheetId = await getSheetIdByName(sheets, sheetName);
  // 구글 시트의 순위 열
  const writeRange = `${sheetName}!I6:I${6 + ranks.length}`;
  const date = new Date();
  const rankRowName = date
    .toLocaleString("sv-SE", { hour12: false })
    .slice(2, 16)
    .replace("T", " ");
  const values = [[rankRowName], ...ranks];
  const colorCellRow = 5; // 6행(0부터 시작)
  const colorCellCol = 8; // I열(0부터 시작)

  // 구글 시트에 순위 열 추가
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          insertDimension: {
            range: {
              sheetId: sheetId,
              dimension: "COLUMNS",
              startIndex: colorCellCol, // I열(0부터 시작하는 열 번호)
              endIndex: colorCellCol + 1,
            },
            inheritFromBefore: false,
          },
        },
        // 날짜 셀 색상 채우기
        {
          repeatCell: {
            range: {
              sheetId: sheetId,
              startRowIndex: colorCellRow,
              endRowIndex: colorCellRow + 1,
              startColumnIndex: colorCellCol,
              endColumnIndex: colorCellCol + 1,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: {
                  red: 1,
                  green: 0.949,
                  blue: 0.8,
                },
              },
            },
            fields: "userEnteredFormat.backgroundColor",
          },
        },
      ],
    },
  });

  // 시트에 데이터 입력
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: writeRange,
    valueInputOption: "RAW",
    requestBody: { values },
  });
}

// 시작
(async () => {
  // 구글 인증
  const auth = new google.auth.GoogleAuth({
    keyFile: keyFile,
    scopes: scopes,
  });

  const sheets = google.sheets({ version: "v4", auth });

  const rows = await getRowsFromSheet(sheets);
  let ranks = [];

  //오늘의 집 페이지 키워드 검색 준비
  const pageYn = await pageOpen();
  if (pageYn) {
    // 각각의 순위 조회
    for (const [keyword, mid] of rows) {
      const rank = await getRankFromOhouse(keyword, mid);
      ranks.push([rank]);
      console.log(`keyword: ${keyword}, mid: ${mid}, rank: ${rank}`);
    }
    await browser.close();
    console.log("순위 조회 완료!");

    await sendDataToSheet(sheets, ranks);
    console.log("순위 업데이트 완료!");
  } else {
    console.log("오늘의 집 로딩 오류");
    ranks.push(["오늘의 집 로딩 오류 발생"]);
    await sendDataToSheet(sheets, ranks);
  }
})();
