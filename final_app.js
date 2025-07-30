const express = require("express");
const bodyParser = require("body-parser");
const { google } = require("googleapis");
const path = require("path");
const puppeteer = require("puppeteer");

const keyFile =
  JSON.parse(process.env.GOOGLE_KEY_JSON) ||
  path.join(__dirname, "package-google-key.json");
const scopes = ["https://www.googleapis.com/auth/spreadsheets"];

const app = express();
const PORT = process.env.PORT || 3000;
app.use(bodyParser.json());

// 오늘의 집에서 순위 조회하는 함수
async function getRankFromOhouse(keyword, mid, page) {
  try {
    let rank = "";

    // keyword가 유효한 경우만 검색
    if (keyword !== undefined && keyword !== null && keyword !== "") {
      const inputSelector =
        "input[placeholder='쇼핑 검색'].css-1pneado.e1rynmtb2";
      await page.waitForSelector(inputSelector);
      await page.type(inputSelector, keyword);
      await page.keyboard.press("Enter");
      console.log("키워드 검색 됨");

      const totalUrls = [];
      let found = false;
      let repeatCount = 0;
      const MAX_REPEAT = 100;
      let prevLastFour = "";

      while (!found) {
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

        if (repeatCount >= MAX_REPEAT) {
          console.log(`keyword: ${keyword}, mid: ${mid} 해당 상품 없음`);
          break;
        }

        // 중복 제거하며 순서대로 저장
        for (const url of newUrls) {
          if (!totalUrls.includes(url)) {
            totalUrls.push(url);
          }
        }

        // 순위 계산
        for (let i = 0; i < totalUrls.length; i++) {
          const match = totalUrls[i].match(
            /productions\/(\d+).*affect_id=(\d+)/
          );
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

      // 검색어 초기화(검색창 클리어)
      const clearBtnSelector = "button.css-ytyqhb.e1rynmtb1";
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

// 시트에 있는 데이터 가져오는 함수
async function getRowsFromSheet(sheets, spreadsheetId, sheetName) {
  const range = `${sheetName}!G7:H`;
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
  });
  return res.data.values || [];
}

// 시트에 순위 데이터 업데이트 하는 함수
async function sendDataToSheet(
  sheets,
  ranks,
  sheetId,
  sheetName,
  spreadsheetId
) {
  const writeRange = `${sheetName}!I6:I${6 + ranks.length}`;
  const date = new Date();
  const rankRowName = date
    .toLocaleString("sv-SE", { hour12: false })
    .slice(2, 16)
    .replace("T", " ");
  const values = [[rankRowName], ...ranks];
  const colorCellRow = 5; // 6행(0부터 시작)
  const colorCellCol = 8; // I열(0부터 시작)

  // 새 열 삽입 + 날짜 셀 색상 지정
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          insertDimension: {
            range: {
              sheetId: sheetId,
              dimension: "COLUMNS",
              startIndex: colorCellCol,
              endIndex: colorCellCol + 1,
            },
            inheritFromBefore: false,
          },
        },
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

// POST 요청 받는 엔드포인트
app.post("/trigger", async (req, res) => {
  const { sheetId, sheetName, spreadsheetId } = req.body;
  if (!sheetId || !sheetName || !spreadsheetId) {
    return res.status(400).json({ error: "필수값 누락" });
  }

  let browser, page;
  try {
    // 구글 인증
    const auth = new google.auth.GoogleAuth({
      keyFile: keyFile,
      scopes: scopes,
    });
    const sheets = google.sheets({ version: "v4", auth });

    // 시트 데이터 읽기
    const rows = await getRowsFromSheet(sheets, spreadsheetId, sheetName);

    browser = await puppeteer.launch({
      headless: true,
      args: ["--no-sandbox", "--disable-setuid-sandbox"]
    });
    page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 800 });
    await page.setUserAgent(
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
    );
    await page.goto("https://store.ohou.se/", { waitUntil: "networkidle2" });
    console.log("오늘의 집-쇼핑 페이지 열림");

    let ranks = [];
    for (const [keyword, mid] of rows) {
      const rank = await getRankFromOhouse(keyword, mid, page);
      ranks.push([rank]);
      console.log(`keyword: ${keyword}, mid: ${mid}, rank: ${rank}`);
    }
    await browser.close();
    console.log("순위 조회 완료!");

    await sendDataToSheet(sheets, ranks, sheetId, sheetName, spreadsheetId);
    console.log("순위 업데이트 완료!");

    return res.json({ status: "success" });
  } catch (e) {
    if (browser) await browser.close();
    console.error(e);
    return res.status(500).json({ error: e.message });
  }
});

// 서버 실행
app.listen(PORT, () => {
  console.log(`서버가 실행중입니다. 포트: ${PORT}`);
});
