const puppeteer = require("puppeteer");

(async () => {
  let browser;
  try {
    browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();

    // headless: true일 때 모바일 뷰로 판단되어 오류나는 것을 방지하기 위해
    // 데스크탑 환경처럼 설정
    await page.setViewport({ width: 1280, height: 800 });
    await page.setUserAgent(
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
    );

    const keyword = "층간소음매트";
    const mid = "3127444";

    await page.goto(
      `https://ohou.se/search/index?query=${encodeURIComponent(keyword)}`,
      { waitUntil: "networkidle2" }
    );

    // 더보기 버튼 클릭
    await page.click(".search-item__more-btn--pc");
    console.log("클릭 됨");

    const totalUrls = [];
    let found = false;
    let rank = 0;

    await page.waitForSelector(".css-18yij6n.eh8ehv77");

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
        console.log("해당 상품 없음");
        break;
      }

      // 중복 제거하며 순서대로 저장
      for (const url of newUrls) {
        if (!totalUrls.includes(url)) {
          totalUrls.push(url);
          // console.log(`url ${url}`);
        }
      }

      // 순위 계산
      for (let i = 0; i < totalUrls.length; i++) {
        const match = totalUrls[i].match(/productions\/(\d+).*affect_id=(\d+)/);
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

    console.log(`상품 순위: ${rank}`);
  } catch (e) {
    console.error("오류 발생:", e.message);
  } finally {
    if (browser) await browser.close();
  }
})();
