const reader = require("xlsx");
const puppeteer = require("puppeteer-extra");
const StealthPlugin = require("puppeteer-extra-plugin-stealth");
const userAgent = require("user-agents");
const Excel = require("exceljs");
const cheerio = require("cheerio");
const proxyChain = require("proxy-chain");
const fs = require("fs");
puppeteer.use(StealthPlugin());

const sleep = (milliseconds) => {
  const date = Date.now();
  let currentDate = null;
  do {
    currentDate = Date.now();
  } while (currentDate - date < milliseconds);
};

const ScrapeTickers = async () => {
  try {
    const Tickers = [];
    const tickers = reader.readFile(__dirname + "/Tickers.xlsx");
    const keyField = reader.utils.sheet_to_csv(
      tickers.Sheets[tickers.SheetNames[0]]
    );
    keyField.split("\n").forEach((row) => {
      Tickers.push(row);
    });
    var Data = [];
    const oldProxyUrl =
      "http://lum-customer-c_ef78a635-zone-data_center:qid5pp9zjd0e@zproxy.lum-superproxy.io:22225";
    const newProxyUrl = await proxyChain.anonymizeProxy(oldProxyUrl);
    const browser = await puppeteer.launch({
      // headless: false,
      executablePath:
        "C://Program Files (x86)//Google//Chrome//Application//chrome.exe",
      ignoreDefaultArgs: ["--disable-extensions", "--enable-automation"],
      args: ["--start-maximized", `--proxy-server=${newProxyUrl}`],
      ignoreHTTPSErrors: true,
      slowMo: 100,
    });
    console.log("Browser Opened ( In Headless Mode )");
    const page = await browser.newPage();
    await page.setDefaultNavigationTimeout(0);
    await page.viewport({
      width: 1024 + Math.floor(Math.random() * 100),
      height: 768 + Math.floor(Math.random() * 100),
    });
    await page.setRequestInterception(true);
    page.on("request", (req) => {
      const url = req.url();
      if (
        req.resourceType() == "stylesheet" ||
        req.resourceType() == "font" ||
        req.resourceType() == "image" ||
        url.endsWith("uat.js")
      ) {
        req.abort();
      } else {
        req.continue();
      }
    });
    await page.setUserAgent(userAgent.toString());
    const href = "https://www.etfscreen.com/";
    console.log("Loading: https://www.etfscreen.com/");
    await page.goto(href, { timeout: 0 });
    console.log("Page Loaded...");
    for (let i = 0; i <= Tickers.length - 2; i++) {
      console.log("Loading the Ticker...");
      const ticker = Tickers[i];
      await page.type("#headerSym", `${ticker}`);
      await page.$eval('form[name="headerSym"]', (form) => form.submit());
      console.log(`Getting the Required Data for ${ticker}`);
      await page.waitForSelector("#headerSym", { timeout: 0 });
      const content = await page.content();
      const $ = cheerio.load(content);
      const m1 = $(
        "body > table.mTable > tbody > tr > td.mPanel > div:nth-child(10) > div:nth-child(2) > table > tbody > tr:nth-child(2) > td:nth-child(3) > table > tbody > tr:nth-child(4) > td.taR"
      ).text();
      let m1__text = m1;
      const m3 = $(
        "body > table.mTable > tbody > tr > td.mPanel > div:nth-child(10) > div:nth-child(2) > table > tbody > tr:nth-child(2) > td:nth-child(3) > table > tbody > tr:nth-child(5) > td.taR"
      ).text();
      let m3__text = m3;
      const m6 = $(
        "body > table.mTable > tbody > tr > td.mPanel > div:nth-child(10) > div:nth-child(2) > table > tbody > tr:nth-child(2) > td:nth-child(3) > table > tbody > tr:nth-child(6) > td.taR"
      ).text();
      let m6__text = m6;

      const _1 = m1__text.split("%");
      const _3 = m3__text.split("%");
      const _6 = m6__text.split("%");

      const MEAN = String((Number(_1[0]) + Number(_3[0]) + Number(_6[0])) / 3);
      if (MEAN.length >= 5) {
        const Mean = MEAN.slice(0, 5) + "%";
        Data[i] = {
          ticker: ticker,
          m1: m1__text,
          m3: m3__text,
          m6: m6__text,
          mean: Mean,
        };
      } else {
        const Mean = MEAN + "%";
        Data[i] = {
          ticker: ticker,
          m1: m1__text,
          m3: m3__text,
          m6: m6__text,
          mean: Mean,
        };
      }
      console.log(Data[i]);
    }
    console.log("All the Data Has Been Scraped Successfully...");
    await browser.close();
    return Data;
  } catch (error) {
    console.log(error);
  }
};

const insertData = async () => {
  try {
    fs.exists(__dirname + "/Data.xlsx", function (exists) {
      if (exists) {
        fs.unlinkSync(__dirname + "/Data.xlsx");
      }
    });
    console.log("Getting the Tickers from Tickers.xlsx");
    const Data = await ScrapeTickers();
    let workbook = new Excel.Workbook();
    let worksheet = workbook.addWorksheet("sheet1");
    worksheet.columns = [
      { header: "Ticker", key: "ticker" },
      { header: "1-Month", key: "m1" },
      { header: "3-Month", key: "m3" },
      { header: "6-Month", key: "m6" },
      { header: "Mean", key: "mean" },
    ];
    worksheet.columns.forEach((column) => {
      column.width = column.header.length < 12 ? 12 : column.header.length;
    });
    worksheet.getRow(1).font = { bold: true };
    Data.forEach((e) => {
      worksheet.addRow({
        ...e,
      });
    });
    workbook.xlsx.writeFile("Data.xlsx");
  } catch (err) {
    console.log(err);
  }
};

insertData();
