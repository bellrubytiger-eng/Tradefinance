// 서식 이미지 캡처 (서식 명세서 삽입용): node tools/shot_forms.js <out_dir>
const { chromium } = require("playwright");
const path = require("path");
const out = process.argv[2];
const forms = "file://" + path.resolve(__dirname, "../screens/forms.html");
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1000, height: 1300 }, deviceScaleFactor: 1.4 });
  for (const h of ["firc-fdi", "out-imp", "in-exp", "fx-acct"]) {
    await p.goto(`${forms}#${h}`);
    await p.waitForTimeout(400);
    await (await p.$("#page")).screenshot({ path: path.join(out, `${h}.png`) });
  }
  await b.close();
  console.log("saved", out);
})();
