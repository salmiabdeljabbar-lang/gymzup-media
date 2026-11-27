// Usage: node render.js <file.html> <outdir> <basename>
// Writes <basename>-US-Letter.pdf, <basename>-A4.pdf and page PNG previews.
const { chromium } = require("playwright");
const path = require("path");
const fs = require("fs");

(async () => {
  const [file, outdir, base] = process.argv.slice(2);
  fs.mkdirSync(path.join(outdir, "pages"), { recursive: true });
  const browser = await chromium.launch();
  const url = "file://" + path.resolve(file);
  for (const [size, label] of [["letter", "US-Letter"], ["a4", "A4"]]) {
    const page = await browser.newPage();
    await page.goto(url + "?size=" + size);
    await page.evaluate(() => document.fonts.ready);
    await page.pdf({ path: path.join(outdir, `${base}-${label}.pdf`), preferCSSPageSize: true, printBackground: true });
    if (size === "letter") {
      await page.setViewportSize({ width: 816, height: 1056 });
      const pages = await page.$$(".page");
      for (let i = 0; i < pages.length; i++) {
        await pages[i].screenshot({ path: path.join(outdir, "pages", `page-${i + 1}.png`) });
      }
    }
    await page.close();
  }
  await browser.close();
  console.log("done");
})();
