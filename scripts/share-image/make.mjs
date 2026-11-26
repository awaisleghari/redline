// Renders share-image.html to public/share-image.jpg (the link-preview picture). Run: npm run make:share-image
import { chromium } from 'playwright';
import { fileURLToPath } from 'node:url';
const here = (name) => fileURLToPath(new URL(name, import.meta.url));
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1200, height: 630 }, deviceScaleFactor: 1 });
await page.goto(`file://${here('share-image.html')}`);
await page.evaluate(() => document.fonts.ready);
await page.screenshot({ path: here('../../public/share-image.jpg'), type: 'jpeg', quality: 88 });
await browser.close();
