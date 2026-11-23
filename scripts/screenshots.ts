/**
 * Takes screenshots of the page for the owner to review (an "owner check").
 * Needs the built site being served: run `npm run build && npm run preview`
 * in another terminal first.
 *
 * Usage: npx tsx scripts/screenshots.ts <output folder>
 */
import { chromium } from '@playwright/test';
import { mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { buildPair, pairById } from '../tests/fixtures/pairs';

const folder = process.argv[2];
if (!folder) throw new Error('usage: screenshots.ts <output folder>');
mkdirSync(folder, { recursive: true });
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, deviceScaleFactor: 2 });
const shot = (name: string): Promise<Buffer> => page.screenshot({ path: join(folder, name), fullPage: true });

await page.goto('http://localhost:4173/');
await shot('1 - empty page.png');

const pair = await buildPair(pairById('several-changes'));
await page.locator('#original').setInputFiles({ name: 'Services Agreement v1.docx', mimeType: DOCX, buffer: Buffer.from(pair.original) });
await page.locator('#modified').setInputFiles({ name: 'Services Agreement v2.docx', mimeType: DOCX, buffer: Buffer.from(pair.modified) });
await shot('2 - files chosen.png');
await page.getByRole('button', { name: 'Compare' }).click();
await page.locator('#status[data-state="done"]').waitFor({ timeout: 60_000 });
await page.frameLocator('#viewer').getByText('Comparison of documents').waitFor();
await shot('3 - after comparing.png');

await page.goto('http://localhost:4173/');
await page.locator('#original').setInputFiles({ name: 'Old contract.doc', mimeType: 'application/msword', buffer: Buffer.from('x') });
await page.locator('#modified').setInputFiles({ name: 'Services Agreement v2.docx', mimeType: DOCX, buffer: Buffer.from(pair.modified) });
await page.getByRole('button', { name: 'Compare' }).click();
await page.locator('#status[data-state="refused"]').waitFor();
await shot('4 - a refused file.png');

await page.setViewportSize({ width: 420, height: 900 });
await page.goto('http://localhost:4173/');
await shot('5 - narrow screen.png');
await browser.close();
console.log(`screenshots written to ${folder}`);
