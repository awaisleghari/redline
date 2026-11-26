/**
 * Browser tests of the landing page (/) and the docs (/docs, /docs/<article>).
 * The workspace is in compare.spec.ts.
 */
import { expect, test, type Page } from '@playwright/test';

function watchForErrors(page: Page): string[] {
  const errors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });
  page.on('pageerror', (error) => errors.push(error.message));
  return errors;
}

/** Whether anything is drawn on a canvas, and a fingerprint of what. */
const inkOf = (page: Page, id: string): Promise<{ lit: number; sum: number }> =>
  page.locator(`#${id}`).evaluate((canvas: HTMLCanvasElement) => {
    const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
    let lit = 0;
    let sum = 0;
    for (let i = 3; i < data.length; i += 4) {
      if (data[i]! > 0) lit++;
      sum += data[i]! * ((i % 977) + 1);
    }
    return { lit, sum };
  });

test.describe('the landing page', () => {
  test('shows the headline over two animated grids, with the two buttons leading to the workspace and the docs', async ({ page, baseURL }) => {
    const errors = watchForErrors(page);
    const outside: string[] = [];
    page.on('request', (request) => {
      if (!request.url().startsWith(new URL(baseURL!).origin + '/') && !request.url().startsWith('data:')) outside.push(request.url());
    });
    await page.setViewportSize({ width: 1440, height: 900 });
    await page.goto('/');
    await expect(page).toHaveTitle('redline — For the paragraphs someone tweaked at 11:58 PM.');
    await expect(page.locator('link[rel="icon"]')).toHaveAttribute('href', '/redline-favicon.svg');
    expect((await page.request.get('/redline-favicon.svg')).headers()['content-type']).toContain('image/svg+xml');
    const headline = page.getByRole('heading', { level: 1 });
    await expect(headline).toHaveText('For the paragraphs someone tweaked at 11:58 PM.');
    await expect(headline).toHaveCSS('font-family', /Newsreader/);
    await expect(headline).toHaveCSS('font-size', '60px');
    await expect(headline.locator('span')).toHaveCSS('white-space', 'nowrap');
    await expect(page.getByRole('img', { name: 'redline' })).toBeVisible();

    // Two links of exactly the same size.
    const start = page.getByRole('link', { name: 'get started' });
    const docs = page.getByRole('link', { name: 'docs', exact: true });
    await expect(start).toHaveAttribute('href', '/app');
    await expect(docs).toHaveAttribute('href', '/docs');
    for (const link of [start, docs]) {
      const box = (await link.boundingBox())!;
      expect([box.width, box.height]).toEqual([132, 40]);
    }
    await start.focus();
    await expect(start).toHaveCSS('outline-color', 'rgb(85, 133, 255)');

    // The fonts come from this site, and both grids are drawn and moving.
    await page.evaluate(() => document.fonts.ready);
    expect(await page.evaluate(() => document.fonts.check("500 60px 'Newsreader'") && document.fonts.check("500 14px 'Geist'"))).toBe(true);
    for (const id of ['grid-top', 'grid-bottom']) {
      const before = await inkOf(page, id);
      expect(before.lit).toBeGreaterThan(10_000);
      await expect.poll(async () => (await inkOf(page, id)).sum, { timeout: 5_000 }).not.toBe(before.sum);
    }
    // The grids are 560px tall as designed (a few pixels more when the window is a little taller than the page).
    const heights = await page.locator('.rl-hero').evaluateAll((all) => all.map((hero) => hero.getBoundingClientRect().height));
    expect(heights[0]).toBeGreaterThanOrEqual(560);
    expect(heights[0]).toBeLessThan(575);
    expect(heights[1]).toBe(heights[0]);
    expect(outside).toEqual([]);
    expect(errors).toEqual([]);

    await start.click();
    await expect(page).toHaveURL(/\/app$/);
    await expect(page.getByRole('heading', { name: 'create a redline' })).toBeVisible();
    // The workspace's logo leads back.
    await page.getByRole('link', { name: 'redline home' }).click();
    await expect(page).toHaveURL(/\/$/);
    await page.getByRole('link', { name: 'docs', exact: true }).click();
    await expect(page).toHaveURL(/\/docs$/);
    await expect(page.getByRole('heading', { level: 1, name: 'Overview' })).toBeVisible();
  });

  test('on a phone: smaller logo, buttons and grids, three headline lines at most, and no sideways scrolling', async ({ page }) => {
    await page.setViewportSize({ width: 390, height: 844 });
    await page.goto('/');
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
    expect(await page.locator('.rl-hero').first().evaluate((hero) => hero.getBoundingClientRect().height)).toBeGreaterThanOrEqual(400);
    expect((await page.getByRole('link', { name: 'get started' }).boundingBox())!.width).toBe(104);
    expect((await page.getByRole('img', { name: 'redline' }).boundingBox())!.height).toBe(24);
    const headline = page.getByRole('heading', { level: 1 });
    await expect(headline).toHaveCSS('font-size', '32px');
    expect((await headline.boundingBox())!.height).toBeLessThanOrEqual(32 * 1.06 * 3 + 1);
  });

  test('on a large display the grids grow to fill the window, with the headline in the middle and no empty strip below', async ({ page }) => {
    for (const [width, height] of [[2560, 1340], [1920, 1080], [1440, 2000]] as const) {
      await page.setViewportSize({ width, height });
      await page.goto('/');
      const bottom = await page.locator('.rl-hero').last().evaluate((hero) => hero.getBoundingClientRect().bottom);
      expect(Math.round(bottom)).toBe(height);
      expect(await page.evaluate(() => document.documentElement.scrollHeight)).toBe(height);
      const headline = (await page.getByRole('heading', { level: 1 }).boundingBox())!;
      expect(Math.abs(headline.y + headline.height / 2 - height / 2)).toBeLessThan(2);
      // The bottom grid is drawn right down to the edge of the window.
      const lit = await page.locator('#grid-bottom').evaluate((canvas: HTMLCanvasElement) => {
        const data = canvas.getContext('2d')!.getImageData(0, canvas.height - 12, canvas.width, 12).data;
        let count = 0;
        for (let i = 3; i < data.length; i += 4) if (data[i]! > 40) count++;
        return count;
      });
      expect(lit).toBeGreaterThan(1000);
    }
  });

  test('the grids stay in the three brand colours: no square washes out to white, even at the bright edges', async ({ page }) => {
    await page.setViewportSize({ width: 1440, height: 900 });
    await page.goto('/');
    // Let the animation run until the bright edges are as busy as they get.
    await page.waitForTimeout(3500);
    for (const id of ['grid-top', 'grid-bottom']) {
      const { coloured, pale } = await page.locator(`#${id}`).evaluate((canvas: HTMLCanvasElement) => {
        const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
        let coloured = 0;
        let pale = 0;
        for (let i = 0; i < data.length; i += 4) {
          const alpha = data[i + 3]! / 255;
          // The colour as it appears over the page's dark background.
          const [r, g, b] = [data[i]! * alpha, data[i + 1]! * alpha, data[i + 2]! * alpha];
          if (Math.max(r, g, b) < 60) continue;
          coloured++;
          // Pale: bright, with its weakest channel close to its strongest (white, or a washed-out pastel).
          if (Math.min(r, g, b) > 150) pale++;
        }
        return { coloured, pale };
      });
      expect(coloured).toBeGreaterThan(5_000);
      expect(pale).toBe(0);
    }
  });

  test('with reduced motion the grids are drawn once and stay still', async ({ page }) => {
    await page.emulateMedia({ reducedMotion: 'reduce' });
    await page.goto('/');
    const before = await inkOf(page, 'grid-top');
    expect(before.lit).toBeGreaterThan(10_000);
    await page.waitForTimeout(700);
    expect((await inkOf(page, 'grid-top')).sum).toBe(before.sum);
  });
});

const ARTICLES = [
  ['overview', 'get started', 'Overview'],
  ['quick-start', 'get started', 'Quick start'],
  ['reading-a-redline', 'get started', 'Reading a redline'],
  ['downloads', 'get started', 'Downloads'],
  ['supported-documents', 'reference', 'Supported documents'],
  ['keyboard-shortcuts', 'reference', 'Keyboard shortcuts'],
] as const;
const SECTION_COLOURS = { 'get started': 'rgb(255, 142, 144)', reference: 'rgb(169, 193, 255)' };
/** Written but not published yet (`hidden: yes`): the whole "how it works" section. */
const HIDDEN = [
  ['why-it-works-this-way', 'Why it works this way'],
  ['the-checks', 'The checks'],
  ['privacy', 'Privacy'],
] as const;

test('a shared link shows a title, a description and the picture, without running any script', async ({ request }) => {
  // Link previews (WhatsApp, LinkedIn, X, iMessage…) read the page as sent, so the tags must be in the HTML itself.
  const tag = (html: string, key: string): string | undefined => new RegExp(`<meta (?:property|name)="${key}" content="([^"]*)"`).exec(html)?.[1];
  for (const [path, title] of [['/', 'redline — For the paragraphs someone tweaked at 11:58 PM.'], ['/docs', 'redline — docs'], ['/docs/quick-start', 'redline — docs'], ['/app', 'redline — workspace']] as const) {
    const html = await (await request.get(path)).text();
    expect(tag(html, 'og:title'), path).toBe(title);
    expect(tag(html, 'og:description')?.length, path).toBeGreaterThan(40);
    expect(tag(html, 'og:image'), path).toBe('https://useredline.dev/share-image.jpg');
    expect(tag(html, 'og:url'), path).toMatch(/^https:\/\/useredline\.dev\//);
    expect(tag(html, 'twitter:card'), path).toBe('summary_large_image');
  }
  // The picture itself: served by the site, a JPEG, and small enough for WhatsApp (which drops large ones).
  const image = await request.get('/share-image.jpg');
  expect(image.status()).toBe(200);
  expect(image.headers()['content-type']).toBe('image/jpeg');
  const bytes = await image.body();
  expect(bytes.length).toBeLessThan(300_000);
  // Width and height from the JPEG's frame header: 1200 × 630.
  let at = 2;
  while (at < bytes.length && !(bytes[at] === 0xff && bytes[at + 1]! >= 0xc0 && bytes[at + 1]! <= 0xc2)) at += 2 + bytes.readUInt16BE(at + 2);
  expect([bytes.readUInt16BE(at + 7), bytes.readUInt16BE(at + 5)]).toEqual([1200, 630]);
});

test.describe('the docs', () => {
  test('every article has its own address, its section colour in the menu, and previous/next in menu order', async ({ page }) => {
    const errors = watchForErrors(page);
    await page.setViewportSize({ width: 1440, height: 900 });
    for (const [index, [slug, section, title]] of ARTICLES.entries()) {
      await page.goto(`/docs/${slug}`);
      await expect(page.getByRole('heading', { level: 1 })).toHaveText(title);
      await expect(page).toHaveTitle(`${title} — redline docs`);
      await expect(page.locator('#eyebrow')).toHaveText(section);
      await expect(page.locator('#eyebrow')).toHaveCSS('color', SECTION_COLOURS[section]);
      await expect(page.locator('#crumb-section')).toHaveText(section);
      const current = page.locator('#docs-nav [aria-current="page"]');
      await expect(current).toHaveCount(1);
      await expect(current).toHaveText(new RegExp(`^${title}`, 'i'));
      await expect(current).toHaveCSS('color', SECTION_COLOURS[section]);
      await expect(page.locator('#crumb-title')).toHaveText((await current.locator('span').first().textContent())!);
      const previous = page.locator('#prev-article');
      const next = page.locator('#next-article');
      if (index === 0) await expect(previous).toBeHidden();
      else await expect(previous).toHaveAttribute('href', index === 1 ? '/docs' : `/docs/${ARTICLES[index - 1]![0]}`);
      if (index === ARTICLES.length - 1) await expect(next).toBeHidden();
      else await expect(next).toHaveAttribute('href', `/docs/${ARTICLES[index + 1]![0]}`);
      // "On this page" lists the article's headings; the first is marked.
      const toc = page.locator('#toc a');
      expect(await toc.count()).toBeGreaterThan(0);
      await expect(toc.first()).toHaveAttribute('aria-current', 'true');
      expect(await page.locator('#article-body h2, #article-body [data-toc]').count()).toBe(await toc.count());
    }
    expect(errors).toEqual([]);
  });

  test('moving between articles changes the address, returns to the top, and the back button works', async ({ page }) => {
    await page.setViewportSize({ width: 1440, height: 900 });
    await page.goto('/docs');
    await expect(page.getByRole('heading', { level: 1 })).toHaveText('Overview');
    await page.locator('#reader').evaluate((reader) => reader.scrollTo(0, 400));
    await page.locator('#docs-nav').getByRole('link', { name: 'downloads' }).click();
    await expect(page).toHaveURL(/\/docs\/downloads$/);
    await expect(page.getByRole('heading', { level: 1 })).toHaveText('Downloads');
    expect(await page.locator('#reader').evaluate((reader) => reader.scrollTop)).toBe(0);
    await page.locator('#next-article').click();
    await expect(page).toHaveURL(/\/docs\/supported-documents$/);
    await page.goBack();
    await expect(page.getByRole('heading', { level: 1 })).toHaveText('Downloads');
    await page.goBack();
    await expect(page).toHaveURL(/\/docs$/);
    await expect(page.getByRole('heading', { level: 1 })).toHaveText('Overview');
    // Out of the docs: the workspace, the source code, the landing page.
    await expect(page.getByRole('link', { name: 'open redline' })).toHaveAttribute('href', '/app');
    await expect(page.getByRole('link', { name: 'source', exact: true })).toHaveAttribute('href', 'https://github.com/awaisleghari/redline');
    await expect(page.getByRole('link', { name: 'source code' })).toHaveAttribute('href', 'https://github.com/awaisleghari/redline');
    await expect(page.getByRole('link', { name: 'redline home' })).toHaveAttribute('href', '/');
    await expect(page.locator('#docs-nav .rl-badge')).toHaveText('soon');
  });

  test('the hidden articles are nowhere: no menu entry, no address, no search result, and links to them are plain words', async ({ page }) => {
    await page.goto('/docs');
    await expect(page.locator('#docs-nav .rl-navtitle')).toHaveText(['get started', 'reference']);
    for (const [slug, title] of HIDDEN) {
      await expect(page.locator(`#docs-nav a[href="/docs/${slug}"]`)).toHaveCount(0);
      await expect(page.locator(`a[href="/docs/${slug}"]`)).toHaveCount(0);
      const reply = await page.request.get(`/docs/${slug}`);
      expect(await reply.text()).not.toContain('id="article-body"');
      await page.getByRole('combobox', { name: 'Search the docs' }).fill(title.toLowerCase());
      await expect(page.getByRole('listbox', { name: 'Search results' }).getByRole('option', { name: new RegExp(`^${title}`) })).toHaveCount(0);
    }
    // The overview's cards keep their words, but are no longer links; nor is "privacy" in its text.
    await page.getByRole('combobox', { name: 'Search the docs' }).fill('');
    await expect(page.locator('.doc-card')).toHaveCount(3);
    await expect(page.locator('a.doc-card')).toHaveCount(0);
    await expect(page.locator('.doc-card').nth(1)).toContainText('Paranoid by default');
    await expect(page.locator('#article-body')).toContainText('See privacy for the exact mechanics.');
    await page.goto('/docs/downloads');
    await expect(page.locator('#article-body a')).toHaveCount(0);
    await expect(page.locator('#article-body')).toContainText('in the checks.');
    // The last placeholder left in a published article is gone.
    await page.goto('/docs/supported-documents');
    await expect(page.locator('#article-body')).not.toContainText('[Confirm');
  });

  test('"on this page" follows the reader down the article, and a click scrolls to the heading', async ({ page }) => {
    await page.setViewportSize({ width: 1440, height: 700 });
    await page.goto('/docs/quick-start');
    const toc = page.locator('#toc a');
    await expect(toc).toHaveText(['Open the workspace', 'Drop in the original', 'Drop in the modified version', 'Run the comparison', 'Read the damage', 'Download your copy', 'Files with tracked changes']);
    await expect(toc.nth(0)).toHaveAttribute('aria-current', 'true');
    await toc.nth(2).click();
    await expect(toc.nth(2)).toHaveAttribute('aria-current', 'true');
    await expect(toc.nth(0)).not.toHaveAttribute('aria-current', 'true');
    await expect(page.getByRole('heading', { name: 'Drop in the modified version' })).toBeInViewport();
    await expect(page).toHaveURL(/\/docs\/quick-start$/);
    await page.locator('#reader').evaluate((reader) => reader.scrollTo(0, reader.scrollHeight));
    await expect(toc.nth(6)).toHaveAttribute('aria-current', 'true');
    await page.locator('#reader').evaluate((reader) => reader.scrollTo(0, 0));
    await expect(toc.nth(0)).toHaveAttribute('aria-current', 'true');
    // Narrower than 1180 px there is no such column.
    await page.setViewportSize({ width: 1100, height: 700 });
    await expect(page.locator('.rl-tocwrap')).toBeHidden();
  });

  test('search opens with Ctrl+K (⌘+K on a Mac) and finds articles by title and by text', async ({ page }) => {
    await page.goto('/docs');
    const search = page.getByRole('combobox', { name: 'Search the docs' });
    const hint = (await page.locator('#search-hint').textContent())!;
    expect(hint).toMatch(/^(Ctrl|⌘) K$/);
    await page.keyboard.press(hint.startsWith('⌘') ? 'Meta+K' : 'Control+K');
    await expect(search).toBeFocused();
    // By title.
    await search.fill('downloads');
    const results = page.getByRole('listbox', { name: 'Search results' });
    await expect(results.getByRole('option').first()).toContainText('Downloads');
    // By words that are only in an article's text.
    await search.fill('struck through');
    await expect(results.getByRole('option')).toHaveCount(1);
    await expect(results.getByRole('option')).toContainText('Reading a redline');
    await search.fill('zzzz');
    await expect(results).toContainText('Nothing found in the docs.');
    await search.fill('25 MB');
    await expect(results.getByRole('option').first()).toHaveAttribute('aria-selected', 'true');
    await page.keyboard.press('ArrowDown');
    await expect(results.getByRole('option').nth(1)).toHaveAttribute('aria-selected', 'true');
    const chosen = (await results.getByRole('option').nth(1).locator('.rl-resulttitle').textContent())!;
    await page.keyboard.press('Enter');
    await expect(page.getByRole('heading', { level: 1 })).toHaveText(chosen);
    await expect(results).toBeHidden();
    await expect(search).toHaveValue('');
    // Escape closes the list; a click on a result opens it.
    await search.fill('download');
    await page.keyboard.press('Escape');
    await expect(results).toBeHidden();
    await search.fill('accept or reject');
    await results.getByRole('option').first().click();
    await expect(page).toHaveURL(/\/docs\/downloads$/);
  });

  test('on a phone the menu folds into one button, cards and marks stack, and nothing scrolls sideways', async ({ page }) => {
    await page.setViewportSize({ width: 390, height: 844 });
    await page.goto('/docs');
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
    const toggle = page.getByRole('button', { name: 'docs menu' });
    await expect(toggle).toHaveAttribute('aria-expanded', 'false');
    await expect(page.locator('#docs-nav')).toBeHidden();
    await expect(page.locator('.rl-crumb')).toBeHidden();
    // The three cards sit one above the other.
    const cards = await page.locator('.doc-card').evaluateAll((all) => all.map((card) => card.getBoundingClientRect().left));
    expect(new Set(cards).size).toBe(1);
    await toggle.click();
    await expect(toggle).toHaveAttribute('aria-expanded', 'true');
    await page.locator('#docs-nav').getByRole('link', { name: 'reading a redline' }).click();
    await expect(page).toHaveURL(/\/docs\/reading-a-redline$/);
    await expect(page.locator('#docs-nav')).toBeHidden();
    await expect(page.getByRole('heading', { level: 1 })).toHaveText('Reading a redline');
    expect(await page.evaluate(() => window.scrollY)).toBe(0);
    const row = page.locator('.doc-marks > div').first();
    expect(await row.evaluate((element) => getComputedStyle(element).gridTemplateColumns.split(' ').length)).toBe(1);
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
  });

  test('the bitmap band is drawn and does not move', async ({ page }) => {
    await page.goto('/docs');
    const before = await inkOf(page, 'band');
    expect(before.lit).toBeGreaterThan(5_000);
    await page.waitForTimeout(600);
    expect((await inkOf(page, 'band')).sum).toBe(before.sum);
  });
});

test('the fonts and their licences are served from this site', async ({ page }) => {
  for (const file of ['Geist-Regular.woff2', 'Geist-Medium.woff2', 'GeistMono-Regular.woff2', 'GeistMono-Medium.woff2', 'Newsreader-Regular.woff2', 'Newsreader-Medium.woff2', 'Geist-OFL.txt', 'Newsreader-OFL.txt']) {
    expect((await page.request.get(`/fonts/${file}`)).status(), file).toBe(200);
  }
});
