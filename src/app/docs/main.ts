/**
 * The docs page. The articles are the files in ./articles/ (one per article);
 * the menu, previous/next, "on this page" and search are all built from them.
 */
import { CHECK_TITLES, CHECKS_FOR, type OutputKind } from '../../checker/report';
import { stillGrid, type GridHandle } from '../grid';

interface Article {
  slug: string;
  section: string;
  label: string;
  title: string;
  intro: string;
  order: number;
  icon: string;
  badge?: string;
  /** Written but not published yet: no menu entry, address, search result or link. */
  hidden: boolean;
  body: string;
  /** Everything a search looks through, in lower case. */
  text: string;
}

function element<T extends HTMLElement>(id: string): T {
  const found = document.getElementById(id);
  if (!found) throw new Error(`missing element #${id}`);
  return found as T;
}

/** Menu icons, drawn at 16 px with a 1.5 px line (the design's own). */
const ICONS: Record<string, string> = {
  lines: '<path d="M2.5 3.5h7M2.5 5.75h7M6 8.25h7.5M6 10.5h7.5M2.5 13h7"></path>',
  bolt: '<path d="M9 1.75 3.5 9h4l-1 5.25L12.5 7h-4z"></path>',
  eye: '<path d="M1.5 8s2.5-4.75 6.5-4.75S14.5 8 14.5 8 12 12.75 8 12.75 1.5 8 1.5 8z"></path><circle cx="8" cy="8" r="2"></circle>',
  download: '<path d="M8 2v8.5M4.25 7 8 10.75 11.75 7M2.5 13.5h11"></path>',
  compass: '<circle cx="8" cy="8" r="6.25"></circle><path d="M10.5 5.5 9 9 5.5 10.5 7 7z"></path>',
  shield: '<path d="M8 1.75 13 3.75v4c0 3-2.25 5.25-5 6.5-2.75-1.25-5-3.5-5-6.5v-4z"></path><path d="M5.75 8 7.4 9.6 10.5 6.5"></path>',
  lock: '<rect x="3" y="7" width="10" height="7" rx="1.75"></rect><path d="M5.25 7V5a2.75 2.75 0 0 1 5.5 0v2"></path>',
  page: '<path d="M4 1.75h5.25L12.5 5v9.25h-8.5z"></path><path d="M9 1.75V5.25h3.5"></path>',
  keyboard: '<rect x="1.75" y="4" width="12.5" height="8.5" rx="1.75"></rect><path d="M4.5 6.75h.01M7 6.75h.01M9.5 6.75h.01M11.5 6.75h.01M5 9.75h6"></path>',
  code: '<path d="M5.5 4.5 2 8l3.5 3.5M10.5 4.5 14 8l-3.5 3.5"></path>',
};
const icon = (name: string): string =>
  `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[name] ?? ICONS.lines}</svg>`;
const SOURCE_URL = 'https://github.com/awaisleghari/redline';
const HOME = '/docs';

// ---------- the articles ----------

const files = import.meta.glob<string>('./articles/*.html', { query: '?raw', import: 'default', eager: true });
const plain = (html: string): string => new DOMParser().parseFromString(html, 'text/html').body.textContent ?? '';

const articles: Article[] = Object.entries(files)
  .map(([path, raw]) => {
    const slug = /([a-z0-9-]+)\.html$/.exec(path)![1]!;
    const header = /^\s*<!--([\s\S]*?)-->/.exec(raw);
    const fields = new Map<string, string>();
    for (const line of (header?.[1] ?? '').split('\n')) {
      const field = /^\s*([a-z]+)\s*:\s*(.*\S)\s*$/.exec(line);
      if (field) fields.set(field[1]!, field[2]!);
    }
    const body = raw.slice(header?.[0].length ?? 0).trim();
    const title = fields.get('title') ?? slug;
    const intro = fields.get('intro') ?? '';
    return {
      slug,
      section: fields.get('section') ?? 'reference',
      label: fields.get('label') ?? title.toLowerCase(),
      title,
      intro,
      order: Number(fields.get('order') ?? 999),
      icon: fields.get('icon') ?? 'lines',
      badge: fields.get('badge'),
      hidden: fields.get('hidden') === 'yes',
      body,
      text: `${title} ${plain(intro)} ${plain(body)}`.replace(/\s+/g, ' ').toLowerCase(),
    };
  })
  .filter((article) => !article.hidden)
  .sort((a, b) => a.order - b.order || a.slug.localeCompare(b.slug));

const addressOf = (article: Article): string => (article === articles[0] ? HOME : `${HOME}/${article.slug}`);
function articleAt(path: string): Article {
  const slug = /^\/docs\/([a-z0-9-]+)\/?$/.exec(path)?.[1];
  return articles.find((article) => article.slug === slug) ?? articles[0]!;
}

// ---------- the menu ----------

const nav = element('docs-nav');
const menuToggle = element<HTMLButtonElement>('menu-toggle');
function buildMenu(): void {
  const sections = [...new Set(articles.map((article) => article.section))];
  nav.replaceChildren(
    ...sections.map((section, index) => {
      const group = document.createElement('div');
      group.className = 'rl-navgroup';
      group.dataset.section = section;
      const title = document.createElement('h2');
      title.className = 'rl-navtitle';
      title.textContent = section;
      group.append(title);
      for (const article of articles.filter((one) => one.section === section)) {
        const link = document.createElement('a');
        link.className = 'rl-btn rl-navitem';
        link.href = addressOf(article);
        link.dataset.slug = article.slug;
        link.innerHTML = icon(article.icon);
        const name = document.createElement('span');
        name.textContent = article.label;
        link.append(name);
        if (article.badge) {
          const badge = document.createElement('span');
          badge.className = 'rl-badge';
          badge.textContent = article.badge;
          link.append(badge);
        }
        group.append(link);
      }
      if (index === sections.length - 1) {
        const source = document.createElement('a');
        source.className = 'rl-btn rl-navitem';
        source.href = SOURCE_URL;
        source.innerHTML = `${icon('code')}<span>source code</span><svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 2.5h5.5V8M9.25 2.75 2.75 9.25"></path></svg>`;
        group.append(source);
      }
      return group;
    }),
  );
}
function setMenuOpen(open: boolean): void {
  nav.classList.toggle('is-open', open);
  menuToggle.setAttribute('aria-expanded', String(open));
}
menuToggle.addEventListener('click', () => setMenuOpen(menuToggle.getAttribute('aria-expanded') !== 'true'));

// ---------- showing an article ----------

const reader = element('reader');
const body = element('article-body');
const toc = element('toc');
let headings: HTMLElement[] = [];

const CHECK_MARK = '<svg width="16" height="16" viewBox="0 0 16 16" role="img" aria-label="Checked"><circle cx="8" cy="8" r="7" fill="rgba(44,194,133,0.14)"></circle><path d="M5 8.2 7 10.2 11 6" fill="none" stroke="#2CC285" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"></path></svg>';

function show(article: Article, toTop: boolean): void {
  const index = articles.indexOf(article);
  document.title = `${article.title} — redline docs`;
  document.querySelector<HTMLElement>('.rl-main')!.dataset.section = article.section;
  element('crumb-section').textContent = article.section;
  element('crumb-title').textContent = article.label;
  element('eyebrow').textContent = article.section;
  element('article-title').textContent = article.title;
  // The articles (intro included) are this site's own files, written by its owner: shown as written.
  element('article-intro').innerHTML = article.intro;
  body.innerHTML = article.body;
  // A link to an article that is not published (hidden, or not written) keeps its words but stops being a link.
  for (const link of Array.from(body.querySelectorAll<HTMLAnchorElement>('a[href^="/docs/"]'))) {
    const slug = /^\/docs\/([a-z0-9-]+)\/?$/.exec(link.getAttribute('href')!)?.[1];
    if (articles.some((one) => one.slug === slug)) continue;
    const plainText = document.createElement(link.classList.contains('doc-card') ? 'div' : 'span');
    plainText.className = link.className;
    plainText.append(...Array.from(link.childNodes));
    link.replaceWith(plainText);
  }
  // The lists of checks come from the app's own list, so the docs cannot fall out of step with it.
  for (const list of Array.from(body.querySelectorAll<HTMLElement>('[data-checks]'))) {
    const kind = list.dataset.checks as OutputKind;
    list.replaceChildren(
      ...(CHECKS_FOR[kind] ?? []).map((id) => {
        const row = document.createElement('div');
        row.innerHTML = CHECK_MARK;
        const text = document.createElement('span');
        text.textContent = CHECK_TITLES[id];
        row.append(text);
        return row;
      }),
    );
  }

  for (const link of Array.from(nav.querySelectorAll<HTMLAnchorElement>('[data-slug]'))) {
    if (link.dataset.slug === article.slug) link.setAttribute('aria-current', 'page');
    else link.removeAttribute('aria-current');
  }

  const previous = articles[index - 1];
  const next = articles[index + 1];
  const previousLink = element<HTMLAnchorElement>('prev-article');
  const nextLink = element<HTMLAnchorElement>('next-article');
  previousLink.hidden = !previous;
  element('prev-gap').hidden = !!previous;
  nextLink.hidden = !next;
  if (previous) {
    previousLink.href = addressOf(previous);
    element('prev-name').textContent = previous.label;
  }
  if (next) {
    nextLink.href = addressOf(next);
    element('next-name').textContent = next.label;
  }

  // "On this page": one link per heading, in the article's own order.
  headings = Array.from(body.querySelectorAll<HTMLElement>('h2, [data-toc]')).filter((heading) => heading.id);
  toc.replaceChildren(
    ...headings.map((heading) => {
      const link = document.createElement('a');
      link.className = 'rl-toc';
      link.href = `#${heading.id}`;
      link.textContent = heading.dataset.toc ?? heading.textContent ?? '';
      return link;
    }),
  );
  if (toTop) {
    reader.scrollTop = 0;
    if (reader.scrollHeight <= reader.clientHeight + 1) window.scrollTo(0, 0);
  }
  markHeading();
}

/** The heading the reader has reached: the last one at or above a line a little below the top of the article view. */
function markHeading(): void {
  const whole = reader.scrollHeight <= reader.clientHeight + 1;
  const line = (whole ? 0 : reader.getBoundingClientRect().top) + 120;
  let active = headings[0];
  for (const heading of headings) if (heading.getBoundingClientRect().top <= line) active = heading;
  // At the very end of the article the last heading may never reach the line.
  const atEnd = whole ? window.innerHeight + window.scrollY >= document.documentElement.scrollHeight - 2 : reader.scrollTop + reader.clientHeight >= reader.scrollHeight - 2;
  if (atEnd && headings.length && (whole ? window.scrollY : reader.scrollTop) > 0) active = headings[headings.length - 1];
  for (const link of Array.from(toc.querySelectorAll<HTMLAnchorElement>('a'))) {
    if (active && link.hash === `#${active.id}`) link.setAttribute('aria-current', 'true');
    else link.removeAttribute('aria-current');
  }
}
reader.addEventListener('scroll', markHeading, { passive: true });
window.addEventListener('scroll', markHeading, { passive: true });
window.addEventListener('resize', markHeading);

// ---------- addresses: each article has its own, and the back button works ----------

function go(address: string, replace = false): void {
  const url = new URL(address, location.href);
  const article = articleAt(url.pathname);
  if (replace) history.replaceState(null, '', url.pathname);
  else if (url.pathname !== location.pathname) history.pushState(null, '', url.pathname);
  show(article, true);
  setMenuOpen(false);
  closeSearch();
}
window.addEventListener('popstate', () => show(articleAt(location.pathname), true));

document.addEventListener('click', (event) => {
  if (event.defaultPrevented || event.button !== 0 || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) return;
  const link = (event.target as Element).closest?.('a');
  if (!link) return;
  const href = link.getAttribute('href') ?? '';
  if (href.startsWith('#')) {
    // A link to a heading on this page: scroll to it, without changing the address.
    event.preventDefault();
    const smooth = !window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    document.getElementById(href.slice(1))?.scrollIntoView({ behavior: smooth ? 'smooth' : 'auto', block: 'start' });
    return;
  }
  const url = new URL(link.href, location.href);
  if (url.origin === location.origin && (url.pathname === HOME || url.pathname.startsWith(`${HOME}/`))) {
    event.preventDefault();
    go(url.pathname);
  }
});

// ---------- search ----------

const search = element<HTMLInputElement>('search');
const results = element('search-results');
let picked = -1;

function closeSearch(): void {
  results.hidden = true;
  results.replaceChildren();
  search.setAttribute('aria-expanded', 'false');
  search.removeAttribute('aria-activedescendant');
  picked = -1;
}

function runSearch(): void {
  const words = search.value.toLowerCase().split(/\s+/).filter(Boolean);
  if (!words.length) return closeSearch();
  const found = articles
    .map((article) => {
      if (!words.every((word) => article.text.includes(word))) return undefined;
      const inTitle = words.filter((word) => article.title.toLowerCase().includes(word) || article.label.includes(word)).length;
      return { article, score: inTitle * 10 + words.reduce((sum, word) => sum + Math.min(5, article.text.split(word).length - 1), 0) };
    })
    .filter((hit) => !!hit)
    .sort((a, b) => b.score - a.score);
  results.replaceChildren();
  if (!found.length) {
    const none = document.createElement('p');
    none.className = 'rl-noresult';
    none.textContent = 'Nothing found in the docs.';
    results.append(none);
  }
  found.forEach(({ article }, index) => {
    const link = document.createElement('a');
    link.className = 'rl-result';
    link.id = `search-result-${index}`;
    link.setAttribute('role', 'option');
    link.href = addressOf(article);
    const title = document.createElement('span');
    title.className = 'rl-resulttitle';
    title.textContent = article.title;
    const text = document.createElement('span');
    text.className = 'rl-resulttext';
    // A line of the article around the first word found.
    const whole = `${plain(article.intro)} ${plain(article.body)}`.replace(/\s+/g, ' ').trim();
    const at = Math.max(0, whole.toLowerCase().indexOf(words[0]!));
    const from = Math.max(0, whole.lastIndexOf(' ', Math.max(0, at - 30)));
    text.textContent = `${from > 0 ? '…' : ''}${whole.slice(from, from + 90).trim()}${from + 90 < whole.length ? '…' : ''}`;
    link.append(title, text);
    results.append(link);
  });
  picked = found.length ? 0 : -1;
  results.hidden = false;
  search.setAttribute('aria-expanded', 'true');
  markPicked();
}
function markPicked(): void {
  const options = Array.from(results.querySelectorAll<HTMLElement>('.rl-result'));
  options.forEach((option, index) => option.setAttribute('aria-selected', String(index === picked)));
  const option = options[picked];
  if (option) {
    search.setAttribute('aria-activedescendant', option.id);
    option.scrollIntoView({ block: 'nearest' });
  }
}
search.addEventListener('input', runSearch);
search.addEventListener('focus', runSearch);
search.addEventListener('keydown', (event) => {
  const options = Array.from(results.querySelectorAll<HTMLAnchorElement>('.rl-result'));
  if (event.key === 'Escape') {
    if (results.hidden) search.blur();
    closeSearch();
  } else if (event.key === 'ArrowDown' && options.length) picked = (picked + 1) % options.length;
  else if (event.key === 'ArrowUp' && options.length) picked = (picked - 1 + options.length) % options.length;
  else if (event.key === 'Enter' && options[picked]) {
    search.value = '';
    go(options[picked]!.href);
    reader.focus();
  } else return;
  event.preventDefault();
  markPicked();
});
document.addEventListener('click', (event) => {
  if (!(event.target as Element).closest?.('.rl-searchwrap')) closeSearch();
});
results.addEventListener('click', () => (search.value = ''));

// ⌘+K on a Mac, Ctrl+K elsewhere, opens the search.
const onMac = /Mac|iPhone|iPad/.test(navigator.platform);
element('search-hint').textContent = onMac ? '⌘ K' : 'Ctrl K';
document.addEventListener('keydown', (event) => {
  if (event.key.toLowerCase() === 'k' && (onMac ? event.metaKey : event.ctrlKey) && !event.altKey && !event.shiftKey) {
    event.preventDefault();
    search.focus();
    search.select();
  }
});

// ---------- the bitmap band: still on this page ----------

let band: GridHandle | undefined = stillGrid(element<HTMLCanvasElement>('band'));
window.addEventListener('pagehide', () => {
  band?.stop();
  band = undefined;
});
window.addEventListener('pageshow', () => {
  band ??= stillGrid(element<HTMLCanvasElement>('band'));
});

buildMenu();
show(articleAt(location.pathname), false);
