/**
 * The brand bitmap: 7 px squares on a 10 px grid in the three brand colours.
 *
 * Ported as is from the design handoff (reference/index.html, app.html and
 * docs.html). Every number here is the design's; do not "simplify" the timing
 * or the falloff maths.
 */

export interface GridHandle {
  stop: () => void;
}

type Lit = { c: number; r: number; k: number; t0: number; d: number; p: number };

const COLORS: [number, number, number][] = [[255, 90, 92], [44, 194, 133], [85, 133, 255]];
const PITCH = 10;
const CELL = 7; /* 7 px squares on a 10 px grid */
const rgba = (k: number, a: number, digits: number): string => {
  const [R, G, B] = COLORS[k]!;
  return 'rgba(' + R + ',' + G + ',' + B + ',' + a.toFixed(digits) + ')';
};
const reducedMotion = (): boolean => !!window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
const env = (x: number): number => (x < 0.22 ? 1 - Math.pow(1 - x / 0.22, 3) : Math.pow(1 - (x - 0.22) / 0.78, 2));

/** The landing page's animated grid. `invert` puts the bright edge at the bottom. */
export function startGrid(cv: HTMLCanvasElement, invert: boolean): GridHandle {
  const h: { raf: number | null; ro: ResizeObserver | null; stop: () => void } = {
    raf: null,
    ro: null,
    stop: () => {
      if (h.raf) cancelAnimationFrame(h.raf);
      h.raf = null;
      if (h.ro) {
        h.ro.disconnect();
        h.ro = null;
      }
    },
  };
  const ctx = cv.getContext('2d')!;
  const base = document.createElement('canvas');
  let W = 0, H = 0, cols = 0, rows = 0, dpr = 1, ox = 0, lit: Lit[] = [], last = 0;
  const reduce = reducedMotion();

  const fall = (r: number): number => Math.pow(Math.max(0, 1 - r / rows), 1.7); /* strong at row 0, gone at the far edge */
  const yOf = (r: number): number => (invert ? H - CELL - r * PITCH : r * PITCH); /* row 0 sits at the top, or at the bottom when inverted */

  const resize = (): void => {
    const rect = cv.getBoundingClientRect();
    dpr = Math.min(window.devicePixelRatio || 1, 2);
    W = Math.max(1, Math.round(rect.width));
    H = Math.max(1, Math.round(rect.height));
    cv.width = W * dpr;
    cv.height = H * dpr;
    cols = Math.ceil(W / PITCH) + 1;
    rows = Math.ceil(H / PITCH);
    ox = Math.round((W - (cols - 1) * PITCH - CELL) / 2);
    base.width = cv.width;
    base.height = cv.height;
    const b = base.getContext('2d')!;
    b.setTransform(dpr, 0, 0, dpr, 0, 0);
    b.clearRect(0, 0, W, H);
    /* resting field: every cell carries a faint tint of one of the three colours */
    for (let r = 0; r < rows; r++) {
      const f = fall(r);
      for (let c = 0; c < cols; c++) {
        const a = (0.07 + Math.random() * 0.13) * f + 0.018;
        b.fillStyle = rgba(Math.floor(Math.random() * 3), a, 4);
        b.fillRect(ox + c * PITCH, yOf(r), CELL, CELL);
      }
    }
    lit = lit.filter((p) => p.c < cols && p.r < rows);
  };

  const pickRow = (): number => Math.floor(rows * Math.pow(Math.random(), 1.9));
  const spawn = (now: number): void => {
    const r = pickRow(), c = Math.floor(Math.random() * cols);
    lit.push({ c, r, k: Math.floor(Math.random() * 3), t0: now, d: 1400 + Math.random() * 2800, p: (0.6 + Math.random() * 0.4) * fall(r) });
  };
  /* now and then, a double rule sweeps across, echoing the mark */
  const sweep = (now: number): void => {
    const r = Math.floor(rows * 0.55 * Math.pow(Math.random(), 1.4));
    const len = 4 + Math.floor(Math.random() * 12);
    const fromRight = Math.random() < 0.5;
    const c0 = Math.floor(Math.random() * Math.max(1, cols - len));
    const k = Math.floor(Math.random() * 3);
    const p = (0.75 + Math.random() * 0.25) * fall(r);
    for (let i = 0; i < len; i++) {
      const c = c0 + i, delay = (fromRight ? len - 1 - i : i) * 45;
      for (const rr of [r, r + 2]) if (rr < rows) lit.push({ c, r: rr, k, t0: now + delay, d: 1500, p });
    }
  };

  const draw = (now: number): void => {
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, W, H);
    ctx.drawImage(base, 0, 0, W, H);
    ctx.globalCompositeOperation = 'lighter';
    const keep: Lit[] = [];
    for (const q of lit) {
      const x = (now - q.t0) / q.d;
      if (x >= 1) continue;
      keep.push(q);
      if (x < 0) continue;
      const a = q.p * env(x);
      if (a < 0.01) continue;
      const px = ox + q.c * PITCH, py = yOf(q.r);
      ctx.fillStyle = rgba(q.k, a * 0.16, 3);
      ctx.fillRect(px - 3, py - 3, CELL + 6, CELL + 6); /* halo */
      ctx.fillStyle = rgba(q.k, a, 3);
      ctx.fillRect(px, py, CELL, CELL);
    }
    ctx.globalCompositeOperation = 'source-over';
    lit = keep;
  };

  const tick = (now: number): void => {
    const dt = Math.min(64, last ? now - last : 16);
    last = now;
    const n = cols * rows * 0.00042 * dt; /* ignition rate scales with area */
    let k = Math.floor(n);
    if (Math.random() < n - k) k++;
    for (let i = 0; i < k; i++) spawn(now);
    if (Math.random() < dt / 650) sweep(now);
    draw(now);
    h.raf = requestAnimationFrame(tick);
  };

  const still = (): void => {
    const now = performance.now();
    lit = [];
    for (let i = 0; i < cols * rows * 0.22; i++) spawn(now - 400);
    draw(now);
  };

  resize();
  if (window.ResizeObserver) {
    h.ro = new ResizeObserver(() => {
      resize();
      if (reduce) still();
    });
    h.ro.observe(cv);
  }
  if (reduce) still();
  else h.raf = requestAnimationFrame(tick);
  return h;
}

/** The workspace's band: still and faint behind the viewer; animated while a redline is being created. */
export function bandGrid(cv: HTMLCanvasElement, animate: boolean): GridHandle {
  const h: { raf: number | null; ro: ResizeObserver | null; stop: () => void } = {
    raf: null,
    ro: null,
    stop: () => {
      if (h.raf) cancelAnimationFrame(h.raf);
      h.raf = null;
      if (h.ro) {
        h.ro.disconnect();
        h.ro = null;
      }
    },
  };
  const live = animate && !reducedMotion();
  const ctx = cv.getContext('2d')!;
  const base = document.createElement('canvas');
  let W = 0, H = 0, cols = 0, rows = 0, dpr = 1, ox = 0, lit: Lit[] = [], last = 0;
  const fall = (r: number): number => Math.pow(Math.max(0, 1 - r / rows), 1.5);

  function draw(now: number): void {
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, W, H);
    ctx.drawImage(base, 0, 0, W, H);
    if (!lit.length) return;
    ctx.globalCompositeOperation = 'lighter';
    const keep: Lit[] = [];
    for (const q of lit) {
      const x = (now - q.t0) / q.d;
      if (x >= 1) continue;
      keep.push(q);
      if (x < 0) continue;
      const a = q.p * env(x);
      if (a < 0.01) continue;
      const px = ox + q.c * PITCH, py = q.r * PITCH;
      ctx.fillStyle = rgba(q.k, a * 0.16, 3);
      ctx.fillRect(px - 3, py - 3, CELL + 6, CELL + 6);
      ctx.fillStyle = rgba(q.k, a, 3);
      ctx.fillRect(px, py, CELL, CELL);
    }
    ctx.globalCompositeOperation = 'source-over';
    lit = keep;
  }
  const resize = (): void => {
    const rect = cv.getBoundingClientRect();
    dpr = Math.min(window.devicePixelRatio || 1, 2);
    W = Math.max(1, Math.round(rect.width));
    H = Math.max(1, Math.round(rect.height));
    cv.width = W * dpr;
    cv.height = H * dpr;
    cols = Math.ceil(W / PITCH) + 1;
    rows = Math.ceil(H / PITCH);
    ox = Math.round((W - (cols - 1) * PITCH - CELL) / 2);
    base.width = cv.width;
    base.height = cv.height;
    const b = base.getContext('2d')!;
    b.setTransform(dpr, 0, 0, dpr, 0, 0);
    b.clearRect(0, 0, W, H);
    const k = animate ? 1.25 : 0.7;
    for (let r = 0; r < rows; r++) {
      const f = fall(r);
      for (let c = 0; c < cols; c++) {
        const a = ((0.05 + Math.random() * 0.12) * f + 0.012) * k;
        b.fillStyle = rgba(Math.floor(Math.random() * 3), a, 4);
        b.fillRect(ox + c * PITCH, r * PITCH, CELL, CELL);
      }
    }
    draw(performance.now());
  };
  const spawn = (now: number): void => {
    const r = Math.floor(rows * Math.pow(Math.random(), 1.6)), c = Math.floor(Math.random() * cols);
    lit.push({ c, r, k: Math.floor(Math.random() * 3), t0: now, d: 900 + Math.random() * 1800, p: (0.6 + Math.random() * 0.4) * fall(r) });
  };
  const sweep = (now: number): void => {
    const r = Math.floor(rows * 0.6 * Math.random()), len = 5 + Math.floor(Math.random() * 14);
    const c0 = Math.floor(Math.random() * Math.max(1, cols - len)), k = Math.floor(Math.random() * 3), back = Math.random() < 0.5;
    const p = (0.8 + Math.random() * 0.2) * fall(r);
    for (let i = 0; i < len; i++) for (const rr of [r, r + 2]) if (rr < rows) lit.push({ c: c0 + i, r: rr, k, t0: now + (back ? len - 1 - i : i) * 40, d: 1300, p });
  };
  const tick = (now: number): void => {
    const dt = Math.min(64, last ? now - last : 16);
    last = now;
    const n = cols * rows * 0.0005 * dt;
    let k = Math.floor(n);
    if (Math.random() < n - k) k++;
    for (let i = 0; i < k; i++) spawn(now);
    if (Math.random() < dt / 420) sweep(now);
    draw(now);
    h.raf = requestAnimationFrame(tick);
  };
  resize();
  if (window.ResizeObserver) {
    h.ro = new ResizeObserver(resize);
    h.ro.observe(cv);
  }
  if (live) h.raf = requestAnimationFrame(tick);
  return h;
}

/** The docs page's band: always still, with a few brighter cells near the top. */
export function stillGrid(cv: HTMLCanvasElement): GridHandle {
  const h: { ro: ResizeObserver | null; stop: () => void } = {
    ro: null,
    stop: () => {
      if (h.ro) {
        h.ro.disconnect();
        h.ro = null;
      }
    },
  };
  const draw = (): void => {
    const rect = cv.getBoundingClientRect(), dpr = Math.min(window.devicePixelRatio || 1, 2);
    const W = Math.max(1, Math.round(rect.width)), H = Math.max(1, Math.round(rect.height));
    cv.width = W * dpr;
    cv.height = H * dpr;
    const ctx = cv.getContext('2d')!;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, W, H);
    const cols = Math.ceil(W / PITCH) + 1, rows = Math.ceil(H / PITCH), ox = Math.round((W - (cols - 1) * PITCH - CELL) / 2);
    for (let r = 0; r < rows; r++) {
      const f = Math.pow(Math.max(0, 1 - r / rows), 1.5);
      for (let c = 0; c < cols; c++) {
        const k = Math.floor(Math.random() * 3);
        const lit = Math.random() < 0.012 * f;
        const a = lit ? 0.35 + Math.random() * 0.3 : ((0.05 + Math.random() * 0.12) * f + 0.012) * 0.7;
        ctx.fillStyle = rgba(k, a, 4);
        ctx.fillRect(ox + c * PITCH, r * PITCH, CELL, CELL);
      }
    }
  };
  draw();
  if (window.ResizeObserver) {
    h.ro = new ResizeObserver(draw);
    h.ro.observe(cv);
  }
  return h;
}
