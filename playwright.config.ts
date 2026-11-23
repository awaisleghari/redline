import { defineConfig, devices } from '@playwright/test';

// By default the browser tests run against the production build served by
// `vite preview`, so they exercise the same bundled files a real deployment
// serves. With REDLINE_URL set they run against that address instead: the
// Worker under `wrangler dev`, or the live site.
//   REDLINE_URL=https://redline.<account>.workers.dev npm run test:e2e
// (Docker is still needed: a few tests make deliberately wrong PDFs locally.)
const target = process.env.REDLINE_URL?.replace(/\/+$/, '');

export default defineConfig({
  testDir: 'tests/e2e',
  globalSetup: './tests/e2e/global-setup.ts',
  // A converter on Cloudflare that has gone to sleep takes a while to start again.
  timeout: target ? 240_000 : 120_000,
  fullyParallel: true,
  // Against a deployed site, a few at a time: the converter handles two documents at once per copy.
  workers: target ? 2 : undefined,
  forbidOnly: !!process.env.CI,
  retries: 0,
  reporter: [['list']],
  use: {
    baseURL: target ?? 'http://localhost:4173',
    trace: 'off',
  },
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'] } }],
  webServer: target
    ? undefined
    : {
        command: 'npm run build && npm run preview',
        url: 'http://localhost:4173',
        reuseExistingServer: !process.env.CI,
        timeout: 180_000,
      },
});
