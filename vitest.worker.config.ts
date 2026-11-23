import { defineConfig } from 'vitest/config';

// The Worker tests run the real Cloudflare Worker (worker/index.ts, wrangler.jsonc)
// with the real converter container, locally under `wrangler dev`. They need Docker.
export default defineConfig({
  test: {
    include: ['tests/worker/**/*.test.ts'],
    environment: 'node',
    globalSetup: ['tests/worker/wrangler-dev.ts'],
    fileParallelism: false,
    testTimeout: 180_000,
    hookTimeout: 900_000,
  },
});
