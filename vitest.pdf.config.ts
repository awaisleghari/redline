import { defineConfig } from 'vitest/config';

// The PDF tests need Docker: they build and run the converter container.
export default defineConfig({
  test: {
    include: ['tests/pdf/**/*.test.ts'],
    environment: 'node',
    globalSetup: ['tests/pdf/container.ts'],
    fileParallelism: false,
    testTimeout: 180_000,
    hookTimeout: 600_000,
  },
});
