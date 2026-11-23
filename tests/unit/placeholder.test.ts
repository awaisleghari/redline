import { describe, expect, it } from 'vitest';
import { APP_NAME } from '../../src/main';

describe('project setup', () => {
  it('runs a unit test', () => {
    expect(APP_NAME).toBe('Redline');
  });
});
