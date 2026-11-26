/**
 * The deployment settings (wrangler.jsonc) carry promises made to the user:
 * nothing is logged, the converter cannot multiply without limit, and no secret
 * lives in the repository. These tests hold the file to them.
 */
import { existsSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const read = (path: string): string => readFileSync(join(root, path), 'utf8');

/** Remove // comments (outside strings) so the file parses as JSON. */
function parseJsonc(text: string): Record<string, any> {
  let out = '';
  let inString = false;
  for (let i = 0; i < text.length; i++) {
    const char = text[i]!;
    if (inString) {
      out += char;
      if (char === '\\') out += text[++i];
      else if (char === '"') inString = false;
    } else if (char === '"') {
      inString = true;
      out += char;
    } else if (char === '/' && text[i + 1] === '/') {
      while (i < text.length && text[i] !== '\n') i++;
      out += '\n';
    } else out += char;
  }
  return JSON.parse(out);
}

const config = parseJsonc(read('wrangler.jsonc'));
const worker = read('worker/index.ts');

describe('deployment settings', () => {
  it('switch off every Cloudflare log, for the Worker and for the converter', () => {
    expect(config.observability.enabled).toBe(false);
    expect(config.observability.logs).toEqual({ enabled: false, invocation_logs: false });
    expect(config.observability.traces.enabled).toBe(false);
    expect(config.containers[0].observability.logs.enabled).toBe(false);
    expect(config.logpush ?? false).toBe(false);
    expect(config.tail_consumers ?? []).toEqual([]);
  });

  it('the Worker code itself writes no log line', () => {
    for (const file of ['worker/index.ts', 'worker/pdf-route.ts', 'worker/state-route.ts']) expect(read(file), file).not.toMatch(/console\s*\./);
  });

  it('cap the converter at the number of copies the Worker uses, with no internet access', () => {
    const instances = Number(/const INSTANCES = (\d+);/.exec(worker)![1]);
    expect(config.containers).toHaveLength(1);
    expect(config.containers[0].max_instances).toBe(instances);
    expect(config.containers[0].class_name).toBe('PdfConverter');
    expect(worker).toMatch(/enableInternet = false;/);
    expect(existsSync(join(root, config.containers[0].image))).toBe(true);
    expect(config.durable_objects.bindings).toEqual([{ name: 'PDF_CONVERTER', class_name: 'PdfConverter' }]);
  });

  it('serve the built site and nothing else, at one address, with no stored data and no secrets', () => {
    expect(config.assets).toEqual({ directory: './dist', binding: 'ASSETS', not_found_handling: 'none' });
    expect(config.preview_urls).toBe(false);
    // The two public addresses, and no others: no workers.dev address, no wildcard routes.
    expect(config.workers_dev).toBe(false);
    expect(config.routes).toEqual([
      { pattern: 'useredline.dev', custom_domain: true },
      { pattern: 'www.useredline.dev', custom_domain: true },
    ]);
    // No storage of any kind is attached, and no variable or secret is declared.
    for (const key of ['vars', 'secrets', 'kv_namespaces', 'r2_buckets', 'd1_databases', 'queues', 'analytics_engine_datasets', 'hyperdrive', 'vectorize', 'services', 'ai']) expect(config[key], key).toBeUndefined();
    expect(Object.keys(config).sort()).toEqual(['$schema', 'assets', 'compatibility_date', 'containers', 'durable_objects', 'main', 'migrations', 'name', 'observability', 'preview_urls', 'routes', 'workers_dev'].sort());
  });

  it('the site-wide headers file is part of the build', () => {
    const headers = read('public/_headers');
    expect(headers).toContain('X-Content-Type-Options: nosniff');
    expect(headers).toContain('X-Frame-Options: DENY');
    expect(headers).toContain('Referrer-Policy: no-referrer');
  });
});
