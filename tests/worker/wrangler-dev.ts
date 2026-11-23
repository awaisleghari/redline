/**
 * Starts the real Worker locally (`wrangler dev`: Cloudflare's own runtime, the
 * settings in wrangler.jsonc, and the converter container run through Docker)
 * for the Worker tests, and stops it afterwards.
 */
import { execFileSync, spawn, spawnSync, type ChildProcess } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const WORKER_PORT = 8788;
export const WORKER_URL = `http://127.0.0.1:${WORKER_PORT}`;

const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
let wrangler: ChildProcess | undefined;
let output = '';

/** The converter containers wrangler has started (not its network helper containers). */
export const converterContainers = (): string[] =>
  execFileSync('docker', ['ps', '--format', '{{.Names}}'], { encoding: 'utf8' })
    .split('\n')
    .filter((name) => name.includes('-PdfConverter-') && !name.endsWith('-proxy'));

const removeContainers = (): void => {
  const all = spawnSync('docker', ['ps', '-aq', '--filter', 'name=-PdfConverter-'], { encoding: 'utf8' }).stdout.split('\n').filter(Boolean);
  if (all.length > 0) spawnSync('docker', ['rm', '-f', ...all]);
};

export async function setup(): Promise<void> {
  // The Worker serves the built site, so build it first.
  execFileSync('npm', ['run', 'build'], { cwd: root, stdio: 'ignore' });
  removeContainers();
  // The first start builds the converter image, which can take several minutes.
  const deadline = Date.now() + 840_000;
  for (let attempt = 1; ; attempt++) {
    output = '';
    wrangler = spawn('npx', ['wrangler', 'dev', '--port', String(WORKER_PORT), '--ip', '127.0.0.1'], { cwd: root, detached: true, stdio: ['ignore', 'pipe', 'pipe'] });
    wrangler.stdout!.on('data', (chunk: Buffer) => (output += chunk.toString()));
    wrangler.stderr!.on('data', (chunk: Buffer) => (output += chunk.toString()));
    while (Date.now() < deadline && wrangler.exitCode === null) {
      try {
        if ((await fetch(`${WORKER_URL}/api/converter-state`)).ok) return;
      } catch {
        // not up yet
      }
      await new Promise((resolve) => setTimeout(resolve, 500));
    }
    // Docker Hub sometimes refuses the base-image lookup for a moment; try again before giving up.
    if (wrangler.exitCode === null || attempt === 5 || !/429|Too Many Requests/.test(output)) break;
    await new Promise((resolve) => setTimeout(resolve, 30_000));
  }
  teardown();
  throw new Error(`wrangler dev did not start:\n${output.slice(-2000)}`);
}

export function teardown(): void {
  if (wrangler?.pid) {
    try {
      process.kill(-wrangler.pid, 'SIGTERM');
    } catch {
      // already gone
    }
  }
  removeContainers();
}
