/**
 * Starts the PDF converter container for the PDF tests and stops it afterwards.
 * Needs Docker. The image is (re)built from pdf-container/ first; Docker's
 * cache makes that quick when nothing has changed.
 */
import { execFileSync, spawnSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const CONTAINER_NAME = 'redline-pdf-test';
export const CONTAINER_PORT = 8791;
export const CONVERT_URL = `http://127.0.0.1:${CONTAINER_PORT}/convert`;
export const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');

function docker(args: string[], allowFailure = false): string {
  const run = spawnSync('docker', args, { encoding: 'utf8' });
  if (run.error) throw new Error(`Docker is needed for the PDF tests and could not be run (${run.error.message}).`);
  if (run.status !== 0 && !allowFailure) throw new Error(`docker ${args[0]} failed: ${run.stderr.trim().split('\n').pop()}`);
  return run.stdout;
}

export async function startContainer(): Promise<void> {
  docker(['build', '-q', '-t', 'redline-pdf', join(root, 'pdf-container')]);
  docker(['rm', '-f', CONTAINER_NAME], true);
  // Read-only file system with only memory-backed scratch space, and no route out to the internet from inside.
  docker(['run', '-d', '--name', CONTAINER_NAME, '--read-only', '--tmpfs', '/tmp', '--tmpfs', '/home/converter', '-p', `127.0.0.1:${CONTAINER_PORT}:8080`, 'redline-pdf']);
  for (let attempt = 0; attempt < 60; attempt++) {
    try {
      if ((await fetch(`http://127.0.0.1:${CONTAINER_PORT}/health`)).ok) return;
    } catch {
      // not up yet
    }
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
  throw new Error('The PDF converter container did not start.');
}

export function stopContainer(): void {
  docker(['rm', '-f', CONTAINER_NAME], true);
}

/** Everything the container has written to its log so far. */
export const containerLog = (name = CONTAINER_NAME): string => {
  const run = spawnSync('docker', ['logs', name], { encoding: 'utf8' });
  return run.stdout + run.stderr;
};

/** Every file and folder the container currently holds in its scratch space. */
export const containerScratchFiles = (name = CONTAINER_NAME): string[] =>
  // As root, so that nothing is hidden from the search by permissions.
  execFileSync('docker', ['exec', '-u', '0', name, 'sh', '-c', 'find /tmp /home/converter /var/tmp -mindepth 1'], { encoding: 'utf8' })
    .split('\n')
    .filter(Boolean);

/** Processes of LibreOffice still running in the container. */
export const containerOfficeProcesses = (name = CONTAINER_NAME): string[] =>
  execFileSync('docker', ['exec', '-u', '0', name, 'sh', '-c', 'for p in /proc/[0-9]*; do tr "\\0" " " < $p/cmdline 2>/dev/null; echo; done | grep "[s]office" || true'], { encoding: 'utf8' })
    .split('\n')
    .filter(Boolean);

/** Run a short Python program inside the container (to reach the server directly, without Docker's port forwarding in between). */
export const runInContainer = (python: string, name = CONTAINER_NAME): string => execFileSync('docker', ['exec', name, 'python3', '-c', python], { encoding: 'utf8' });

/** Start a second converter with its own settings (for example a very short time limit). */
export function startExtraContainer(name: string, port: number, env: Record<string, string>): void {
  docker(['rm', '-f', name], true);
  docker(['run', '-d', '--name', name, '--read-only', '--tmpfs', '/tmp', '--tmpfs', '/home/converter', ...Object.entries(env).flatMap(([k, v]) => ['-e', `${k}=${v}`]), '-p', `127.0.0.1:${port}:8080`, 'redline-pdf']);
}
export const removeContainer = (name: string): void => void docker(['rm', '-f', name], true);

export async function convertToPdf(docx: Uint8Array): Promise<Uint8Array> {
  const response = await fetch(CONVERT_URL, { method: 'POST', headers: { 'Content-Type': DOCX_TYPE }, body: new Blob([docx.slice().buffer]) });
  if (!response.ok) throw new Error(`conversion failed with status ${response.status}`);
  return new Uint8Array(await response.arrayBuffer());
}

// Vitest global setup and teardown.
export async function setup(): Promise<void> {
  await startContainer();
}
export function teardown(): void {
  stopContainer();
}
