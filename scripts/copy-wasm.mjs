// Copies the compare engine's WebAssembly runtime out of the docxodus package
// into public/wasm/, where Vite serves it. The copy is git-ignored: the files
// are large and are always reproducible from node_modules.
import { cpSync, existsSync, mkdirSync, rmSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const source = join(root, 'node_modules', 'docxodus', 'dist', 'wasm', '_framework');
const target = join(root, 'public', 'wasm', '_framework');

if (!existsSync(source)) {
  console.error(`copy-wasm: ${source} not found. Run "npm install" first.`);
  process.exit(1);
}
rmSync(target, { recursive: true, force: true });
mkdirSync(target, { recursive: true });
// Skip the pre-compressed .br/.gz twins: the web host compresses on the fly.
cpSync(source, target, { recursive: true, filter: (path) => !/\.(br|gz)$/.test(path) });
console.log('copy-wasm: engine runtime copied to public/wasm/_framework');
