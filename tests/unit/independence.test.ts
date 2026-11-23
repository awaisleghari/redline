/**
 * The Checker must be independent of the Producer (PLAN.md, "Independence
 * rules"). This test reads the source files and fails if:
 *   - anything in src/checker/ imports from src/producer/ (or the app);
 *   - anything in src/common/ imports from the Producer, the Checker or the app;
 *   - the Checker or the shared helpers use the compare engine (docxodus), or
 *     any package outside a short allow-list.
 */
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const SRC = join(ROOT, 'src');

function sourceFiles(folder: string): string[] {
  const found: string[] = [];
  for (const name of readdirSync(folder)) {
    const path = join(folder, name);
    if (statSync(path).isDirectory()) found.push(...sourceFiles(path));
    else if (/\.(ts|tsx|js|mjs)$/.test(name)) found.push(path);
  }
  return found;
}

/** Every module a source file pulls in: static and dynamic imports, re-exports, require, and worker/asset URLs. */
export function importsOf(source: string): string[] {
  const withoutComments = source.replace(/\/\*[\s\S]*?\*\//g, '').replace(/(^|[^:'"`])\/\/.*$/gm, '$1');
  const patterns = [
    /\bimport\s+(?:type\s+)?[^'"`;]*?\bfrom\s*['"`]([^'"`]+)['"`]/g,
    /\bexport\s+(?:type\s+)?[^'"`;]*?\bfrom\s*['"`]([^'"`]+)['"`]/g,
    /\bimport\s*['"`]([^'"`]+)['"`]/g,
    /\bimport\s*\(\s*['"`]([^'"`]+)['"`]/g,
    /\brequire\s*\(\s*['"`]([^'"`]+)['"`]/g,
    /\bnew\s+URL\s*\(\s*['"`]([^'"`]+)['"`]\s*,\s*import\.meta\.url/g,
  ];
  const found = new Set<string>();
  for (const pattern of patterns) for (const match of withoutComments.matchAll(pattern)) found.add(match[1]!);
  return [...found];
}

interface Violation {
  file: string;
  imports: string;
  why: string;
}

function violations(files: { path: string; source: string }[], rules: { forbiddenFolders: string[]; allowedPackages: string[] }): Violation[] {
  const result: Violation[] = [];
  for (const file of files) {
    for (const specifier of importsOf(file.source)) {
      if (specifier.startsWith('.') || specifier.startsWith('/')) {
        const target = specifier.startsWith('/') ? join(ROOT, specifier) : resolve(dirname(file.path), specifier);
        for (const folder of rules.forbiddenFolders) {
          const inside = relative(join(SRC, folder), target);
          if (!inside.startsWith('..')) result.push({ file: relative(ROOT, file.path), imports: specifier, why: `reaches into src/${folder}/` });
        }
        if (relative(SRC, target).startsWith('..')) result.push({ file: relative(ROOT, file.path), imports: specifier, why: 'reaches outside src/' });
      } else {
        const packageName = specifier.startsWith('@') ? specifier.split('/').slice(0, 2).join('/') : specifier.split('/')[0]!;
        if (!rules.allowedPackages.includes(packageName)) result.push({ file: relative(ROOT, file.path), imports: specifier, why: `package "${packageName}" is not on the allow-list` });
      }
    }
  }
  return result;
}

const load = (folder: string): { path: string; source: string }[] => sourceFiles(join(SRC, folder)).map((path) => ({ path, source: readFileSync(path, 'utf8') }));

const CHECKER_RULES = { forbiddenFolders: ['producer', 'app'], allowedPackages: ['saxes'] };
const COMMON_RULES = { forbiddenFolders: ['producer', 'checker', 'app'], allowedPackages: ['jszip'] };

describe('Checker independence', () => {
  it('src/checker/ imports nothing from the Producer and does not use the compare engine', () => {
    const files = load('checker');
    expect(files.length).toBeGreaterThan(5);
    expect(violations(files, CHECKER_RULES)).toEqual([]);
  });

  it('src/common/ stays neutral: it imports from neither side', () => {
    const files = load('common');
    expect(files.length).toBeGreaterThan(0);
    expect(violations(files, COMMON_RULES)).toEqual([]);
  });

  it('src/common/ stays trivial (small, and no document logic of its own)', () => {
    const files = load('common');
    const lines = files.reduce((sum, f) => sum + f.source.split('\n').filter((l) => l.trim() !== '' && !/^\s*(\/\/|\/?\*)/.test(l)).length, 0);
    expect(lines).toBeLessThan(80);
  });

  it('the Producer does not import the Checker either', () => {
    expect(violations(load('producer'), { forbiddenFolders: ['checker', 'app'], allowedPackages: ['docxodus', 'jszip', '@xmldom/xmldom'] })).toEqual([]);
  });

  // The test itself must have teeth: feed it code that breaks each rule.
  describe('the detector catches every way of importing', () => {
    const at = join(SRC, 'checker', 'fake.ts');
    const check = (source: string): Violation[] => violations([{ path: at, source }], CHECKER_RULES);

    it.each([
      ["import { compare } from '../producer/compare-client';"],
      ["import type { CompareRequest } from '../producer/protocol';"],
      ["export { REVISION_AUTHOR } from '../producer/protocol';"],
      ["export * from '../producer/protocol';"],
      ["import '../producer/compare.worker';"],
      ["const m = await import('../producer/protocol');"],
      ["const m = require('../producer/protocol');"],
      ["new Worker(new URL('../producer/compare.worker.ts', import.meta.url));"],
      ["import { x } from\n  '../producer/protocol';"],
      ["import { x } from './../producer/../producer/protocol';"],
      ["import { runComparison } from '../app/flow';"],
      ["import { compareDocuments } from 'docxodus/core';"],
      ["import { acceptRevisions } from 'docxodus';"],
      ["import JSZip from 'jszip';"],
      ["import x from '../../tests/helpers/tracked-changes';"],
    ])('%s', (source) => {
      expect(check(source).length).toBeGreaterThan(0);
    });

    it.each([
      ["import { NS } from '../common/namespaces';"],
      ["import { SaxesParser } from 'saxes';"],
      ["import { parseXml } from './xml';"],
      ["// import { compare } from '../producer/compare-client';"],
      ["/* import x from '../producer/protocol'; */"],
      ["const text = 'see the producer folder';"],
    ])('allows: %s', (source) => {
      expect(check(source)).toEqual([]);
    });
  });
});
