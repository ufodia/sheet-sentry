import { spawnSync } from 'node:child_process';
import { linkSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterAll, describe, expect, it } from 'vitest';
import { writeWorkbook } from '../src/core/xlsx.ts';

const directory = mkdtempSync(path.join(os.tmpdir(), 'sheet-sentry-cli-'));
const original = path.join(directory, 'original.xlsx');
const edited = path.join(directory, 'edited.xlsx');
const bytes = writeWorkbook([{ name: 'S', cells: { A1: 1 } }]);
writeFileSync(original, bytes);
writeFileSync(edited, bytes);
afterAll(() => rmSync(directory, { recursive: true, force: true }));

const run = (...args: string[]) =>
  spawnSync(process.execPath, ['--import', 'tsx', 'src/cli.ts', ...args], {
    encoding: 'utf8',
    timeout: 15_000,
  });

describe('CLI file and option contracts', () => {
  it('prints help successfully and rejects unknown, duplicate or missing options', () => {
    expect(run('--help').status).toBe(0);
    for (const args of [
      [original, edited, '--jsoon', 'x'],
      [original, edited, '--json'],
      [original, edited, '--json', '--csv', 'x'],
      [original, edited, '--fail-on', 'high', '--fail-on', 'low'],
    ])
      expect(run(...args).status).toBe(2);
  });
  it('refuses to overwrite either input with a report', () => {
    expect(run(original, edited, '--json', original).status).toBe(2);
    expect(run(original, edited, '--csv', edited).status).toBe(2);
    expect(readFileSync(original)).toEqual(Buffer.from(bytes));
    expect(readFileSync(edited)).toEqual(Buffer.from(bytes));
  });
  it('refuses an output hard-link alias of an input', () => {
    const alias = path.join(directory, 'alias.json');
    linkSync(original, alias);
    expect(run(original, edited, '--json', alias).status).toBe(2);
    expect(readFileSync(original)).toEqual(Buffer.from(bytes));
  });
  it('refuses two report formats targeting the same file', () => {
    const output = path.join(directory, 'same.txt');
    expect(run(original, edited, '--json', output, '--csv', output).status).toBe(2);
  });
  it('returns documented output errors and writes valid reports', () => {
    expect(run(original, edited, '--json', directory).status).toBe(2);
    const json = path.join(directory, 'report.json');
    const csv = path.join(directory, 'changes.csv');
    const result = run(original, edited, '--json', json, '--csv', csv);
    expect(result.status).toBe(0);
    expect(JSON.parse(readFileSync(json, 'utf8')).verdict).toBe('clean');
    expect(readFileSync(csv, 'utf8')).toContain('severity,kind');
    expect(result.stdout).toContain('# Sheet Sentry review');
  });
  it('applies the requested severity threshold', () => {
    const changed = path.join(directory, 'changed.xlsx');
    writeFileSync(changed, writeWorkbook([{ name: 'S', cells: { A1: 2 } }]));
    expect(run(original, changed).status).toBe(0);
    expect(run(original, changed, '--fail-on', 'medium').status).toBe(1);
  });
});
