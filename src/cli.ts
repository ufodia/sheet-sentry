/**
 * Sheet Sentry command line: compare two workbooks and print a Markdown review.
 *
 *   npm run review -- original.xlsx edited.xlsx [--json report.json] [--csv changes.csv] [--fail-on high]
 *
 * Exit codes: 0 = no change at or above --fail-on, 1 = changes at or above
 * --fail-on, 2 = usage or file error.
 */
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { compareWorkbooks, type Severity } from './core/compare.ts';
import { csvReport, jsonReport, markdownReport } from './core/report.ts';
import { readWorkbook } from './core/xlsx.ts';

const LEVELS: Severity[] = ['critical', 'high', 'medium', 'low', 'info'];

async function main(argv: string[]): Promise<number> {
  const positional: string[] = [];
  const flags = new Map<string, string>();
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg.startsWith('--')) flags.set(arg.slice(2), argv[++i] ?? '');
    else positional.push(arg);
  }
  if (positional.length !== 2 || flags.has('help')) {
    console.error(
      'Usage: npm run review -- <original.xlsx> <edited.xlsx> [--json file] [--csv file] [--fail-on critical|high|medium|low]',
    );
    return 2;
  }
  const failOn = (flags.get('fail-on') ?? 'critical') as Severity;
  if (!LEVELS.includes(failOn)) {
    console.error(`--fail-on must be one of ${LEVELS.join(', ')}`);
    return 2;
  }
  const [a, b] = positional;
  let result;
  try {
    result = compareWorkbooks(
      readWorkbook(new Uint8Array(await readFile(a))),
      readWorkbook(new Uint8Array(await readFile(b))),
    );
  } catch (error) {
    console.error(`Could not read the workbooks: ${(error as Error).message}`);
    return 2;
  }
  const files = { original: path.basename(a), edited: path.basename(b) };
  if (flags.get('json')) await writeFile(flags.get('json')!, jsonReport(result, files));
  if (flags.get('csv')) await writeFile(flags.get('csv')!, csvReport(result));
  process.stdout.write(markdownReport(result, files));
  const threshold = LEVELS.indexOf(failOn);
  const failing = LEVELS.slice(0, threshold + 1).reduce((n, s) => n + result.counts[s], 0);
  return failing > 0 ? 1 : 0;
}

process.exitCode = await main(process.argv.slice(2));
