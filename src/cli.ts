/**
 * Sheet Sentry command line: compare two workbooks and print a Markdown review.
 *
 *   npm run review -- original.xlsx edited.xlsx [--json report.json] [--csv changes.csv] [--fail-on high]
 *
 * Exit codes: 0 = no change at or above --fail-on, 1 = changes at or above
 * --fail-on, 2 = usage or file error.
 */
import { readFile, realpath, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { compareWorkbooks, type Severity } from './core/compare.ts';
import { csvReport, jsonReport, markdownReport } from './core/report.ts';
import { readWorkbook } from './core/xlsx.ts';

const LEVELS: Severity[] = ['critical', 'high', 'medium', 'low', 'info'];

async function main(argv: string[]): Promise<number> {
  const usage =
    'Usage: npm run review -- <original.xlsx> <edited.xlsx> [--json file] [--csv file] [--fail-on critical|high|medium|low|info]';
  if (argv.length === 1 && argv[0] === '--help') {
    console.log(usage);
    return 0;
  }
  const positional: string[] = [];
  const flags = new Map<string, string>();
  let options = true;
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (options && arg === '--') options = false;
    else if (options && arg.startsWith('--')) {
      const name = arg.slice(2);
      if (
        !['json', 'csv', 'fail-on'].includes(name) ||
        flags.has(name) ||
        !argv[i + 1] ||
        argv[i + 1].startsWith('--')
      ) {
        console.error(`Invalid or missing option: ${arg}\n${usage}`);
        return 2;
      }
      flags.set(name, argv[++i]);
    } else positional.push(arg);
  }
  if (positional.length !== 2) {
    console.error(usage);
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
    // Reject aliases too: an output symlink or hard link must never overwrite either input.
    const inputs = await Promise.all([a, b].map(fileIdentity));
    const outputs = await Promise.all(
      [...flags].filter(([key]) => key === 'json' || key === 'csv').map(([, file]) => fileIdentity(file)),
    );
    for (let i = 0; i < outputs.length; i++) {
      if ([...inputs, ...outputs.slice(0, i)].some((file) => sameFile(file, outputs[i])))
        throw new Error('Report outputs must be separate files and must not overwrite an input workbook.');
    }
    result = compareWorkbooks(
      readWorkbook(new Uint8Array(await readFile(a))),
      readWorkbook(new Uint8Array(await readFile(b))),
    );
  } catch (error) {
    console.error(`Could not review the workbooks: ${(error as Error).message}`);
    return 2;
  }
  const files = { original: path.basename(a), edited: path.basename(b) };
  try {
    if (flags.get('json')) await writeFile(flags.get('json')!, jsonReport(result, files));
    if (flags.get('csv')) await writeFile(flags.get('csv')!, csvReport(result));
  } catch (error) {
    console.error(`Could not write a report: ${(error as Error).message}`);
    return 2;
  }
  process.stdout.write(markdownReport(result, files));
  const threshold = LEVELS.indexOf(failOn);
  const failing = LEVELS.slice(0, threshold + 1).reduce((n, s) => n + result.counts[s], 0);
  return failing > 0 ? 1 : 0;
}

async function fileIdentity(file: string) {
  const resolved = await realpath(file).catch(async (error: NodeJS.ErrnoException) => {
    if (error.code !== 'ENOENT') throw error;
    return path.join(await realpath(path.dirname(path.resolve(file))), path.basename(file));
  });
  const info = await stat(file).catch((error: NodeJS.ErrnoException) => {
    if (error.code !== 'ENOENT') throw error;
    return undefined;
  });
  return { path: process.platform === 'win32' ? resolved.toLowerCase() : resolved, info };
}

function sameFile(a: Awaited<ReturnType<typeof fileIdentity>>, b: Awaited<ReturnType<typeof fileIdentity>>) {
  return (
    a.path === b.path ||
    !!(a.info && b.info && a.info.ino !== 0 && a.info.ino === b.info.ino && a.info.dev === b.info.dev)
  );
}

process.exitCode = await main(process.argv.slice(2));
