/** Writes the demo workbooks to examples/ so they can be opened in a spreadsheet app or used with the CLI. */
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { demoWorkbooks } from '../src/demo.ts';

const dir = path.resolve(import.meta.dirname, '..', 'examples');
await mkdir(dir, { recursive: true });
const { original, edited } = demoWorkbooks();
await writeFile(path.join(dir, 'q3-budget.xlsx'), original);
await writeFile(path.join(dir, 'q3-budget-ai-edit.xlsx'), edited);
console.log('Wrote examples/q3-budget.xlsx and examples/q3-budget-ai-edit.xlsx');
