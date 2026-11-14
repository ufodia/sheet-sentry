/**
 * Row alignment between two versions of a sheet, so an inserted row does not
 * make every cell below it look changed.
 *
 * Rows are compared by a position-independent signature. An LCS (Myers'
 * O(ND) algorithm) matches identical rows; inside each gap, deleted and
 * inserted rows are paired in order as "modified" rows.
 */

export type RowOp =
  | { type: 'same' | 'modified'; a: number; b: number }
  | { type: 'deleted'; a: number }
  | { type: 'inserted'; b: number };

/** Longest common subsequence pairs (indices into a and b), via Myers' algorithm. */
export function lcsPairs(a: string[], b: string[], maxD = 1500): [number, number][] {
  const n = a.length;
  const m = b.length;
  const max = Math.min(n + m, maxD);
  const offset = max + 1;
  const v = new Int32Array(2 * max + 3);
  const trace: Int32Array[] = [];
  let found = -1;
  for (let d = 0; d <= max; d++) {
    // Keep only the diagonals reachable at this depth: memory grows with d², not with sheet size.
    trace.push(v.slice(offset - d, offset + d + 1));
    for (let k = -d; k <= d; k += 2) {
      let x =
        k === -d || (k !== d && v[offset + k - 1] < v[offset + k + 1])
          ? v[offset + k + 1]
          : v[offset + k - 1] + 1;
      let y = x - k;
      while (x < n && y < m && a[x] === b[y]) {
        x++;
        y++;
      }
      v[offset + k] = x;
      if (x >= n && y >= m) {
        found = d;
        break;
      }
    }
    if (found >= 0) break;
  }
  if (found < 0) return fallback(a, b);
  // Backtrack.
  const pairs: [number, number][] = [];
  let x = n;
  let y = m;
  for (let d = found; d > 0; d--) {
    const vv = trace[d];
    const at = (k: number) => vv[k + d];
    const k = x - y;
    const prevK = k === -d || (k !== d && at(k - 1) < at(k + 1)) ? k + 1 : k - 1;
    const prevX = at(prevK);
    const prevY = prevX - prevK;
    while (x > prevX && y > prevY) {
      pairs.push([x - 1, y - 1]);
      x--;
      y--;
    }
    x = prevX;
    y = prevY;
  }
  while (x > 0 && y > 0) {
    pairs.push([x - 1, y - 1]);
    x--;
    y--;
  }
  return pairs.reverse();
}

/** When the sheets differ too much for a bounded search, align rows by position. */
function fallback(a: string[], b: string[]): [number, number][] {
  const pairs: [number, number][] = [];
  for (let i = 0; i < Math.min(a.length, b.length); i++) if (a[i] === b[i]) pairs.push([i, i]);
  return pairs;
}

export function alignRows(a: string[], b: string[]): RowOp[] {
  const pairs = lcsPairs(a, b);
  const ops: RowOp[] = [];
  let i = 0;
  let j = 0;
  const flushGap = (ai: number, bj: number) => {
    const dels: number[] = [];
    const ins: number[] = [];
    while (i < ai) dels.push(i++);
    while (j < bj) ins.push(j++);
    const paired = Math.min(dels.length, ins.length);
    for (let k = 0; k < paired; k++) ops.push({ type: 'modified', a: dels[k], b: ins[k] });
    for (let k = paired; k < dels.length; k++) ops.push({ type: 'deleted', a: dels[k] });
    for (let k = paired; k < ins.length; k++) ops.push({ type: 'inserted', b: ins[k] });
  };
  for (const [pa, pb] of pairs) {
    flushGap(pa, pb);
    ops.push({ type: 'same', a: pa, b: pb });
    i = pa + 1;
    j = pb + 1;
  }
  flushGap(a.length, b.length);
  return ops;
}

/**
 * Map an original row index to the edited sheet for reference remapping.
 * Deleted rows map to the next surviving row, as Excel does when a range
 * shrinks.
 */
export function rowMapper(ops: RowOp[], originalRows: number): (row: number) => number {
  const map = new Int32Array(originalRows + 1).fill(-1);
  for (const op of ops) if (op.type === 'same' || op.type === 'modified') map[op.a] = op.b;
  // Fill deleted rows (and the row after the end) from the next mapped row.
  let next = -1;
  let lastB = -1;
  for (const op of ops) if (op.type !== 'deleted' && op.type !== 'inserted') lastB = Math.max(lastB, op.b);
  map[originalRows] = lastB + 1;
  for (let r = originalRows; r >= 0; r--) {
    if (map[r] >= 0) next = map[r];
    else map[r] = next;
  }
  return (row) => (row >= 0 && row < originalRows ? map[row] : row + (map[originalRows] - originalRows));
}
