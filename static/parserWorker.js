/**
 * parserWorker.js
 * ───────────────
 * Runs entirely off the main thread.
 *
 * The main thread sends a raw clipboard string.  This worker splits it into
 * a 2-D matrix and posts the result back.  For a 10 000-row paste the parse
 * work (~25 ms) no longer blocks the UI; the grid stays interactive.
 *
 * Message in:
 *   { type: 'parse', text: string, startR: number, startC: number, gridId: string }
 *
 * Message out:
 *   { type: 'result', matrix: string[][], startR: number, startC: number, gridId: string }
 */

self.onmessage = function (e) {
  const { type, text, startR, startC, gridId } = e.data;
  if (type !== 'parse' || typeof text !== 'string') return;

  // ── Split into lines ─────────────────────────────────────────────────────
  // Handle Windows (\r\n), Unix (\n), and legacy Mac (\r) line endings.
  const lines = text.split(/\r?\n|\r/);

  // Strip the trailing blank line that copy operations append.
  if (lines.length > 0 && lines[lines.length - 1].trim() === '') {
    lines.pop();
  }

  // ── Split each line into cells ────────────────────────────────────────────
  // Prefer tab-delimited (Excel / Sheets default); fall back to comma (CSV).
  const matrix = lines.map(line =>
    line.includes('\t') ? line.split('\t') : line.split(',')
  );

  self.postMessage({ type: 'result', matrix, startR, startC, gridId });
};