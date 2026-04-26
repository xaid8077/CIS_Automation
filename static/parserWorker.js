/**
 * parserWorker.js
 * Runs entirely off the main thread.
 *
 * Message in:
 *   { type: 'parse', text: string, startR: number, startC: number, gridId: string, requestId: number }
 *
 * Message out:
 *   { type: 'result', matrix: string[][], startR: number, startC: number, gridId: string, requestId: number }
 *   { type: 'error', message: string, startR: number, startC: number, gridId: string, requestId: number }
 */

self.onmessage = function (e) {
  const { type, text, startR, startC, gridId, requestId } = e.data || {};
  if (type !== 'parse' || typeof text !== 'string') return;

  try {
    const lines = text.split(/\r?\n|\r/);

    if (lines.length > 0 && lines[lines.length - 1].trim() === '') {
      lines.pop();
    }

    const matrix = lines.map(line =>
      line.includes('\t') ? line.split('\t') : line.split(',')
    );

    self.postMessage({ type: 'result', matrix, startR, startC, gridId, requestId });
  } catch (err) {
    self.postMessage({
      type: 'error',
      message: err && err.message ? err.message : 'Clipboard parse failed',
      startR,
      startC,
      gridId,
      requestId,
    });
  }
};
