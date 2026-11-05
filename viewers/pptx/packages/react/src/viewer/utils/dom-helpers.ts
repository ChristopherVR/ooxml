/**
 * DOM utility helpers: escapeHtml, downloadBlob.
 *
 * `sanitizeDownloadFilename` and `downloadBlob` now live once in
 * `pptx-viewer-shared` (`export/download-helpers`), which sanitizes every
 * filename internally, so this module just re-exports them and keeps the
 * historical React import path.
 */
export { downloadBlob, sanitizeDownloadFilename } from 'ooxml-ui/pptx';

// HTML entity escaping is shared (see `export/print-document`).
export { escapeHtml } from 'ooxml-ui/pptx';
