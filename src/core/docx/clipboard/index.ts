/** Reading the HTML Word puts on the clipboard (see `word-html.ts`). */
export {
	normalizeWordClipboardHtml,
	type WordClipboardOptions,
	type WordPasteMode,
} from './word-html';
export { clipboardFragment, isWordClipboardHtml } from './word-fragment';
export type { WordClipboardList } from './word-list-scan';
export { styleIdForName, styleNameFromClass } from './word-styles';
