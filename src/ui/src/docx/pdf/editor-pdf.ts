/** `<docx-editor>`'s PDF export: the open document, its pictures and the page's local fonts. */
import { collectLocalFontFiles } from '../../export/pdf-font-faces';
import { currentModel } from '../current-model';
import type { EditorCore } from '../editor-core';
import { decodePictureInBrowser } from './browser-picture-decoder';
import { exportDocxPdf, type DocxPdfOptions, type DocxPdfResult } from './export-pdf';
import { editorPageStatus } from './pdf-context';
import { selectionModel } from './selection-model';

/**
 * What a host may set; the editor supplies the pictures and the text measurer. Besides the page
 * range of `exportDocxPdf`, `pages` takes Word's other two ranges: `current` (the page Print
 * Layout shows, as its status bar counts it; only in Print Layout) and `selection` (only the
 * selected body content, laid out on its own pages).
 */
export type EditorPdfOptions = Omit<DocxPdfOptions, 'media' | 'measurer' | 'pages'> & {
	pages?: DocxPdfOptions['pages'] | 'current' | 'selection';
};

export async function editorPdf(
	core: EditorCore,
	options: EditorPdfOptions = {},
): Promise<DocxPdfResult> {
	const { pages, ...rest } = options;
	let model = currentModel(core);
	let range: DocxPdfOptions['pages'];
	if (pages === 'current') {
		const status = editorPageStatus(core);
		if (!status) throw new RangeError('The current page is only known in Print Layout.');
		range = { from: status.current, to: status.current };
	} else if (pages === 'selection') {
		const state = core.view?.state;
		const selected = state ? selectionModel(state, model) : undefined;
		if (!selected) throw new RangeError('There is no selection to export.');
		model = selected;
	} else range = pages;
	const media = (partName: string) =>
		core.loaded?.media?.get(partName) ?? core.inserts.pendingMedia.get(partName)?.bytes;
	// Fonts the page already holds as data: or blob: URLs; remote fonts are never fetched.
	const fonts =
		options.fonts ?? (typeof document === 'undefined' ? [] : await collectLocalFontFiles(document));
	return exportDocxPdf(model, {
		decodePicture: decodePictureInBrowser,
		...rest,
		...(range ? { pages: range } : {}),
		fonts,
		media,
	});
}
