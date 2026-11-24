/** `<docx-editor>`'s PDF export: the open document, its pictures and the page's local fonts. */
import { collectLocalFontFiles } from '../../export/pdf-font-faces';
import { currentModel } from '../current-model';
import type { EditorCore } from '../editor-core';
import { decodePictureInBrowser } from './browser-picture-decoder';
import { exportDocxPdf, type DocxPdfOptions, type DocxPdfResult } from './export-pdf';

/** What a host may set; the editor supplies the pictures and the text measurer. */
export type EditorPdfOptions = Omit<DocxPdfOptions, 'media' | 'measurer'>;

export async function editorPdf(
	core: EditorCore,
	options: EditorPdfOptions = {},
): Promise<DocxPdfResult> {
	const model = currentModel(core);
	const media = (partName: string) =>
		core.loaded?.media?.get(partName) ?? core.inserts.pendingMedia.get(partName)?.bytes;
	// Fonts the page already holds as data: or blob: URLs; remote fonts are never fetched.
	const fonts =
		options.fonts ?? (typeof document === 'undefined' ? [] : await collectLocalFontFiles(document));
	return exportDocxPdf(model, {
		decodePicture: decodePictureInBrowser,
		...options,
		fonts,
		media,
	});
}
