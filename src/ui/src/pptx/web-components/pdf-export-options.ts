/**
 * The options step of File > Export > PDF as every binding runs it: mount
 * `pptx-ui-pdf-export-dialog` in the viewer (so it inherits the viewer theme), wait for Export or
 * Cancel, remove it, and resolve to the slide selection to export (`null` when cancelled). The
 * picture export asks nothing (`pdfExportAsksForRange`) and resolves to `{}` at once. The
 * binding's own UI commands call this; its programmatic `exportPdf()` never does.
 */
import type { PdfSlideSelection } from '../export/pdf-vector-select';
import { pdfExportAsksForRange } from '../render/pdf-export-range';
import type {
	PdfExportDialogConfirmEvent,
	PdfExportDialogTranslate,
	PptxUiPdfExportDialogElement,
} from './pdf-export-dialog';

export interface PdfExportOptionsInput {
	/** Options > Advanced > Save as PDF. */
	settings: { mode: 'vector' | 'pictures'; includeHiddenSlides: boolean };
	/** The deck's slides (only `hidden` is read). */
	slides: readonly { hidden?: boolean }[];
	/** 0-based index of the current slide. */
	currentSlide: number;
	translate?: PdfExportDialogTranslate | undefined;
}

/** Ask for the PDF range in `container`; see the module comment. */
export function choosePdfExportSelection(
	container: HTMLElement | null | undefined,
	input: PdfExportOptionsInput,
): Promise<PdfSlideSelection | null> {
	if (!pdfExportAsksForRange(input.settings) || !container) {
		return Promise.resolve({});
	}
	const doc = container.ownerDocument;
	if (!doc.defaultView?.customElements.get('pptx-ui-pdf-export-dialog')) {
		return Promise.resolve({});
	}
	const dialog = doc.createElement('pptx-ui-pdf-export-dialog') as PptxUiPdfExportDialogElement;
	dialog.slides = input.slides;
	dialog.currentSlide = input.currentSlide;
	dialog.includeHiddenSlides = input.settings.includeHiddenSlides;
	dialog.translateLabel = input.translate;
	return new Promise((resolve) => {
		let selection: PdfSlideSelection | null = null;
		dialog.addEventListener('pdf-export-confirm', (event) => {
			selection = (event as PdfExportDialogConfirmEvent).detail.selection;
		});
		dialog.addEventListener('dialog-close', () => {
			dialog.remove();
			resolve(selection);
		});
		container.append(dialog);
		dialog.open = true;
	});
}
