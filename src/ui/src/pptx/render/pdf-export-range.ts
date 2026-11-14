/**
 * The Range step of File > Export > PDF, like PowerPoint's Save as PDF options: All slides,
 * the Current slide, or Slides from-to, plus "Include hidden slides". Pure: the dialog
 * (`pptx-ui-pdf-export-dialog`) edits a {@link PdfExportChoice}; the binding turns the confirmed
 * choice into the slide selection the vector export reads.
 */
import { selectPdfSlides, type PdfSlideSelection } from '../export/pdf-vector-select';

export type PdfExportRange = 'all' | 'current' | 'slides';

/** What the PDF options dialog asks. Slide numbers are 1-based. */
export interface PdfExportChoice {
	range: PdfExportRange;
	from: number;
	to: number;
	includeHiddenSlides: boolean;
}

/** Why a choice cannot be exported (a dictionary key suffix under `pptx.pdfExport.dialog.`). */
export type PdfExportChoiceError = 'rangeInvalid' | 'rangeEmpty';

/** The choice the dialog opens with: every slide, hidden ones as Options > Advanced says. */
export function initialPdfExportChoice(
	slideCount: number,
	includeHiddenSlides: boolean,
): PdfExportChoice {
	return { range: 'all', from: 1, to: Math.max(1, slideCount), includeHiddenSlides };
}

/**
 * The selection a confirmed choice exports. The current slide is exported even when it is
 * hidden: choosing it is explicit.
 */
export function resolvePdfExportSelection(
	choice: PdfExportChoice,
	currentIndex: number,
): PdfSlideSelection {
	switch (choice.range) {
		case 'current':
			return { slideIndexes: [Math.max(0, Math.floor(currentIndex))], includeHiddenSlides: true };
		case 'slides':
			return {
				from: choice.from,
				to: choice.to,
				includeHiddenSlides: choice.includeHiddenSlides,
			};
		default:
			return { includeHiddenSlides: choice.includeHiddenSlides };
	}
}

/** Whether `choice` can be exported from `slides`; the error explains why not. */
export function validatePdfExportChoice(
	choice: PdfExportChoice,
	slides: readonly { hidden?: boolean }[],
	currentIndex: number,
): PdfExportChoiceError | undefined {
	if (choice.range === 'slides') {
		const { from, to } = choice;
		if (
			!Number.isInteger(from) ||
			!Number.isInteger(to) ||
			from < 1 ||
			to < from ||
			to > slides.length
		)
			return 'rangeInvalid';
	}
	const selection = resolvePdfExportSelection(choice, currentIndex);
	return selectPdfSlides(slides, selection).length ? undefined : 'rangeEmpty';
}

/** Parse a slide-number field: a whole number, or `NaN`. */
export function parsePdfSlideNumber(value: string): number {
	const text = value.trim();
	return /^\d+$/u.test(text) ? Number(text) : Number.NaN;
}
