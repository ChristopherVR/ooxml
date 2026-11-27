// The DOM half of line numbering; the marks are shared with the PDF export
// (`ooxml-core/docx/pdf-export`, whose page-marks file loads without the writer).
import type { LineNumberMark } from 'ooxml-core/docx/pdf-export';

export {
	lineNumberMarks,
	pageLineNumberMarks,
	suppressedLineNumberParagraphs,
	type LineNumberMark,
	type PrintLineNumbering,
} from 'ooxml-core/docx/pdf-export';

/** Width of the box a Print Layout label right-aligns its number in. */
const LABEL_WIDTH_PX = 40;

/** A line number as a Print Layout label. */
export function lineNumberLabel(mark: LineNumberMark): HTMLElement {
	const label = document.createElement('span');
	label.className = 'dve-print-line-number';
	label.textContent = String(mark.number);
	Object.assign(label.style, {
		left: `${mark.rightPx - LABEL_WIDTH_PX}px`,
		top: `${mark.topPx}px`,
		height: `${mark.heightPx}px`,
		lineHeight: `${mark.heightPx}px`,
	});
	return label;
}
