/**
 * Page borders and line numbers for the Word PDF export, in the places Print Layout draws them
 * (`print-page-borders.ts`, `print-line-numbers.ts`): the borders as vector lines around the
 * section's border box, the numbers as real text right-aligned at the numbering distance left of
 * each counted line, on its baseline, in the Line Number style's font. DOM-free.
 */
import {
	resolveRunFormatting,
	type DocumentModel,
	type PageBorders,
	type TextRun,
} from 'ooxml-core/docx';
import {
	DEFAULT_FONT_SIZE_PT,
	layoutBorder,
	type LayoutFragment,
	type LayoutPageBox,
	type TextMeasurer,
} from 'ooxml-core/docx/layout';
import { pageBorderRect } from '../print-page-borders';
import {
	pageLineNumberMarks,
	suppressedLineNumberParagraphs,
	type LineNumberMark,
} from '../print-line-numbers';
import { paragraphFrame, type DocxPdfPainter } from './pdf-blocks';
import type { DocxPdfReportBuilder } from './pdf-report';

/** The marks of one page. */
export interface DocxPdfPageMarks {
	borders?: PageBorders;
	lineNumbers: LineNumberMark[];
}

/** Every page's marks; line numbers count through all pages, written or not. */
export function pageMarks(
	model: DocumentModel,
	pages: readonly LayoutPageBox[],
): DocxPdfPageMarks[] {
	const numbering = (model.sections ?? []).map((section) => section.lineNumberSettings);
	const numbered = numbering.some(Boolean);
	const suppressed = numbered ? suppressedLineNumberParagraphs(model) : undefined;
	const counter = { next: 1, section: -1 };
	return pages.map((page) => {
		const borders = model.sections?.[page.sectionIndex]?.pageBorders;
		return {
			...(borders ? { borders } : {}),
			lineNumbers: numbered ? pageLineNumberMarks(page, numbering, counter, suppressed) : [],
		};
	});
}

function paintBorders(
	painter: DocxPdfPainter,
	page: LayoutPageBox,
	borders: PageBorders,
	model: DocumentModel,
	report: DocxPdfReportBuilder,
): void {
	const rect = pageBorderRect(page, borders);
	if (!rect) return;
	const side = (name: 'top' | 'left' | 'bottom' | 'right') => {
		const resolved = layoutBorder(borders[name], model.theme);
		if (borders[name]?.art && resolved)
			report.note('Page art borders are written to PDF as plain lines.');
		return resolved && { ...resolved, spacePx: 0 };
	};
	const sides = {
		top: side('top'),
		left: side('left'),
		bottom: side('bottom'),
		right: side('right'),
	};
	const present = Object.fromEntries(Object.entries(sides).filter(([, value]) => value));
	if (!Object.keys(present).length) return;
	paragraphFrame(
		painter.page,
		{ leftPx: rect.leftPx, widthPx: rect.widthPx, borders: present },
		0,
		rect.topPx,
		rect.heightPx,
	);
}

/** The Line Number character style over the document defaults, with theme fonts resolved. */
function lineNumberFont(model: DocumentModel) {
	const formatting = resolveRunFormatting({ text: '', style: 'LineNumber' } as TextRun, {
		runCatalog: model.characterStyles,
	});
	const role = formatting.fontTheme?.ascii ?? formatting.fontTheme?.hAnsi;
	const family =
		formatting.fontFamily ?? (role ? model.theme?.fonts[role]?.latin : undefined) ?? 'Calibri';
	return {
		fontFamily: family,
		fontSizePt: formatting.fontSize ?? DEFAULT_FONT_SIZE_PT,
		...(formatting.color ? { color: formatting.color } : {}),
		...(formatting.bold ? { bold: true } : {}),
		...(formatting.italic ? { italic: true } : {}),
	};
}

function paintLineNumbers(
	painter: DocxPdfPainter,
	marks: readonly LineNumberMark[],
	model: DocumentModel,
	measurer: TextMeasurer,
): void {
	if (!marks.length) return;
	const font = lineNumberFont(model);
	const spec = {
		family: font.fontFamily,
		sizePx: (font.fontSizePt * 96) / 72,
		...(font.bold ? { bold: true } : {}),
		...(font.italic ? { italic: true } : {}),
	};
	for (const mark of marks) {
		const text = String(mark.number);
		const widthPx = measurer.widthOf(text, spec);
		const fragment: LayoutFragment = { ...font, text, xPx: 0, widthPx, runIndex: 0 };
		const line = {
			yPx: 0,
			heightPx: mark.heightPx,
			baselinePx: mark.baselinePx - mark.topPx,
			fragments: [fragment],
			sourceStart: 0,
			sourceEnd: 0,
		};
		painter.text.draw(painter.page, fragment, mark.rightPx - widthPx, line, mark.topPx);
	}
}

/**
 * Draws a page's marks: `back` the borders set behind the text (`w:zOrder="back"`), `front` the
 * line numbers and the other borders.
 */
export function paintPageMarks(
	painter: DocxPdfPainter,
	page: LayoutPageBox,
	marks: DocxPdfPageMarks | undefined,
	layer: 'back' | 'front',
	model: DocumentModel,
	measurer: TextMeasurer,
	report: DocxPdfReportBuilder,
): void {
	if (!marks) return;
	const back = marks.borders?.zOrder === 'back';
	if (marks.borders && back === (layer === 'back'))
		paintBorders(painter, page, marks.borders, model, report);
	if (layer === 'front') paintLineNumbers(painter, marks.lineNumbers, model, measurer);
}
