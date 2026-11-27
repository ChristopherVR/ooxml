// Line numbers on a laid-out page (`w:lnNumType`): DOM-free, shared by Print Layout and the PDF
// export (`page-marks.ts`).
import { resolveParagraphFormatting, type DocumentModel } from '../index';
import type { LayoutResult } from '../layout/index';

/** Line numbering of a section, as the section model gives it. */
export interface PrintLineNumbering {
	countBy: number;
	start: number;
	restart: 'newPage' | 'newSection' | 'continuous';
	distanceTwips?: number;
}

/** One printed line number: its line's box on the page, and where the number ends. */
export interface LineNumberMark {
	number: number;
	/** The number's right edge: the column's left edge less the numbering distance. */
	rightPx: number;
	topPx: number;
	heightPx: number;
	/** The line's baseline on the page. */
	baselinePx: number;
}

type Page = LayoutResult['pages'][number];

/** Advances the body-line counter across columns/pages, skipping suppressed paragraphs. */
export function lineNumberMarks(
	page: Page,
	settings: PrintLineNumbering,
	counter: { next: number; section: number },
	suppressed?: ReadonlySet<string>,
): LineNumberMark[] {
	const newSection = counter.section !== page.sectionIndex;
	if (
		counter.section === -1 ||
		settings.restart === 'newPage' ||
		(newSection && settings.restart === 'newSection')
	)
		counter.next = settings.start;
	counter.section = page.sectionIndex;
	const marks: LineNumberMark[] = [];
	const gap = (settings.distanceTwips ?? 360) / 15;
	for (const column of page.columns)
		for (const block of column.blocks) {
			if (block.kind !== 'paragraph' || suppressed?.has(block.blockId)) continue;
			for (const line of block.lines) {
				const number = counter.next++;
				if (number % settings.countBy !== 0) continue;
				const topPx = page.marginTopPx + block.yPx + line.yPx;
				marks.push({
					number,
					rightPx: page.marginLeftPx + column.xPx - gap,
					topPx,
					heightPx: line.heightPx,
					baselinePx: topPx + (line.baselinePx ?? line.heightPx * 0.8),
				});
			}
		}
	return marks;
}

/**
 * A page's line numbers: each section sharing the page (continuous section breaks) numbers its own
 * columns with its own settings; sections without settings print none.
 */
export function pageLineNumberMarks(
	page: Page,
	numbering: ReadonlyArray<PrintLineNumbering | undefined>,
	counter: { next: number; section: number },
	suppressed?: ReadonlySet<string>,
): LineNumberMark[] {
	const bands = new Map<number, Page['columns']>();
	for (const column of page.columns) {
		const index = column.sectionIndex ?? page.sectionIndex;
		const band = bands.get(index) ?? [];
		band.push(column);
		bands.set(index, band);
	}
	const marks: LineNumberMark[] = [];
	for (const [sectionIndex, columns] of bands) {
		const settings = numbering[sectionIndex];
		if (settings)
			marks.push(
				...lineNumberMarks({ ...page, sectionIndex, columns }, settings, counter, suppressed),
			);
	}
	return marks;
}

/** Body paragraphs with effective `w:suppressLineNumbers`, which the counter skips. */
export function suppressedLineNumberParagraphs(model: DocumentModel): Set<string> {
	return new Set(
		model.blocks
			.filter(
				(block) =>
					block.type === 'paragraph' &&
					(model.paragraphStyles
						? resolveParagraphFormatting(block, model.paragraphStyles).suppressLineNumbers
						: block.suppressLineNumbers),
			)
			.map((block) => block.id),
	);
}
