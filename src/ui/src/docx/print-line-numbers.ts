import { resolveParagraphFormatting, type DocumentModel } from 'ooxml-core/docx';
import type { LayoutResult } from 'ooxml-core/docx/layout';

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

/** Width of the box a Print Layout label right-aligns its number in. */
const LABEL_WIDTH_PX = 40;

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
