import type { DocumentModel, HeaderFooterContent, HeaderFooterSlots, TextRun } from '../model.js';
import { formatNoteNumber } from '../notes.js';
import { fieldName } from '../field-runs.js';
import { dateFieldResult } from '../field-date.js';
import type { LayoutPageBox } from './result.js';

export interface PageFieldContext {
	page: string;
	numPages: string;
	sectionPages: string;
	/** DATE and TIME fields update to the pagination time. */
	now?: Date;
}

/** Recalculate page/date fields; other fields retain their saved result. */
export function fieldDisplayText(run: TextRun, values: PageFieldContext): string {
	if (!run.field) return run.text;
	const name = fieldName(run.field.instr);
	if (name === 'PAGE') return values.page;
	if (name === 'NUMPAGES') return values.numPages;
	if (name === 'SECTIONPAGES') return values.sectionPages;
	if (name === 'DATE' || name === 'TIME')
		return dateFieldResult(name, run.field.instr, values.now ?? new Date());
	return run.text;
}

/** Numeric page values, before formatting as Roman, letter or decimal labels. */
export function pageNumberValues(model: DocumentModel, pages: LayoutPageBox[]): number[] {
	const sharedStarts = new Map<number, number>();
	const seen = new Set<number>();
	for (const page of pages)
		for (const index of page.sectionIndices ?? [page.sectionIndex]) {
			if (!seen.has(index) && index !== page.sectionIndex) sharedStarts.set(index, page.index);
			seen.add(index);
		}
	let previous = 0;
	return pages.map((page) => {
		const start = model.sections?.[page.sectionIndex]?.pageNumbering?.start;
		// With distinct odd/even headers, native Word aligns a restart's parity with the
		// shared physical starting page, decrementing an even start on an odd page.
		const sharedPage = sharedStarts.get(page.sectionIndex);
		const offset =
			model.evenAndOddHeaders &&
			sharedPage !== undefined &&
			start !== undefined &&
			start % 2 !== (sharedPage + 1) % 2
				? 1
				: 0;
		const value = start === undefined ? previous + 1 : start + page.pageInSection - offset;
		previous = value;
		for (const index of page.sectionIndices ?? []) {
			if (index === page.sectionIndex) continue;
			const restart = model.sections?.[index]?.pageNumbering?.start;
			if (restart !== undefined)
				previous =
					restart - (model.evenAndOddHeaders && restart % 2 !== (page.index + 1) % 2 ? 1 : 0);
		}
		return value;
	});
}

/** Word page labels, honoring each section's restart and number format. */
export function pageNumbers(model: DocumentModel, pages: LayoutPageBox[]): string[] {
	return pageNumberValues(model, pages).map((value, index) =>
		formatNoteNumber(
			value,
			model.sections?.[pages[index]!.sectionIndex]?.pageNumbering?.format ?? 'decimal',
		),
	);
}

/** Physical pages occupied by each section, including shared continuous-section pages. */
export function sectionPageCounts(pages: LayoutPageBox[]): ReadonlyMap<number, number> {
	const counts = new Map<number, number>();
	for (const page of pages)
		for (const index of page.sectionIndices ?? [page.sectionIndex])
			counts.set(index, (counts.get(index) ?? 0) + 1);
	return counts;
}

/** Resolve first/even/default header or footer slots, inheriting missing slots by section. */
export function headerFooterForPage(
	model: DocumentModel,
	page: LayoutPageBox,
	pageNumber: number,
	kind: 'headers' | 'footers',
): HeaderFooterContent | undefined {
	const sections = model.sections ?? [];
	const section = sections[page.sectionIndex];
	if (!section) return undefined;
	const slot: keyof HeaderFooterSlots =
		page.pageInSection === 0 && section.titlePage
			? 'first'
			: model.evenAndOddHeaders && pageNumber % 2 === 0
				? 'even'
				: 'default';
	for (let index = page.sectionIndex; index >= 0; index--) {
		const content = sections[index]?.[kind]?.[slot];
		if (content) return content;
	}
	return undefined;
}
