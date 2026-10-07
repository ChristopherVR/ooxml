import { SectionFlow } from './section-flow.js';
import { balanceColumns, canBalanceBlock } from './balance-columns.js';
import type { TextMeasurer } from './measure.js';
import type { LayoutDocumentInput, LayoutSection } from './input.js';
import type { LayoutPageBox, LayoutResult } from './result.js';
import type { Exclusion } from './wrap.js';

/**
 * Word starts an even-page (odd-page) section on the next even (odd) page, leaving a blank page
 * when needed. The blank page belongs to the previous section.
 */
function insertParityBlank(pages: LayoutPageBox[], section: LayoutSection): void {
	const previous = pages.at(-1);
	if (!previous || (section.break !== 'evenPage' && section.break !== 'oddPage')) return;
	const nextNumber = pages.length + 1;
	const wanted = section.break === 'evenPage' ? 0 : 1;
	if (nextNumber % 2 === wanted) return;
	pages.push({
		...previous,
		index: pages.length,
		pageInSection: previous.pageInSection + 1,
		columns: previous.columns.map((column) => ({ ...column, blocks: [] })),
	});
}

/** Shifts each page's content down for centered or bottom-aligned sections (`w:vAlign`). */
function alignVertically(
	pages: LayoutPageBox[],
	section: LayoutSection,
	note: (m: string) => void,
) {
	const align = section.verticalAlign;
	if (!align || align === 'top') return;
	if (align === 'both') {
		note('Vertically justified sections are laid out top-aligned.');
		return;
	}
	for (const page of pages) {
		const available = page.heightPx - page.marginTopPx - page.marginBottomPx;
		const used = Math.max(
			0,
			...page.columns.flatMap((column) => column.blocks.map((block) => block.yPx + block.heightPx)),
		);
		const shift = (available - used) / (align === 'center' ? 2 : 1);
		if (shift <= 0) continue;
		for (const column of page.columns) for (const block of column.blocks) block.yPx += shift;
	}
}

/** Paginates an already-adapted engine input. See `layout.ts` for the DocumentModel-facing entry point. */
export function layoutSections(
	input: LayoutDocumentInput,
	measurer: TextMeasurer,
	exclusions?: ReadonlyMap<number, Exclusion[]>,
): LayoutResult {
	const approximations = new Set<string>();
	const note = (message: string) => approximations.add(message);
	const pages: LayoutPageBox[] = [];
	let previousFlow: SectionFlow | undefined;
	input.sections.forEach((section, index) => {
		insertParityBlank(pages, section);
		const first = pages.length;
		const previous = input.sections[index - 1];
		const page = pages.at(-1);
		let continueAt: number | undefined;
		if (
			section.break === 'continuous' &&
			previousFlow &&
			previous &&
			page &&
			page.widthPx === section.page.widthPx &&
			page.heightPx === section.page.heightPx
		) {
			if (page.footnotes?.length)
				note('Continuous sections sharing a page with footnotes start a new page.');
			else if (
				(previous.verticalAlign && previous.verticalAlign !== 'top') ||
				(section.verticalAlign && section.verticalAlign !== 'top')
			)
				note('Continuous sections with vertical alignment changes start a new page.');
			else if ((previous.columns?.count ?? 1) > 1 && !previousFlow.balanced)
				note(
					'Continuous breaks after multi-column sections requiring unsupported balancing start a new page; balancing those layouts (splittable or repeated-header tables, floats or explicit breaks) is not yet modeled.',
				);
			else continueAt = previousFlow.cursor.finishBand();
		}
		const next = input.sections[index + 1];
		const canBalance =
			(section.columns?.count ?? 1) > 1 &&
			next?.break === 'continuous' &&
			next.page.widthPx === section.page.widthPx &&
			next.page.heightPx === section.page.heightPx &&
			(!section.verticalAlign || section.verticalAlign === 'top') &&
			section.blocks.every(canBalanceBlock);
		const basePages = canBalance ? structuredClone(pages) : [];
		let flow = new SectionFlow(section, pages, measurer, note, index, exclusions, continueAt);
		flow.run();
		const last = pages.at(-1)!;
		if (canBalance && !last.footnotes?.length) {
			const balancedNotes = new WeakMap<SectionFlow, string[]>();
			const balanced = balanceColumns(
				basePages,
				last.index,
				flow.cursor.physicalColumnHeightPx,
				(trialPages, heightPx) => {
					const trialNotes: string[] = [];
					const trial = new SectionFlow(
						section,
						trialPages,
						measurer,
						(message) => trialNotes.push(message),
						index,
						exclusions,
						continueAt,
						{ pageIndex: last.index, heightPx },
					);
					trial.run();
					balancedNotes.set(trial, trialNotes);
					return trial;
				},
			);
			if (balanced.flow.balanceValid) {
				pages.splice(0, pages.length, ...balanced.pages);
				flow = balanced.flow;
				for (const message of balancedNotes.get(flow) ?? []) note(message);
				flow.balanced = true;
			}
		}
		previousFlow = flow;
		alignVertically(pages.slice(first), section, note);
	});
	return { pages, approximations: [...approximations] };
}
