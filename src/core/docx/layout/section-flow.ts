import { createReferenceFootnotes, type ReferenceFootnotes } from './reference-footnotes';
import { cellParagraphs } from './table-cell-layout';
import { NO_TWIPS } from './units';
import { PageCursor } from './page-cursor';
import { placeParagraph } from './flow-paragraph';
import { placeTable } from './flow-table';
import { resolveTableAutofit } from './table-autofit';
import { layoutParagraph, type ParagraphLayoutResult } from './paragraph-layout';
import { layoutRow, stackParagraphs } from './table-layout';
import { suppressesSpacing } from './keep-rules';
import type { TextMeasurer } from './measure';
import type { LayoutBlock, LayoutSection, LayoutNote, StoryMargins } from './input';
import type { LayoutFootnoteBox, LayoutLine, LayoutPageBox } from './result';
import { lineBoxesFor, type Exclusion } from './wrap';
import { expectDefined } from '../expect-defined';
import { tokenPosition, type TokenPosition } from './text-breaks';

/** Resumes a section part-way and observes each placed block (incremental layout). */
export interface SectionFlowResume {
	/** First block to place; earlier blocks are already on the pages handed to the flow. */
	startIndex: number;
	/** `pageInSection` of the first new page. */
	pageInSection: number;
	/** Called after each block is placed; returning true stops the flow. */
	afterBlock(index: number, cursor: PageCursor): boolean;
}

export class SectionFlow {
	readonly cursor: PageCursor;
	balanceValid = true;
	balanced = false;
	private readonly blocks: LayoutBlock[];
	private readonly measurer: TextMeasurer;
	private readonly note: (message: string) => void;
	private readonly paragraphCache = new Map<string, ParagraphLayoutResult>();
	private readonly referenceFootnotes: ReferenceFootnotes;
	private readonly firstRowHeightCache = new Map<string, number>();

	constructor(
		private readonly section: LayoutSection,
		pages: LayoutPageBox[],
		measurer: TextMeasurer,
		note: (m: string) => void,
		sectionIndex = 0,
		private readonly exclusions?: ReadonlyMap<number, Exclusion[]>,
		continueAt?: number,
		private readonly balance?: { pageIndex: number; heightPx: number },
		private readonly resume?: SectionFlowResume,
		storyMargins?: StoryMargins,
	) {
		const noteColumns =
			section.footnoteColumns ??
			(section.modernFootnoteSeparator ? (section.columns ?? { count: 1, gapPx: 0 }) : undefined);
		if (
			!noteColumns &&
			!section.legacyFootnoteColumns &&
			(section.columns?.count ?? 1) > 1 &&
			section.blocks.some((block) =>
				block.kind === 'paragraph'
					? !!block.footnotes?.length
					: block.rows.some((row) =>
							row.cells.some((cell) =>
								cellParagraphs(cell).some((paragraph) => !!paragraph.footnotes?.length),
							),
						),
			)
		)
			note(
				'Multi-column footnotes use a page-wide area; native per-column footnote balancing is not yet modeled.',
			);
		const sources = new Map<string, LayoutNote>();
		for (const block of section.blocks) {
			const paragraphs =
				block.kind === 'paragraph'
					? [block]
					: block.rows.flatMap((row) => row.cells.flatMap(cellParagraphs));
			for (const paragraph of paragraphs)
				for (const source of paragraph.footnotes ?? []) sources.set(source.id, source);
		}
		const reflowCache = new Map<string, LayoutFootnoteBox>();
		const reflow = (box: LayoutFootnoteBox, width: number): LayoutFootnoteBox => {
			const source = sources.get(box.id);
			if (!source) return box;
			const ids = new Set(box.paragraphs.map((paragraph) => paragraph.blockId));
			const key = JSON.stringify([
				box.id,
				width,
				[...ids],
				box.sourceTokenStarts,
				box.sourceTokenOffsets,
			]);
			const cached = reflowCache.get(key);
			if (cached) return { ...cached, continued: !!box.continued };
			const paragraphs = source.paragraphs
				.filter((paragraph) => ids.has(paragraph.id))
				.map((paragraph) =>
					box.sourceTokenStarts?.[paragraph.id] || box.sourceTokenOffsets?.[paragraph.id]
						? { ...paragraph, spacingBeforeTwips: NO_TWIPS }
						: paragraph,
				);
			const stacked = stackParagraphs(
				paragraphs,
				width,
				measurer,
				note,
				0,
				new Map(
					Object.entries(box.sourceTokenStarts ?? {}).map(([id, index]) => [
						id,
						tokenPosition(index, box.sourceTokenOffsets?.[id]),
					]),
				),
			);
			const result = { ...box, yPx: 0, heightPx: stacked.heightPx, paragraphs: stacked.boxes };
			reflowCache.set(key, result);
			return result;
		};
		this.blocks = section.blocks;
		this.measurer = measurer;
		this.note = note;
		const separatorCache = new Map<string, NonNullable<LayoutPageBox['footnoteSeparator']>>();
		this.cursor = new PageCursor(
			pages,
			section.page,
			section.columns ?? { count: 1, gapPx: 0 },
			sectionIndex,
			continueAt,
			balance,
			note,
			(page, continued) => {
				const paragraphs =
					section.footnoteSeparators?.[
						continued && !section.modernFootnoteSeparator ? 'continuationSeparator' : 'separator'
					];
				if (!paragraphs) return undefined;
				const width = page.widthPx - page.marginLeftPx - page.marginRightPx;
				const key = `${width}:${continued}`;
				const cached = separatorCache.get(key);
				if (cached) return cached;
				const adapted = paragraphs.map((paragraph) => {
					const probe = section.legacyFootnoteColumns
						? layoutParagraph(
								{
									...paragraph,
									runs: paragraph.runs.map((run) =>
										run.object?.noteSeparator
											? { ...run, object: { ...run.object, widthPx: 0 } }
											: run,
									),
								},
								width,
								measurer,
								note,
							)
						: undefined;
					return {
						...paragraph,
						runs: paragraph.runs.map((run, index) =>
							run.object?.noteSeparator
								? {
										...run,
										object: {
											...run.object,
											widthPx: section.legacyFootnoteColumns
												? Math.min(
														run.object.noteSeparator === 'continuationSeparator' ? width : 192,
														Math.max(
															0,
															width -
																(probe?.lines
																	.flatMap((line) => line.fragments)
																	.find(
																		(fragment) =>
																			fragment.runIndex === index && fragment.object?.noteSeparator,
																	)?.xPx ?? 0),
														),
													)
												: run.object.noteSeparator === 'continuationSeparator'
													? width
													: Math.min(192, width),
										},
									}
								: run,
						),
					};
				});
				const stacked = stackParagraphs(adapted, width, measurer, note);
				const result = { heightPx: stacked.heightPx, paragraphs: stacked.boxes };
				separatorCache.set(key, result);
				return result;
			},
			noteColumns,
			reflow,
			section.legacyFootnoteColumns,
			resume?.pageInSection ?? 0,
			storyMargins,
		);
		this.referenceFootnotes = createReferenceFootnotes(
			() => this.cursor.footnoteWidthPx,
			measurer,
			note,
		);
	}

	private paragraphLayout(index: number): ParagraphLayoutResult {
		const key = `${index}:${this.cursor.columnWidthPx}`;
		let cached = this.paragraphCache.get(key);
		if (!cached) {
			cached = layoutParagraph(
				this.blocks[index] as Extract<LayoutBlock, { kind: 'paragraph' }>,
				this.cursor.columnWidthPx,
				this.measurer,
				this.note,
			);
			this.paragraphCache.set(key, cached);
		}
		return cached;
	}

	private firstRowHeight(index: number): number {
		const key = `${index}:${this.cursor.columnWidthPx}`;
		let cached = this.firstRowHeightCache.get(key);
		if (cached === undefined) {
			const table = resolveTableAutofit(
				this.blocks[index] as Extract<LayoutBlock, { kind: 'table' }>,
				this.cursor.columnWidthPx,
				this.measurer,
			);
			const [firstRow] = table.rows;
			cached = firstRow
				? layoutRow(firstRow, this.cursor.columnWidthPx, this.measurer, this.note).heightPx
				: 0;
			this.firstRowHeightCache.set(key, cached);
		}
		return cached;
	}

	private spacingBeforeFor(index: number): number {
		const block = this.blocks[index];
		if (block?.kind !== 'paragraph') return 0;
		const layout = this.paragraphLayout(index);
		const previous = index > 0 ? this.blocks[index - 1] : undefined;
		if (index === 0 && !this.section.addParagraphSpacing)
			return Math.max(0, layout.spacingBeforePx - (this.section.leadingSpacingAfterPx ?? 0));
		if (previous?.kind !== 'paragraph') return layout.spacingBeforePx;
		if (suppressesSpacing(previous, block)) return 0;
		// Unless the document sets doNotUseHTMLParagraphAutoSpacing, Word leaves the larger of the space
		// after the first paragraph and the space before the second (measured, spacing-probe.json), not
		// their sum: the first already took its own. Tables and sections are handled by their callers.
		if (this.section.addParagraphSpacing) return layout.spacingBeforePx;
		return Math.max(0, layout.spacingBeforePx - this.paragraphLayout(index - 1).spacingAfterPx);
	}

	/** Contextual spacing: no space after a paragraph that is followed by one of its own style. */
	private spacingAfterSuppressed(index: number): boolean {
		const block = this.blocks[index];
		const next = this.blocks[index + 1];
		return (
			block?.kind === 'paragraph' && next?.kind === 'paragraph' && suppressesSpacing(block, next)
		);
	}

	/** Minimum height that must fit before `index` may start, walking a `keepNext` chain. */
	private requiredKeepHeight(index: number, depth = 0): number {
		if (index >= this.blocks.length || depth > 64) return 0;
		const block = expectDefined(this.blocks[index], 'keep-chain block index');
		// A floating table takes no room from the flow.
		if (block.kind === 'table') return block.floating ? 0 : this.firstRowHeight(index);
		const layout = this.paragraphLayout(index);
		const before = this.spacingBeforeFor(index);
		const own =
			before +
			(block.keepLines
				? layout.contentHeightPx + layout.spacingAfterPx
				: (layout.lines[0]?.heightPx ?? 0));
		return block.keepNext ? own + this.requiredKeepHeight(index + 1, depth + 1) : own;
	}

	/**
	 * Lay the remaining paragraph around the wrapped objects on the current page.
	 */
	private layoutAtCursor(
		index: number,
		spacingBeforePx: number,
		start: TokenPosition = 0,
	): ParagraphLayoutResult {
		const page = this.cursor.page;
		const exclusions = this.exclusions?.get(page.index);
		const block = this.blocks[index] as Extract<LayoutBlock, { kind: 'paragraph' }>;
		const width = this.cursor.columnWidthPx;
		const plain = start
			? layoutParagraph(block, width, this.measurer, this.note, undefined, start)
			: this.paragraphLayout(index);
		if (!exclusions) return plain;
		const topPx = page.marginTopPx + this.cursor.y + spacingBeforePx;
		const bottomPx = topPx + plain.contentHeightPx;
		if (!exclusions.some((item) => item.yPx < bottomPx && item.yPx + item.heightPx > topPx))
			return plain;
		return layoutParagraph(
			block,
			width,
			this.measurer,
			this.note,
			lineBoxesFor(exclusions, topPx, page.marginLeftPx + this.cursor.column.xPx, width),
			start,
		);
	}

	run(): void {
		for (let index = this.resume?.startIndex ?? 0; index < this.blocks.length; index++) {
			const block = expectDefined(this.blocks[index], 'flow block index');
			if (block.kind === 'table') {
				placeTable(
					this.cursor,
					block,
					this.measurer,
					this.note,
					block.rows.some((row) =>
						row.cells.some((cell) =>
							cellParagraphs(cell).some((paragraph) => paragraph.footnotes?.length),
						),
					)
						? this.referenceFootnotes
						: undefined,
				);
				if (this.balance?.pageIndex === this.cursor.page.index && !this.cursor.columnContentFits)
					this.balanceValid = false;
				if (this.resume?.afterBlock(index, this.cursor)) return;
				continue;
			}
			if (block.afterTableSectionBreak) this.cursor.endBalance();
			const spacingBeforePx = this.spacingBeforeFor(index);
			const layout = this.layoutAtCursor(
				index,
				this.cursor.atContinuationTop ? 0 : spacingBeforePx,
			);
			const footnotes = (lines: readonly LayoutLine[], includeUnanchored: boolean) =>
				this.referenceFootnotes(block, lines, includeUnanchored);
			const requiredTogetherPx = block.keepNext
				? this.requiredKeepHeight(index)
				: block.keepLines
					? spacingBeforePx + layout.contentHeightPx + layout.spacingAfterPx
					: 0;
			if (
				requiredTogetherPx > this.cursor.columnHeightPx &&
				requiredTogetherPx <= this.cursor.physicalColumnHeightPx
			)
				this.balanceValid = false;
			placeParagraph(
				this.cursor,
				{
					paragraph: block,
					layout,
					footnotes,
					spacingBeforePx,
					...(this.section.keepSpaceBeforePageBreak ? { keepSpaceBeforePageBreak: true } : {}),
					...(this.spacingAfterSuppressed(index) ? { spacingAfterPx: 0 } : {}),
					...(this.section.modernFootnoteSeparator || this.section.legacyFootnoteColumns
						? { trimLeadingColumnBreak: true }
						: {}),
					requiredTogetherPx,
					reflow: (token, before) => this.layoutAtCursor(index, before, token),
					...(this.exclusions
						? { hasExclusions: (pageIndex) => this.exclusions!.has(pageIndex) }
						: {}),
				},
				this.note,
			);
			if (
				this.balance?.pageIndex === this.cursor.page.index &&
				this.cursor.usedColumnHeightPx > this.cursor.columnHeightPx
			)
				this.balanceValid = false;
			if (this.resume?.afterBlock(index, this.cursor)) return;
		}
		this.cursor.finishFootnotes();
	}
}
