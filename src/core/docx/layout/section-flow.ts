import { createReferenceFootnotes, type ReferenceFootnotes } from './reference-footnotes';
import { NO_TWIPS } from './units';
import { PageCursor } from './page-cursor';
import { placeParagraph } from './flow-paragraph';
import { placeTable } from './flow-table';
import { layoutParagraph, type ParagraphLayoutResult } from './paragraph-layout';
import { layoutRow, stackParagraphs } from './table-layout';
import { suppressesSpacing } from './keep-rules';
import type { TextMeasurer } from './measure';
import type { LayoutBlock, LayoutSection, LayoutNote } from './input';
import type { LayoutFootnoteBox, LayoutLine, LayoutPageBox } from './result';
import { lineBoxesFor, type Exclusion } from './wrap';
import { expectDefined } from '../expect-defined';

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
								cell.paragraphs.some((paragraph) => !!paragraph.footnotes?.length),
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
					: block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs));
			for (const paragraph of paragraphs)
				for (const source of paragraph.footnotes ?? []) sources.set(source.id, source);
		}
		const reflowCache = new Map<string, LayoutFootnoteBox>();
		const reflow = (box: LayoutFootnoteBox, width: number): LayoutFootnoteBox => {
			const source = sources.get(box.id);
			if (!source) return box;
			const ids = new Set(box.paragraphs.map((paragraph) => paragraph.blockId));
			const key = JSON.stringify([box.id, width, [...ids], box.sourceTokenStarts]);
			const cached = reflowCache.get(key);
			if (cached) return { ...cached, continued: !!box.continued };
			const paragraphs = source.paragraphs
				.filter((paragraph) => ids.has(paragraph.id))
				.map((paragraph) =>
					box.sourceTokenStarts?.[paragraph.id]
						? { ...paragraph, spacingBeforeTwips: NO_TWIPS }
						: paragraph,
				);
			const stacked = stackParagraphs(
				paragraphs,
				width,
				measurer,
				note,
				0,
				new Map(Object.entries(box.sourceTokenStarts ?? {})),
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
			const table = this.blocks[index] as Extract<LayoutBlock, { kind: 'table' }>;
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
		if (previous?.kind === 'paragraph' && suppressesSpacing(previous, block)) return 0;
		return layout.spacingBeforePx;
	}

	/** Minimum height that must fit before `index` may start, walking a `keepNext` chain. */
	private requiredKeepHeight(index: number, depth = 0): number {
		if (index >= this.blocks.length || depth > 64) return 0;
		const block = expectDefined(this.blocks[index], 'keep-chain block index');
		if (block.kind === 'table') return this.firstRowHeight(index);
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
	 * The paragraph laid out around wrapped pictures on the current page, when any of them reaches
	 * the paragraph's lines; lines continuing onto a later page keep these widths.
	 */
	private wrappedLayout(index: number, spacingBeforePx: number): ParagraphLayoutResult | undefined {
		const page = this.cursor.page;
		const exclusions = this.exclusions?.get(page.index);
		if (!exclusions) return undefined;
		const plain = this.paragraphLayout(index);
		const topPx = page.marginTopPx + this.cursor.y + spacingBeforePx;
		const bottomPx = topPx + plain.contentHeightPx;
		if (!exclusions.some((item) => item.yPx < bottomPx && item.yPx + item.heightPx > topPx))
			return undefined;
		return layoutParagraph(
			this.blocks[index] as Extract<LayoutBlock, { kind: 'paragraph' }>,
			this.cursor.columnWidthPx,
			this.measurer,
			this.note,
			lineBoxesFor(
				exclusions,
				topPx,
				page.marginLeftPx + this.cursor.column.xPx,
				this.cursor.columnWidthPx,
			),
		);
	}

	run(): void {
		for (let index = 0; index < this.blocks.length; index++) {
			const block = expectDefined(this.blocks[index], 'flow block index');
			if (block.kind === 'table') {
				placeTable(
					this.cursor,
					block,
					this.measurer,
					this.note,
					block.rows.some((row) =>
						row.cells.some((cell) =>
							cell.paragraphs.some((paragraph) => paragraph.footnotes?.length),
						),
					)
						? this.referenceFootnotes
						: undefined,
				);
				if (this.balance?.pageIndex === this.cursor.page.index && !this.cursor.columnContentFits)
					this.balanceValid = false;
				continue;
			}
			if (block.afterTableSectionBreak) this.cursor.endBalance();
			const spacingBeforePx = this.spacingBeforeFor(index);
			const layout = this.wrappedLayout(index, spacingBeforePx) ?? this.paragraphLayout(index);
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
					...(this.section.modernFootnoteSeparator || this.section.legacyFootnoteColumns
						? { trimLeadingColumnBreak: true }
						: {}),
					requiredTogetherPx,
					reflow: (token) =>
						layoutParagraph(
							block,
							this.cursor.columnWidthPx,
							this.measurer,
							this.note,
							undefined,
							token,
						),
				},
				this.note,
			);
			if (
				this.balance?.pageIndex === this.cursor.page.index &&
				this.cursor.usedColumnHeightPx > this.cursor.columnHeightPx
			)
				this.balanceValid = false;
		}
		this.cursor.finishFootnotes();
	}
}
