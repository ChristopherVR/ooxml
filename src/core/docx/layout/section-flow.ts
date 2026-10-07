import { PageCursor } from './page-cursor';
import { placeParagraph } from './flow-paragraph';
import { placeTable } from './flow-table';
import { layoutParagraph, type ParagraphLayoutResult } from './paragraph-layout';
import { layoutRow, stackParagraphs } from './table-layout';
import { suppressesSpacing } from './keep-rules';
import type { TextMeasurer } from './measure';
import type { LayoutBlock, LayoutSection } from './input';
import type { LayoutPageBox } from './result';
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
	private readonly firstRowHeightCache = new Map<string, number>();

	constructor(
		section: LayoutSection,
		pages: LayoutPageBox[],
		measurer: TextMeasurer,
		note: (m: string) => void,
		sectionIndex = 0,
		private readonly exclusions?: ReadonlyMap<number, Exclusion[]>,
		continueAt?: number,
		private readonly balance?: { pageIndex: number; heightPx: number },
	) {
		this.blocks = section.blocks;
		this.measurer = measurer;
		this.note = note;
		this.cursor = new PageCursor(
			pages,
			section.page,
			section.columns ?? { count: 1, gapPx: 0 },
			sectionIndex,
			continueAt,
			balance,
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
				placeTable(this.cursor, block, this.measurer, this.note);
				if (this.balance?.pageIndex === this.cursor.page.index && !this.cursor.columnContentFits)
					this.balanceValid = false;
				continue;
			}
			if (block.afterTableSectionBreak) this.cursor.endBalance();
			const spacingBeforePx = this.spacingBeforeFor(index);
			const layout = this.wrappedLayout(index, spacingBeforePx) ?? this.paragraphLayout(index);
			if (block.footnotes?.length)
				this.cursor.holdFootnotes(
					block.id,
					block.footnotes.map((note) => {
						const stacked = stackParagraphs(
							note.paragraphs,
							this.cursor.columnWidthPx,
							this.measurer,
							this.note,
						);
						return { id: note.id, yPx: 0, heightPx: stacked.heightPx, paragraphs: stacked.boxes };
					}),
				);
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
					spacingBeforePx,
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
	}
}
