import type { LayoutTableRow } from './input';
import type { RowLayout } from './table-layout';

/**
 * Heights of rows that a vertical merge spans. Word sizes every row from its own cells, then lets
 * the text of a merged cell that is taller than the rows it covers grow the LAST spanned row
 * (measured in `word-table.test.ts`). A merge that would not fit the space left in the column
 * keeps the older behaviour: its text counts in the first row, which the row splitter can cut.
 */
export class MergedRowPlan {
	private readonly legacy = new Set<LayoutTableRow>();
	private readonly minimum = new Map<LayoutTableRow, number>();
	private readonly cache = new Map<LayoutTableRow, LayoutTableRow>();

	constructor(private readonly rows: readonly LayoutTableRow[]) {}

	private apply(source: LayoutTableRow): LayoutTableRow {
		const minimum = this.minimum.get(source);
		if (minimum === undefined && !this.legacy.has(source)) return source;
		return {
			...source,
			...(minimum === undefined ? {} : { heightPx: minimum, heightRule: 'atLeast' as const }),
			...(this.legacy.has(source) ? { mergedTextInRow: true } : {}),
		};
	}

	/** The row to lay out and place (the source row itself when the plan leaves it alone). */
	effective(source: LayoutTableRow): LayoutTableRow {
		let row = this.cache.get(source);
		if (!row) {
			row = this.apply(source);
			this.cache.set(source, row);
		}
		return row;
	}

	/** Plans the rows a merge starting at `index` covers, given the room left for them. */
	plan(index: number, layout: (row: LayoutTableRow) => RowLayout, availablePx: number): void {
		const start = this.rows[index];
		const spans = start?.cells.map((cell) => cell.rowSpan ?? 1) ?? [];
		const longest = Math.max(1, ...spans);
		if (!start || longest < 2) return;
		const covered = this.rows.slice(index, index + longest);
		const natural = covered.map((row) => layout(this.apply(row)));
		const heights = natural.map((item) => item.heightPx);
		spans.forEach((span, cell) => {
			if (span < 2) return;
			const count = Math.min(span, covered.length);
			// The merged cell has no bottom edge of its own: the last spanned row's border closes it.
			const need =
				(natural[0]?.geometry[cell]?.contentHeightPx ?? 0) +
				(natural[count - 1]?.borderBottomPx ?? 0);
			const have = heights.slice(0, count).reduce((sum, h) => sum + h, 0);
			if (need > have) heights[count - 1] = (heights[count - 1] ?? 0) + need - have;
		});
		if (heights.reduce((sum, h) => sum + h, 0) > availablePx) {
			this.legacy.add(start);
			return;
		}
		covered.forEach((row, at) => {
			const grown = heights[at]! - natural[at]!.heightPx;
			if (grown > 1e-6 && row.heightRule !== 'exact')
				this.minimum.set(row, heights[at]! - (natural[at]!.borderBottomPx ?? 0));
		});
	}
}
