import type { LayoutTableRow } from './input';
import type { NestedFrameProfile } from './nested-frame-profile';
import type { LayoutParagraphBox } from './result';
import type { TextMeasurer } from './measure';
import { layoutRow, type RowLayout } from './table-layout';
/** Native row heights include the top edge; the physical fragment also reserves its last bottom. */
export function layoutBorderedRow(
	row: LayoutTableRow,
	width: number,
	measurer: TextMeasurer,
	note: (message: string) => void,
	profile: NestedFrameProfile | undefined,
	remaining?: LayoutParagraphBox[][],
): RowLayout {
	const reserved =
		profile !== undefined && (profile.leftPx > 0 || profile.rightPx > 0 || profile.bottomPx > 0);
	const layout = layoutRow(
		row,
		width,
		measurer,
		note,
		remaining,
		undefined,
		reserved,
		profile?.leftPx,
	);
	if (!reserved) return layout;
	const bottom = Math.max(0, ...layout.geometry.map((cell) => cell.borders?.bottom?.widthPx ?? 0));
	return {
		...layout,
		heightPx: layout.heightPx + bottom,
		borderBottomPx: bottom,
		cellPadding: row.cells.map((cell) => ({
			top: (cell.padding?.top ?? 0) + (cell.borders?.top?.widthPx ?? 0),
			bottom: (cell.padding?.bottom ?? 0) + (cell.borders?.bottom?.widthPx ?? 0),
		})),
	};
}

/** Share a horizontal stroke with the next row, retaining it when a fragment is flushed. */
export class TableFragmentBorders {
	private bottom = 0;
	private included = true;
	record(bottom: number | undefined): void {
		this.bottom = bottom ?? 0;
		this.included = true;
	}
	remove(row: import('./result').LayoutTableRowBox | undefined): number {
		if (!row || !this.included) return 0;
		row.heightPx -= this.bottom;
		this.included = false;
		return -this.bottom;
	}
	restore(row: import('./result').LayoutTableRowBox | undefined): number {
		if (!row || this.included) return 0;
		row.heightPx += this.bottom;
		this.included = true;
		return this.bottom;
	}
}
