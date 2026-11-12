// What the Format Cells dialog shows for a selection, as data: each aspect holds the properties
// every examined cell agrees on, and names the ones that differ (mixed).
import {
	normalizeRange,
	styleAt,
	type Alignment,
	type Border,
	type CellRange,
	type CellStyle,
	type Fill,
	type Font,
	type Protection,
	type Workbook,
} from '../index';
import { effectiveStyleId } from '../layout/cell-view';

/** The properties all examined cells share, and the names of those that differ. */
export interface FormatAspect<T> {
	/** Properties with one value across the selection (a property nobody sets is absent). */
	value: T;
	/** Property names whose value differs between cells. */
	mixed: string[];
}

export interface FormatCellsState {
	font: FormatAspect<Font>;
	alignment: FormatAspect<Alignment>;
	border: FormatAspect<Border>;
	protection: FormatAspect<Protection>;
	/** The number format code, or null when the selection mixes formats. */
	numFmt: string | null;
	/** The fill, or null when the selection mixes fills. */
	fill: Fill | null;
	/** False when the selection was too large to examine fully (its mixed lists are then partial). */
	complete: boolean;
}

/** Cells examined per request; a larger selection reports `complete: false`. */
export const FORMAT_CELLS_SAMPLE = 4096;

function aspect<T extends object>(items: readonly T[]): FormatAspect<T> {
	const value: Record<string, unknown> = {};
	const mixed: string[] = [];
	const keys = new Set(items.flatMap((item) => Object.keys(item)));
	for (const key of keys) {
		const seen = new Set(
			items.map((item) => JSON.stringify((item as Record<string, unknown>)[key])),
		);
		if (seen.size > 1) mixed.push(key);
		else {
			const first = (items[0] as Record<string, unknown> | undefined)?.[key];
			if (first !== undefined) value[key] = first;
		}
	}
	return { value: value as T, mixed };
}

/** The common and mixed formatting of `ranges` on one sheet. */
export function formatCellsState(
	workbook: Workbook | undefined,
	sheetIndex: number,
	ranges: readonly CellRange[],
): FormatCellsState {
	const sheet = workbook?.sheets[sheetIndex];
	const ids = new Set<number>();
	let remaining = FORMAT_CELLS_SAMPLE;
	let complete = true;
	if (workbook && sheet)
		outer: for (const raw of ranges) {
			const range = normalizeRange(raw);
			for (let row = range.start.row; row <= range.end.row; row++)
				for (let col = range.start.col; col <= range.end.col; col++) {
					if (--remaining < 0) {
						complete = false;
						break outer;
					}
					ids.add(effectiveStyleId(sheet, row, col) ?? 0);
				}
		}
	const styles: CellStyle[] = workbook ? [...ids].map((id) => styleAt(workbook, id)) : [];
	const numFmts = new Set(styles.map((s) => s.numFmt));
	const fills = new Set(styles.map((s) => JSON.stringify(s.fill)));
	return {
		font: aspect(styles.map((s) => s.font)),
		alignment: aspect(styles.map((s) => s.alignment ?? {})),
		border: aspect(styles.map((s) => s.border)),
		protection: aspect(styles.map((s) => s.protection ?? {})),
		numFmt: numFmts.size === 1 ? [...numFmts][0]! : null,
		fill: fills.size === 1 && styles[0] ? styles[0].fill : null,
		complete,
	};
}
