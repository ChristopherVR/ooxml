import type { ColumnInfo, Worksheet } from '../model.js';

/** Column properties without the span, for comparing neighbouring entries. */
function propsKey(info: ColumnInfo): string {
	const { min: _min, max: _max, ...rest } = info;
	return JSON.stringify(
		Object.entries(rest)
			.filter(([, v]) => v !== undefined)
			.sort(),
	);
}

const hasProps = (info: ColumnInfo): boolean => propsKey(info) !== '[]';

/** The column entry covering `col`, if any. */
export const columnAt = (sheet: Worksheet, col: number): ColumnInfo | undefined =>
	sheet.columns.find((c) => col >= c.min && col <= c.max);

/**
 * Edits the properties of individual columns. Spans covering them are split, each target column
 * gets its own entry for `edit`, and equal neighbours are joined again afterwards.
 */
export function editColumns(
	sheet: Worksheet,
	cols: number[],
	edit: (info: ColumnInfo) => void,
): void {
	const targets = new Set(cols);
	const singles = new Map<number, ColumnInfo>();
	const kept: ColumnInfo[] = [];
	for (const info of sheet.columns) {
		let runStart = info.min;
		for (let c = info.min; c <= info.max + 1; c++) {
			const isTarget = c <= info.max && targets.has(c);
			if (c > info.max || isTarget) {
				if (c > runStart) kept.push({ ...info, min: runStart, max: c - 1 });
				runStart = c + 1;
			}
			if (isTarget) singles.set(c, { ...info, min: c, max: c });
		}
	}
	for (const c of targets) {
		const info = singles.get(c) ?? { min: c, max: c };
		edit(info);
		singles.set(c, info);
	}
	sheet.columns = normalizeColumns([...kept, ...singles.values()]);
}

/** Sorts entries, drops empty ones and joins adjacent entries with equal properties. */
export function normalizeColumns(columns: ColumnInfo[]): ColumnInfo[] {
	const sorted = columns.filter(hasProps).sort((a, b) => a.min - b.min);
	const out: ColumnInfo[] = [];
	for (const info of sorted) {
		const last = out.at(-1);
		if (last && last.max + 1 === info.min && propsKey(last) === propsKey(info)) last.max = info.max;
		else out.push({ ...info });
	}
	return out;
}
