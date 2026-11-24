// Shares a table width between AutoFit columns (see table-autofit.ts for the measured rules).

export interface AutofitColumn {
	/** Content minimum and maximum, cell margins included (CSS pixels, whole twips). */
	minPx: number;
	maxPx: number;
	/** The largest preferred width (`w:tcW` dxa) of the column's cells. */
	preferredPx?: number;
	/** The largest preferred percentage (`w:tcW` pct) as a fraction of the table width. */
	preferredFraction?: number;
}

export interface AutofitTarget {
	/** Room in the text column. */
	availablePx: number;
	/** The table's preferred width (`w:tblW` dxa or pct), when it has one. */
	targetPx?: number;
}

const toTwips = (px: number) => Math.round(px * 15);

/** Moves each column from `from` towards `to` by the same fraction so that they add up to `total`. */
function interpolate(from: number[], to: number[], total: number): number[] {
	const low = from.reduce((sum, value) => sum + value, 0);
	const high = to.reduce((sum, value) => sum + value, 0);
	const k = high > low ? Math.min(1, Math.max(0, (total - low) / (high - low))) : 0;
	return from.map((value, index) => value + (to[index]! - value) * k);
}

/** Whole twips that keep the rounded total (largest remainders take the spare twips). */
function roundToTotal(widths: number[]): number[] {
	const total = Math.round(widths.reduce((sum, value) => sum + value, 0));
	const floors = widths.map(Math.floor);
	let spare = total - floors.reduce((sum, value) => sum + value, 0);
	const order = widths
		.map((value, index) => ({ index, rest: value - Math.floor(value) }))
		.sort((a, b) => b.rest - a.rest);
	for (const { index } of order) {
		if (spare <= 0) break;
		floors[index]! += 1;
		spare -= 1;
	}
	return floors;
}

/**
 * Caps the widest columns at one common width so that they add up to `total`: narrow columns keep
 * their minimum and only the columns that need more than their share give way.
 */
export function capToTotal(widths: number[], total: number): number[] {
	if (!widths.length || total <= 0) return widths.map(() => 0);
	const sorted = [...widths].sort((a, b) => a - b);
	let rest = total;
	let cap = total / widths.length;
	for (const [index, width] of sorted.entries()) {
		cap = rest / (sorted.length - index);
		if (width > cap) break;
		rest -= width;
	}
	return widths.map((width) => Math.min(width, cap));
}

/** Column widths in whole twips. */
export function distributeColumns(columns: AutofitColumn[], target: AutofitTarget): number[] {
	if (!columns.length) return [];
	const available = toTwips(target.availablePx);
	const tableWidth = target.targetPx !== undefined ? toTwips(target.targetPx) : undefined;
	const min = columns.map((column) => toTwips(column.minPx));
	const contentMax = columns.map((column, index) => Math.max(min[index]!, toTwips(column.maxPx)));
	// A percentage column takes its share of the table width; an automatic-width table is as wide
	// as the other columns at their maximum need, within the text column.
	const fractions = columns.map((column) =>
		column.preferredPx === undefined ? (column.preferredFraction ?? 0) : 0,
	);
	const fractionSum = Math.min(
		0.99,
		fractions.reduce((sum, value) => sum + value, 0),
	);
	const percentBase =
		tableWidth ??
		Math.min(
			available,
			columns.reduce(
				(sum, column, index) =>
					fractions[index]
						? sum
						: sum + Math.max(contentMax[index]!, toTwips(column.preferredPx ?? 0)),
				0,
			) /
				(1 - fractionSum),
		);
	const preferred = columns.map((column, index) =>
		column.preferredPx !== undefined
			? toTwips(column.preferredPx)
			: fractions[index]
				? fractions[index]! * percentBase
				: undefined,
	);
	const fixed = preferred.map((value) => value !== undefined);
	const max = contentMax.map((value, index) =>
		fixed[index] ? Math.max(min[index]!, preferred[index]!) : value,
	);
	const sum = (values: number[]) => values.reduce((acc, value) => acc + value, 0);
	const limit = tableWidth ?? available;
	// A preferred table width below the content minimum gives way to the minimum (Word's grids),
	// but the minimum itself stays inside the text column: an unbreakable token (a URL, an IBAN)
	// then breaks inside its cell instead of pushing the table off the page (#320).
	if (sum(min) >= limit)
		return roundToTotal(sum(min) > available ? capToTotal(min, available) : min);
	if (sum(max) <= limit) {
		if (tableWidth === undefined) return roundToTotal(max);
		// A table wider than its content grows the automatic columns in proportion to their
		// maximum (every column, when all have a preferred width).
		const grow = fixed.every(Boolean) ? fixed.map(() => true) : fixed.map((value) => !value);
		const growing = sum(max.filter((_, index) => grow[index]));
		const extra = tableWidth - sum(max);
		return roundToTotal(
			max.map((value, index) =>
				grow[index] ? value + (growing > 0 ? (extra * value) / growing : 0) : value,
			),
		);
	}
	// Shrink the automatic columns first; preferred widths give way only when those are at minimum.
	const pick = (values: number[], wanted: boolean) =>
		values.map((value, index) => (fixed[index] === wanted ? value : 0));
	const fixedMax = sum(pick(max, true));
	const autoMin = sum(pick(min, false));
	const widths =
		fixedMax + autoMin <= limit
			? interpolate(pick(min, false), pick(max, false), limit - fixedMax).map((value, index) =>
					fixed[index] ? max[index]! : value,
				)
			: interpolate(pick(min, true), pick(max, true), limit - autoMin).map((value, index) =>
					fixed[index] ? value : min[index]!,
				);
	return roundToTotal(widths);
}
