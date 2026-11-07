/**
 * A table cell's edge lines from `TableCellInput.borders`, mapped onto the
 * per-edge {@link PptxTableCellStyle} fields the save path writes as
 * `a:lnL` / `a:lnR` / `a:lnT` / `a:lnB`. A px dash pattern becomes
 * `a:custDash` segments the way a shape outline's does.
 *
 * @module sdk/table-border-input
 */

import type { PptxCustomDashSegment } from '../../types/shape-style';
import type { PptxTableCellStyle } from '../../types/table';
import { customDashSegments } from './shape-style-input';
import type { TableCellBorderInput, TableCellBordersInput } from './types';

/** `a:ds/@d` and `@sp` are `ST_PositivePercentage` values, 32-bit integers. */
const MAX_DASH_UNITS = 2147483647;

const EDGES = [
	['top', 'borderTop'],
	['right', 'borderRight'],
	['bottom', 'borderBottom'],
	['left', 'borderLeft'],
] as const;

type EdgePrefix = (typeof EDGES)[number][1];

/**
 * A px dash pattern as `a:ds` segments relative to the line width.
 *
 * @throws RangeError for an empty pattern, a length that is not a finite
 *   number above 0, or one too short or too long to write.
 */
function borderDashSegments(
	edge: string,
	lengths: readonly number[],
	width: number | undefined,
): PptxCustomDashSegment[] {
	const name = `TableCellInput.borders.${edge}.dash`;
	if (lengths.length === 0) {
		throw new RangeError(`${name} must give at least one length`);
	}
	if (lengths.some((n) => !Number.isFinite(n) || n <= 0)) {
		throw new RangeError(`${name} lengths must be finite and above 0, got [${lengths.join(', ')}]`);
	}
	const segments = customDashSegments([...lengths], width !== undefined && width > 0 ? width : 1);
	if (segments.some((s) => s.dash <= 0 || s.space <= 0)) {
		throw new RangeError(`${name} has a length too short to write: [${lengths.join(', ')}]`);
	}
	if (segments.some((s) => s.dash > MAX_DASH_UNITS || s.space > MAX_DASH_UNITS)) {
		throw new RangeError(`${name} has a length too long to write: [${lengths.join(', ')}]`);
	}
	return segments;
}

/** One edge as cell style fields, keyed by its field prefix. */
function edgeStyle(
	edge: string,
	prefix: EdgePrefix,
	input: TableCellBorderInput,
): Partial<PptxTableCellStyle> {
	if (input.width !== undefined && (!Number.isFinite(input.width) || input.width < 0)) {
		throw new RangeError(
			`TableCellInput.borders.${edge}.width must be a finite width of 0 or more, got ${input.width}`,
		);
	}
	const out: Record<string, unknown> = {};
	if (input.width !== undefined) out[`${prefix}Width`] = input.width;
	if (input.color !== undefined) out[`${prefix}Color`] = input.color;
	if (Array.isArray(input.dash)) {
		out[`${prefix}Dash`] = 'custom';
		out[`${prefix}DashSegments`] = borderDashSegments(edge, input.dash, input.width);
	} else if (input.dash !== undefined) {
		if ((input.dash as string) === 'custom') {
			throw new RangeError(`TableCellInput.borders.${edge}.dash: give 'custom' as a length list`);
		}
		out[`${prefix}Dash`] = input.dash;
	}
	return out as Partial<PptxTableCellStyle>;
}

/**
 * The cell style fields of a cell's edge lines: `all` for every edge, an
 * edge's own input over it.
 *
 * @throws RangeError for a width that is negative or not finite, or a dash
 *   pattern {@link borderDashSegments} refuses.
 */
export function tableCellBorderStyle(
	borders: TableCellBordersInput | undefined,
): Partial<PptxTableCellStyle> {
	if (!borders) {
		return {};
	}
	const out: Partial<PptxTableCellStyle> = {};
	for (const [edge, prefix] of EDGES) {
		const own = borders[edge];
		const input = borders.all || own ? { ...borders.all, ...own } : undefined;
		if (input) {
			Object.assign(out, edgeStyle(own ? edge : 'all', prefix, input));
		}
	}
	return out;
}
