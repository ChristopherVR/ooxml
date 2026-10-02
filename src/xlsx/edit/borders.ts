import { type CellRange, normalizeRange } from '../address.js';
import type { Border, BorderEdge } from '../model.js';
import { ensureCell, type EditContext, forEachPosition, pruneCell, sheetAt } from './context.js';
import { stylePatcher } from './format.js';
import type { BorderPreset } from './types.js';

type Side = 'top' | 'bottom' | 'left' | 'right';

/** The border edges a preset sets on the cell at (`row`, `col`) of range `r`. */
export function presetEdges(
	preset: BorderPreset,
	r: CellRange,
	row: number,
	col: number,
	edge: BorderEdge,
): Partial<Border> | undefined {
	const atTop = row === r.start.row;
	const atBottom = row === r.end.row;
	const atLeft = col === r.start.col;
	const atRight = col === r.end.col;
	const out: Partial<Border> = {};
	const set = (side: Side, when: boolean, value: BorderEdge = edge): void => {
		if (when) out[side] = value;
	};
	const thick: BorderEdge = { ...edge, style: 'thick' };
	switch (preset) {
		case 'all':
			for (const side of ['top', 'bottom', 'left', 'right'] as const) set(side, true);
			break;
		case 'outside':
		case 'thickOutside': {
			const value = preset === 'thickOutside' ? thick : edge;
			set('top', atTop, value);
			set('bottom', atBottom, value);
			set('left', atLeft, value);
			set('right', atRight, value);
			break;
		}
		case 'none':
			return {
				top: undefined,
				bottom: undefined,
				left: undefined,
				right: undefined,
				diagonal: undefined,
				diagonalUp: undefined,
				diagonalDown: undefined,
			} as unknown as Partial<Border>;
		case 'top':
			set('top', atTop);
			break;
		case 'bottom':
			set('bottom', atBottom);
			break;
		case 'left':
			set('left', atLeft);
			break;
		case 'right':
			set('right', atRight);
			break;
		case 'thickBottom':
			set('bottom', atBottom, thick);
			break;
		case 'doubleBottom':
			set('bottom', atBottom, { ...edge, style: 'double' });
			break;
		case 'topAndBottom':
			set('top', atTop);
			set('bottom', atBottom);
			break;
		case 'insideH':
		case 'insideV':
		case 'inside':
			if (preset !== 'insideV') {
				set('top', !atTop);
				set('bottom', !atBottom);
			}
			if (preset !== 'insideH') {
				set('left', !atLeft);
				set('right', !atRight);
			}
			break;
	}
	return Object.keys(out).length ? out : undefined;
}

export function setBorders(
	ctx: EditContext,
	s: number,
	range: CellRange,
	preset: BorderPreset,
	edge: BorderEdge = { style: 'thin' },
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const r = normalizeRange(range);
	const patchers = new Map<string, (id: number | undefined) => number>();
	ctx.run(
		'Borders',
		'format',
		[{ kind: 'cells', sheet: s, ranges: [r] }],
		() =>
			forEachPosition(sheet, r, (row, col) => {
				const border = presetEdges(preset, r, row, col, edge);
				if (!border) return;
				const key = JSON.stringify(border, (_k, v: unknown) => (v === undefined ? null : v));
				let patcher = patchers.get(key);
				if (!patcher) {
					patcher = stylePatcher(ctx.workbook, { border });
					patchers.set(key, patcher);
				}
				const cell = ensureCell(sheet, row, col);
				const id = patcher(cell.styleId);
				if (id) cell.styleId = id;
				else delete cell.styleId;
				pruneCell(sheet, row, col);
			}),
		{ sheet: s, ranges: [r] },
	);
}
