import { type CellRange, normalizeRange } from '../address.js';
import { deleteCell, getCell, putCell } from '../cells.js';
import type { Cell } from '../model.js';
import { type EditContext, sheetAt } from './context.js';
import { translateFormula } from './deps.js';
import { detectSeries, mod } from './fill-series.js';

/** A formula moved by (`dRow`, `dCol`); unparseable formulas are kept as they are. */
export function moveFormula(formula: string, dRow: number, dCol: number): string {
	if (!dRow && !dCol) return formula;
	try {
		return translateFormula(formula, dRow, dCol);
	} catch {
		return formula;
	}
}

/** A deep copy of a cell moved by (`dRow`, `dCol`). */
export function copyCell(cell: Cell, dRow: number, dCol: number): Cell {
	const copy = structuredClone(cell);
	if (copy.formula !== undefined) copy.formula = moveFormula(copy.formula, dRow, dCol);
	if (copy.arrayRange) delete copy.arrayRange;
	// A spilled value copies as a constant (its anchor's formula does not move with it).
	delete (copy as { spillAnchor?: unknown }).spillAnchor;
	return copy;
}

/**
 * AutoFill: extends `source` over `target` (which may include the source). The fill runs down,
 * up, right or left depending on where the target lies; each column (or row) of the source is a
 * lane whose series is detected separately. Formulas are translated relative to their new place.
 */
export function fillRange(ctx: EditContext, s: number, source: CellRange, target: CellRange): void {
	const sheet = sheetAt(ctx.workbook, s);
	const src = normalizeRange(source);
	const tgt = normalizeRange(target);
	const all: CellRange = {
		start: {
			row: Math.min(src.start.row, tgt.start.row),
			col: Math.min(src.start.col, tgt.start.col),
		},
		end: { row: Math.max(src.end.row, tgt.end.row), col: Math.max(src.end.col, tgt.end.col) },
	};
	const vertical = all.start.col === src.start.col && all.end.col === src.end.col;
	const horizontal = all.start.row === src.start.row && all.end.row === src.end.row;
	if (vertical && horizontal) return;
	if (!vertical && !horizontal) throw new Error('A fill must extend the source in one direction.');
	ctx.run(
		'Fill',
		'cells',
		[{ kind: 'cells', sheet: s, ranges: [all] }],
		() => {
			const laneCount = vertical
				? src.end.col - src.start.col + 1
				: src.end.row - src.start.row + 1;
			const length = vertical ? src.end.row - src.start.row + 1 : src.end.col - src.start.col + 1;
			for (let lane = 0; lane < laneCount; lane++) {
				const at = (i: number): [number, number] =>
					vertical
						? [src.start.row + i, src.start.col + lane]
						: [src.start.row + lane, src.start.col + i];
				const cells = Array.from({ length }, (_v, i) => getCell(sheet, ...at(i)));
				const series = detectSeries(ctx.workbook, cells);
				const from = vertical ? all.start.row - src.start.row : all.start.col - src.start.col;
				const to = vertical ? all.end.row - src.start.row : all.end.col - src.start.col;
				for (let k = from; k <= to; k++) {
					if (k >= 0 && k < length) continue;
					const [row, col] = at(k);
					const i = mod(k, length);
					const origin = cells[i];
					if (!origin) {
						deleteCell(sheet, row, col);
						continue;
					}
					const [srcRow, srcCol] = at(i);
					const step = series(k);
					const cell =
						step.kind === 'copy'
							? copyCell(origin, row - srcRow, col - srcCol)
							: { value: step.value, ...(origin.styleId ? { styleId: origin.styleId } : {}) };
					putCell(sheet, row, col, cell);
				}
			}
		},
		{ sheet: s, ranges: [all] },
	);
}
