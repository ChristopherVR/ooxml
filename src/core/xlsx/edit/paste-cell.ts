import { deleteCell, getCell, putCell } from '../cells.js';
import type { Cell, Workbook } from '../model.js';
import { internStyle, styleAt } from '../styles.js';
import { clearContents } from './cell-values.js';
import { ensureCell, pruneCell, type sheetAt } from './context.js';
import type { ClipboardCell, PasteMode, PasteOperation } from './types.js';
import { pasteArithmetic } from './paste-arithmetic.js';
export function writeClip(
	workbook: Workbook,
	sheet: ReturnType<typeof sheetAt>,
	row: number,
	col: number,
	clip: ClipboardCell | null,
	mode: PasteMode,
	formula: string | undefined,
	styleOf: (clip: ClipboardCell) => number | undefined,
	operation: PasteOperation,
): void {
	if (mode === 'formats') {
		const id = clip ? styleOf(clip) : 0;
		if (id === undefined) return;
		const cell = ensureCell(sheet, row, col);
		if (id) cell.styleId = id;
		else delete cell.styleId;
		pruneCell(sheet, row, col);
		return;
	}
	const existing = getCell(sheet, row, col);
	if (!clip && operation === 'none' && mode !== 'noBorders') {
		if (mode === 'all' || mode === 'transpose') deleteCell(sheet, row, col);
		else if (existing) {
			clearContents(existing);
			pruneCell(sheet, row, col);
		}
		return;
	}
	const keepStyle = mode === 'values' || mode === 'formulas' || (clip && !clip.style);
	const styleId =
		mode === 'noBorders'
			? internStyle(workbook, {
					...structuredClone(clip?.style ?? styleAt(workbook, 0)),
					border: structuredClone(styleAt(workbook, existing?.styleId).border),
				})
			: keepStyle
				? existing?.styleId
				: clip
					? styleOf(clip)
					: 0;
	const source = {
		value:
			operation !== 'none' &&
			mode === 'values' &&
			clip?.formula !== undefined &&
			typeof clip.value === 'boolean'
				? Number(clip.value)
				: (clip?.value ?? null),
		...(mode !== 'values' && formula !== undefined ? { formula } : {}),
	};
	const result = operation === 'none' ? source : pasteArithmetic(existing, source, operation);
	const cell: Cell = { value: result.value };
	if (styleId) cell.styleId = styleId;
	if (result.formula !== undefined) {
		cell.formula = result.formula;
		if (clip?.legacyFormula || (operation !== 'none' && existing?.legacyFormula))
			cell.legacyFormula = true;
	}
	if (clip?.richText && mode !== 'values' && mode !== 'formulas' && operation === 'none')
		cell.richText = structuredClone(clip.richText);
	if (clip?.numFmt && styleAt(workbook, cell.styleId).numFmt === 'General') {
		const base = styleAt(workbook, cell.styleId);
		const id = internStyle(workbook, { ...base, numFmt: clip.numFmt });
		if (id) cell.styleId = id;
	}
	putCell(sheet, row, col, cell);
	pruneCell(sheet, row, col);
}
