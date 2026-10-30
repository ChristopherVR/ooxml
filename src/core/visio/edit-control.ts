import { rewriteControlFormula } from './control-formula';
import {
	cachedCellNumber,
	commitSheetCells,
	sheetScope,
	type SheetCellWrite,
} from './edit-sheet-cells';
import { effectiveFormula, effectiveNode, type InstanceCell } from './edit-instance-sheet';
import type { MasterTemplate } from './edit-text-instance';
import { evaluateVisioFormula, visioFormulaCachedValue, type VisioFormulaValue } from './formula';
import { fail } from './package-common';
import { attribute } from './sheet';

/**
 * Move a control handle: `x` and `y` are the handle's new position in the shape's local y-up
 * inches. An axis left out stays where it is.
 */
export interface VisioControlEdit {
	type: 'set-control';
	pageId: string;
	shapeId: string;
	/** The Controls row: its IX, or its name (`Row_1`) when it has no IX. */
	row: string;
	x?: number;
	y?: number;
}

export const isVisioControlEdit = (edit: { type: string }): edit is VisioControlEdit =>
	edit.type === 'set-control';

const finite = (value: unknown, label: string): number => {
	if (typeof value !== 'number' || !Number.isFinite(value) || Math.abs(value) > 1e6)
		fail('INVALID_EDIT', `${label} must be a finite number.`);
	return value as number;
};

/** Copy and validate a control edit; the shape and row are checked when it is applied. */
export function snapshotControlEdit(edit: VisioControlEdit): VisioControlEdit {
	if (typeof edit.pageId !== 'string' || !edit.pageId || edit.pageId.length > 256)
		fail('INVALID_EDIT', 'Invalid edit target.');
	if (typeof edit.shapeId !== 'string' || !/^[1-9]\d{0,9}$/.test(edit.shapeId))
		fail('INVALID_EDIT', 'Control shape IDs require canonical positive unsigned integers.');
	if (typeof edit.row !== 'string' || !/^[A-Za-z0-9_]{1,64}$/.test(edit.row))
		fail('INVALID_EDIT', 'Invalid control row.');
	if (edit.x === undefined && edit.y === undefined)
		fail('INVALID_EDIT', 'A control edit moves x or y.');
	return {
		type: 'set-control',
		pageId: edit.pageId,
		shapeId: edit.shapeId,
		row: edit.row,
		...(edit.x === undefined ? {} : { x: finite(edit.x, 'x') }),
		...(edit.y === undefined ? {} : { y: finite(edit.y, 'y') }),
	};
}

const rowKey = (cell: InstanceCell): string | undefined =>
	cell.row && (attribute(cell.row, 'IX') ?? attribute(cell.row, 'N'));
const isControl = (cell: InstanceCell) =>
	cell.section?.name === 'Control' || cell.section?.name === 'Controls';

/**
 * Drag a control handle the way Visio saves it: the X or Y cell of the Controls row becomes a
 * ratio of the shape size (inside the cell's own BOUND limits), and every cell that follows it
 * (geometry, connection points, text block) is recalculated. A handle whose position is computed
 * some other way, is locked on that axis (XCon or YCon not 0) or has a dependent this editor
 * cannot recompute is refused with the reason, and nothing is written.
 */
export async function applyControlEdit(
	roots: ReadonlyMap<string, Element>,
	edit: VisioControlEdit,
	template: MasterTemplate,
	check: () => void,
): Promise<readonly string[]> {
	const root = roots.get(edit.pageId);
	if (!root) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const scope = await sheetScope(root, edit.shapeId, template);
	check();
	const sheet = scope.root;
	const byName = (name: string) => sheet.byName.get(name.toLowerCase());
	const cellOf = (name: string) =>
		sheet.cells.find((cell) => isControl(cell) && rowKey(cell) === edit.row && cell.name === name);
	if (!cellOf('X') || !cellOf('Y'))
		fail('EDIT_TARGET_NOT_FOUND', 'The shape has no such control handle.');
	const dimensions = {
		x: cachedCellNumber(byName('Width'), 'The width'),
		y: cachedCellNumber(byName('Height'), 'The height'),
	};
	if (!(dimensions.x > 0) || !(dimensions.y > 0))
		fail('UNSUPPORTED_INSTANCE_EDIT', 'The shape has no positive size.');
	// The limits of BOUND read the cached sheet: the size and cells the move does not touch.
	const LENGTH =
		/^(width|height|(loc)?pin[xy]|geometry\d+\.[a-f]?[xy]?\d+|controls\..+|connections\..+)$/;
	const read = (reference: { cell: string; shapeId?: string }): VisioFormulaValue => {
		if (reference.shapeId === undefined && /^(true|false)$/i.test(reference.cell))
			return { value: Number(reference.cell.toLowerCase() === 'true'), unit: 'scalar' };
		const cell = scope.resolve(cellOf('X')!, reference.cell, reference.shapeId, false);
		const node = cell && effectiveNode(cell);
		if (!cell || !node || node.hasAttribute('E'))
			return fail(
				'EDIT_UNSUPPORTED_DEPENDENCY',
				'A control limit reads a cell this editor cannot read.',
			);
		const value = visioFormulaCachedValue(attribute(node, 'V') ?? '', attribute(node, 'U'));
		return value.unit === 'scalar' && LENGTH.test(reference.cell.toLowerCase())
			? { ...value, unit: 'length' }
			: value;
	};
	const writes: SheetCellWrite[] = [];
	for (const axis of ['x', 'y'] as const) {
		const requested = edit[axis];
		if (requested === undefined) continue;
		const name = axis === 'x' ? 'X' : 'Y';
		const cell = cellOf(name)!;
		const constraint = cellOf(axis === 'x' ? 'XCon' : 'YCon');
		if (constraint && cachedCellNumber(constraint, `${name}Con`) !== 0)
			fail(
				'EDIT_PROTECTED_CELL',
				`This control handle is fixed along ${axis === 'x' ? 'the width' : 'the height'}.`,
			);
		const ratio = requested / dimensions[axis];
		const formula = rewriteControlFormula(effectiveFormula(cell), axis, ratio);
		if (formula === undefined)
			fail(
				'EDIT_UNSUPPORTED_DEPENDENCY',
				`The control handle's ${name} is computed by a formula (${effectiveFormula(cell)}) that dragging cannot rewrite.`,
			);
		let value = requested;
		if (formula) {
			try {
				value = evaluateVisioFormula(formula, read, { bareLengths: true }).value;
			} catch {
				fail(
					'EDIT_UNSUPPORTED_DEPENDENCY',
					`The limits of the control handle's ${name} use a formula this editor cannot evaluate.`,
				);
			}
		}
		writes.push({ cell, value, ...(formula ? { formula } : {}) });
	}
	return commitSheetCells(roots, edit.pageId, edit.shapeId, scope, writes, check);
}
