import { parseActionSteps } from './action-formula';
import { executableCellFormula } from './cell-formula';
import {
	cachedCellNumber,
	commitSheetCells,
	sheetScope,
	type SheetCellWrite,
} from './edit-sheet-cells';
import {
	effectiveFormula,
	effectiveNode,
	writeInstanceCell,
	type InstanceCell,
} from './edit-instance-sheet';
import type { MasterTemplate } from './edit-text-instance';
import {
	evaluateVisioFormula,
	parseVisioFormula,
	visioFormulaCachedValue,
	type VisioFormulaValue,
} from './formula';
import { evaluateMenuText } from './action-menu';
import { fail } from './package-common';
import { attribute } from './sheet';

/** Run one Actions row of a shape: its SETF steps change cells of the shape itself. */
export interface VisioActionEdit {
	type: 'run-action';
	pageId: string;
	shapeId: string;
	/** The Actions row: its IX, or its name when it has no IX. */
	row: string;
}

export const isVisioActionEdit = (edit: { type: string }): edit is VisioActionEdit =>
	edit.type === 'run-action';

/** Copy and validate an action edit; the shape and row are checked when it is applied. */
export function snapshotActionEdit(edit: VisioActionEdit): VisioActionEdit {
	if (typeof edit.pageId !== 'string' || !edit.pageId || edit.pageId.length > 256)
		fail('INVALID_EDIT', 'Invalid edit target.');
	if (typeof edit.shapeId !== 'string' || !/^[1-9]\d{0,9}$/.test(edit.shapeId))
		fail('INVALID_EDIT', 'Action shape IDs require canonical positive unsigned integers.');
	if (typeof edit.row !== 'string' || !/^[A-Za-z0-9_]{1,128}$/.test(edit.row))
		fail('INVALID_EDIT', 'Invalid action row.');
	return { type: 'run-action', pageId: edit.pageId, shapeId: edit.shapeId, row: edit.row };
}

const rowKey = (cell: InstanceCell): string | undefined =>
	cell.row && (attribute(cell.row, 'IX') ?? attribute(cell.row, 'N'));

const LENGTH =
	/^(width|height|(loc)?pin[xy]|geometry\d+\.[a-f]?[xy]?\d+|controls\..+|connections\..+)$/;

/**
 * Run an action the way Visio does: `SETF(GetRef(cell), value)` stores the evaluated value in the
 * cell as a plain value, then everything that follows is recalculated and the shape's menu texts
 * are refreshed (`Hide Box` after `Show Box`). Disabled actions, other functions, targets outside
 * the shape and values this editor cannot evaluate are refused with the reason.
 */
export async function applyActionEdit(
	roots: ReadonlyMap<string, Element>,
	edit: VisioActionEdit,
	template: MasterTemplate,
	check: () => void,
): Promise<readonly string[]> {
	const root = roots.get(edit.pageId);
	if (!root) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const scope = await sheetScope(root, edit.shapeId, template);
	check();
	const sheet = scope.root;
	const rowCell = (name: string) =>
		sheet.cells.find(
			(cell) => cell.section?.name === 'Actions' && rowKey(cell) === edit.row && cell.name === name,
		);
	const action = rowCell('Action');
	if (!action) fail('EDIT_TARGET_NOT_FOUND', 'The shape has no such action.');
	const disabled = rowCell('Disabled');
	if (disabled && cachedCellNumber(disabled, 'Disabled') !== 0)
		fail('EDIT_PROTECTED_CELL', 'This action is disabled.');
	const parsed = parseActionSteps(effectiveFormula(action!));
	if ('reason' in parsed) return fail('UNSUPPORTED_ACTION', parsed.reason);
	const read = (reference: { cell: string; shapeId?: string }): VisioFormulaValue => {
		if (reference.shapeId === undefined && /^(true|false)$/i.test(reference.cell))
			return { value: Number(reference.cell.toLowerCase() === 'true'), unit: 'scalar' };
		const cell = scope.resolve(action!, reference.cell, reference.shapeId, false);
		const node = cell && effectiveNode(cell);
		if (!cell || !node || node.hasAttribute('E'))
			return fail('UNSUPPORTED_ACTION', `${reference.cell} cannot be read.`);
		const value = visioFormulaCachedValue(attribute(node, 'V') ?? '', attribute(node, 'U'));
		return value.unit === 'scalar' && LENGTH.test(reference.cell.toLowerCase())
			? { ...value, unit: 'length' as const }
			: value;
	};
	const writes: SheetCellWrite[] = [];
	for (const step of parsed.steps) {
		const target = sheet.byName.get(step.target.toLowerCase());
		if (!target) fail('UNSUPPORTED_ACTION', `${step.target} is not a cell of this shape.`);
		let value: number;
		try {
			value = evaluateVisioFormula(step.value, read, { bareLengths: true }).value;
		} catch {
			return fail('UNSUPPORTED_ACTION', `The value for ${step.target} cannot be evaluated here.`);
		}
		const duplicate = writes.find((write) => write.cell === target);
		if (duplicate) duplicate.value = value;
		else writes.push({ cell: target!, value });
	}
	const pages = commitSheetCells(roots, edit.pageId, edit.shapeId, scope, writes, check);
	// Menu texts are strings: the numeric recalculation leaves them, so they are refreshed here.
	for (const cell of sheet.cells) {
		if (cell.section?.name !== 'Actions' || cell.name !== 'Menu') continue;
		const source = effectiveFormula(cell);
		if (!source || !executableCellFormula(source)) continue;
		let text: string | undefined;
		try {
			text = evaluateMenuText(parseVisioFormula(source), read);
		} catch {
			text = undefined;
		}
		if (text === undefined) continue;
		const node = effectiveNode(cell);
		if (attribute(node, 'V') === text) continue;
		const own = cell.local;
		if (own && executableCellFormula(attribute(own, 'F'))) own.setAttribute('V', text);
		else writeInstanceCell(sheet, cell, text, { formula: 'Inh' });
	}
	return pages;
}
