import { attribute } from './sheet';
import { executableCellFormula } from './cell-formula';
import { analyzeVisioFormula } from './formula';
import { editableCell } from './edit-geometry-admission';
import { indexCells, key } from './edit-recalculate-index';
import { createVisioCellEvaluator } from './edit-recalculate-values';
import type { VisioGeometryEdit } from './edit-commands';

/** Native rotations and flips replace unguarded formulas after their source dependency/cache proof. */
export function editableTransformCell(
	roots: ReadonlyMap<string, Element>,
	edit: VisioGeometryEdit,
	node: Element | undefined,
	check: () => void,
): void {
	const formula = executableCellFormula(attribute(node, 'F'));
	if (!formula || !analyzeVisioFormula(formula).references.length) return editableCell(node);
	const evaluate = createVisioCellEvaluator(indexCells(roots, { check }), { check }, true);
	editableCell(node, (reference) =>
		evaluate(
			key({
				pageId: edit.pageId,
				shapeId: reference.shapeId ?? edit.shapeId,
				cell: reference.cell,
			}),
		),
	);
	evaluate(key({ pageId: edit.pageId, shapeId: edit.shapeId, cell: attribute(node, 'N')! }));
}
