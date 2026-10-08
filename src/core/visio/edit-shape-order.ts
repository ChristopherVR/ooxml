import type { VisioShapeOrderEdit } from './edit-commands';
import { assertShapeLocks, effectiveShapeCell } from './edit-style-admission';
import { executableCellFormula } from './cell-formula';
import { analyzeVisioFormula, evaluateVisioFormula, visioFormulaCachedValue } from './formula';
import { fail } from './package-common';
import { attribute, children } from './sheet';
import type { VisioPackage } from './package';

/** Container lookup and dynamic references can depend on stacking order without naming a cell. */
export async function assertShapeOrderPackageScope(
	pkg: VisioPackage,
	check: () => void,
): Promise<void> {
	for (const path of pkg.paths()) {
		if (!/^visio\/.*\.xml$/i.test(path)) continue;
		const pending = [await pkg.readXml(path)];
		while (pending.length) {
			check();
			const node = pending.pop()!;
			const source = executableCellFormula(attribute(node, 'F'));
			if (source) {
				try {
					if (analyzeVisioFormula(source, { onStep: check }).dynamic)
						fail(
							'UNSUPPORTED_SHAPE_ORDER',
							'Dynamic or container formulas may depend on stacking order.',
						);
				} catch {
					fail(
						'UNSUPPORTED_SHAPE_ORDER',
						'Package formulas cannot be proven independent of stacking order.',
					);
				}
			}
			for (const child of Array.from(node.childNodes))
				if (child.nodeType === 1) pending.push(child as Element);
		}
	}
}

/** Ordering is limited to the ordinary display band until native band transitions are proven. */
function ordinaryDisplayBand(shape: Element, document: Element): void {
	for (const category of ['LineStyle', 'FillStyle', 'TextStyle'] as const) {
		const node = effectiveShapeCell(shape, document, 'DisplayLevel', category);
		if (!node) continue;
		const cached = visioFormulaCachedValue(attribute(node, 'V') ?? '', attribute(node, 'U'));
		if (node.hasAttribute('E') || cached.unit !== 'scalar' || cached.value !== 0)
			fail('UNSUPPORTED_SHAPE_ORDER', 'Only the ordinary zero display band can be reordered.');
		const source = executableCellFormula(attribute(node, 'F'));
		if (!source) continue;
		const analysis = analyzeVisioFormula(source);
		if (analysis.dynamic || analysis.references.length || analysis.unsupportedFunctions.length)
			fail('UNSUPPORTED_SHAPE_ORDER', 'Display band formulas must be constant.');
		const result = evaluateVisioFormula(source, () =>
			fail('UNSUPPORTED_SHAPE_ORDER', 'Display band dependencies cannot be resolved.'),
		);
		if (result.unit !== 'scalar' || result.value !== cached.value)
			fail('UNSUPPORTED_SHAPE_ORDER', 'Display band formula cache is stale.');
	}
}

/** Move an intact local shape node; identifiers, connections, geometry and text stay unchanged. */
export function reorderVisioShape(
	root: Element,
	document: Element,
	edit: VisioShapeOrderEdit,
	check: () => void,
): boolean {
	check();
	const containers = children(root, 'Shapes');
	if (containers.length !== 1)
		fail('UNSUPPORTED_SHAPE_ORDER', 'One local Shapes container is required.');
	const container = containers[0]!;
	const siblings = children(container, 'Shape');
	const ids = new Set<string>();
	for (const shape of siblings) {
		check();
		const id = attribute(shape, 'ID');
		if (!id || ids.has(id)) fail('INVALID_SHAPE_ID', 'Shape IDs must be present and unique.');
		ids.add(id);
		if (
			shape.hasAttribute('Master') ||
			shape.hasAttribute('MasterShape') ||
			shape.hasAttribute('Del') ||
			children(shape, 'Shapes').length ||
			children(shape, 'ForeignData').length ||
			children(shape, 'Rel').length ||
			(attribute(shape, 'Type') !== undefined && attribute(shape, 'Type') !== 'Shape')
		)
			fail(
				'UNSUPPORTED_SHAPE_ORDER',
				'Ordering requires local ordinary shapes in one display band.',
			);
		ordinaryDisplayBand(shape, document);
	}
	const index = siblings.findIndex((shape) => attribute(shape, 'ID') === edit.shapeId);
	if (index < 0) fail('EDIT_TARGET_NOT_FOUND', 'A top-level local shape is required.');
	const target = siblings[index]!;
	assertShapeLocks(target, document, ['LockSelect', 'LockFormat']);
	for (const category of ['LineStyle', 'FillStyle', 'TextStyle'] as const)
		if (effectiveShapeCell(target, document, 'LayerMember', category))
			fail('UNSUPPORTED_SHAPE_ORDER', 'Layered shape ordering is not yet supported.');
	const destination =
		edit.order === 'front'
			? siblings.length - 1
			: edit.order === 'back'
				? 0
				: edit.order === 'forward'
					? Math.min(siblings.length - 1, index + 1)
					: Math.max(0, index - 1);
	if (destination === index) return false;
	const anchor = siblings[destination]!;
	container.insertBefore(target, destination > index ? anchor.nextSibling : anchor);
	return true;
}
