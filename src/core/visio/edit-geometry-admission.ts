import { isLocalBitmapShape } from './edit-foreign-bitmap';
import { cells, numeric } from './edit-geometry-cells';
export { cells, numeric, setCell } from './edit-geometry-cells';
import { assertEllipseResizeRow } from './edit-ellipse-geometry';
import { assertRelativeArcResizeRow } from './edit-path-create';
import { attribute, children } from './sheet';
import { fail } from './package-common';
import {
	analyzeVisioFormula,
	evaluateVisioFormula,
	visioFormulaCachedValue,
	type VisioFormulaReference,
	type VisioFormulaValue,
} from './formula';
import { createVisioDependencyQuery } from './edit-recalculate';
import type { VisioGeometryEdit } from './edit-commands';
import { executableCellFormula } from './cell-formula';

const locks = [
	'LockMoveX',
	'LockMoveY',
	'LockWidth',
	'LockHeight',
	'LockAspect',
	'LockDelete',
] as const;
/** Shared GUARD recognition for cells a native transform leaves in place. */
export function guardedCell(node: Element | undefined): boolean {
	const source = executableCellFormula(attribute(node, 'F'));
	return !!(source && analyzeVisioFormula(source).guarded);
}
export function editableCell(
	node: Element | undefined,
	resolve?: (reference: VisioFormulaReference) => VisioFormulaValue,
): void {
	if (!node) return;
	if (node.hasAttribute('E')) fail('UNSUPPORTED_GEOMETRY_EDIT', 'Cannot overwrite an error cell.');
	const source = attribute(node, 'F');
	if (source === 'Inh')
		fail('EDIT_PROTECTED_CELL', 'Inherited transform formulas cannot be overwritten.');
	const formula = executableCellFormula(source);
	if (!formula) return;
	const analysis = analyzeVisioFormula(formula);
	if (
		analysis.guarded ||
		analysis.dynamic ||
		(analysis.references.length && !resolve) ||
		analysis.unsupportedFunctions.length
	)
		fail(
			'EDIT_PROTECTED_CELL',
			'Transform formulas, GUARD and SETATREF redirection cannot be overwritten.',
		);
	evaluateVisioFormula(
		formula,
		resolve ?? (() => fail('UNSUPPORTED_GEOMETRY_EDIT', 'Unexpected formula dependency.')),
	);
}
export const isLineSheet = (local: ReadonlyMap<string, Element>): boolean =>
	['BeginX', 'BeginY', 'EndX', 'EndY'].some((name) => local.has(name)) ||
	numeric(local.get('OneD'), 0) !== 0;
export function admitted(
	root: Element,
	shapeId: string,
	masterMovePins: ReadonlySet<Element> = new Set(),
	masterDimensions: ReadonlyMap<Element, { width: number; height: number }> = new Map(),
	lineOperation?: 'move' | 'delete' | 'resize',
	rotationGroups: ReadonlySet<Element> = new Set(),
): Element {
	const containers = children(root, 'Shapes');
	if (containers.length !== 1)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'One local Shapes container is required.');
	const candidates = children(containers[0], 'Shape').filter(
		(node) => attribute(node, 'ID') === shapeId,
	);
	if (candidates.length !== 1)
		fail('EDIT_TARGET_NOT_FOUND', 'A unique top-level local shape is required.');
	const shape = candidates[0]!;
	if (
		((shape.hasAttribute('Master') || shape.hasAttribute('MasterShape')) &&
			!['PinX', 'PinY'].every((name) => {
				const cell = cells(shape).get(name);
				return cell && masterMovePins.has(cell);
			})) ||
		['1', 'true'].includes(attribute(shape, 'Del') ?? '') ||
		(children(shape, 'Shapes').length && !rotationGroups.has(shape)) ||
		(!isLocalBitmapShape(shape) &&
			(children(shape, 'ForeignData').length ||
				children(shape, 'Rel').length ||
				(attribute(shape, 'Type') &&
					attribute(shape, 'Type') !== 'Shape' &&
					!(attribute(shape, 'Type') === 'Group' && rotationGroups.has(shape)))))
	)
		fail(
			'UNSUPPORTED_GEOMETRY_EDIT',
			'Master, group, foreign, deleted and non-shape operations are unsupported.',
		);
	const local = cells(shape);
	const line = isLineSheet(local);
	// Removing a leaf does not rewrite or rely on its geometry caches.
	if (line && lineOperation === 'delete') return shape;
	assertLengthTransformCells(local);
	if (line && lineOperation !== 'move' && lineOperation !== 'resize')
		fail(
			'UNSUPPORTED_GEOMETRY_EDIT',
			'Only local 2D shapes are admitted; line routing and glue are unsupported.',
		);
	const proven = masterDimensions.get(shape);
	if (
		!(numeric(local.get('Width'), proven?.width) > 0) ||
		!(
			numeric(local.get('Height'), proven?.height) > 0 ||
			(line && !!lineOperation && numeric(local.get('Height')) === 0)
		)
	)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Positive proven Width and Height caches are required.');
	return shape;
}
/** Validate scalar protection; return effective rotation lock for flips that retain Angle.
 * Certified master moves preserve active non-movement locks.
 */
export function protectedShape(
	shape: Element,
	document: Element,
	masterMovePins: ReadonlySet<Element> = new Set(),
	additionalLocks: readonly ('LockBegin' | 'LockEnd' | 'LockRotate')[] = [],
	preserveRotation = false,
): boolean {
	const local = cells(shape);
	const protectionLocks = [...locks, ...additionalLocks];
	const certifiedMove =
		shape.hasAttribute('Master') &&
		['PinX', 'PinY'].every((name) => {
			const node = local.get(name);
			return node !== undefined && masterMovePins.has(node);
		});
	const styles = new Map<string, Element>();
	for (const container of children(document, 'StyleSheets'))
		for (const node of children(container, 'StyleSheet')) {
			const id = attribute(node, 'ID');
			if (!id || styles.has(id)) fail('EDIT_PROTECTED_CELL', 'Ambiguous protection style IDs.');
			styles.set(id, node);
		}
	const validateProtection = (cell: Element): void => {
		try {
			const cached = numeric(cell);
			const operationIndependent =
				(preserveRotation && attribute(cell, 'N') === 'LockRotate') ||
				(certifiedMove &&
					['LockWidth', 'LockHeight', 'LockAspect', 'LockDelete'].includes(
						attribute(cell, 'N') ?? '',
					));
			if (
				cell.hasAttribute('E') ||
				visioFormulaCachedValue(attribute(cell, 'V') ?? '', attribute(cell, 'U')).unit !==
					'scalar' ||
				!(cached === 0 || (operationIndependent && cached === 1))
			)
				fail('EDIT_PROTECTED_CELL', 'Inherited protection is active or invalid.');
			const source = executableCellFormula(attribute(cell, 'F'));
			if (!source) return;
			const analysis = analyzeVisioFormula(source);
			if (analysis.dynamic || analysis.references.length || analysis.unsupportedFunctions.length)
				fail('EDIT_PROTECTED_CELL', 'Inherited protection formula cannot be resolved safely.');
			const result = evaluateVisioFormula(source, () =>
				fail('EDIT_PROTECTED_CELL', 'Inherited protection has dependencies.'),
			);
			if (result.unit !== 'scalar' || result.value !== cached)
				fail('EDIT_PROTECTED_CELL', 'Inherited protection formula cache is stale.');
		} catch (error) {
			fail(
				'EDIT_PROTECTED_CELL',
				`Inherited protection cannot be proven safe for this operation: ${error instanceof Error ? error.message : 'invalid cache'}`,
			);
		}
	};
	const resolve = (
		id: string,
		category: string,
		lock: string,
		seen = new Set<string>(),
	): Element | undefined => {
		if (seen.size >= 64 || seen.has(id))
			fail('EDIT_PROTECTED_CELL', 'Protection style ancestry is cyclic or too deep.');
		seen.add(id);
		const style = styles.get(id);
		if (!style) fail('EDIT_PROTECTED_CELL', 'Protection style ancestry cannot be resolved.');
		const cell = cells(style).get(lock),
			parent = attribute(style, category);
		if (cell && attribute(cell, 'F') !== 'Inh') {
			validateProtection(cell);
			return cell;
		}
		if (cell) {
			// An Inh cache is not evidence: a real parent must prove the effective protection value.
			validateProtection(cell);
			if (parent === undefined || parent === id)
				fail('EDIT_PROTECTED_CELL', 'Inherited protection has no provable parent.');
			const inherited = resolve(parent, category, lock, seen);
			if (!inherited)
				fail('EDIT_PROTECTED_CELL', 'Inherited protection has no explicit provable ancestor.');
			if (numeric(cell) !== numeric(inherited))
				fail('EDIT_PROTECTED_CELL', 'Delegated protection cache disagrees with its ancestor.');
			return inherited;
		}
		return parent !== undefined && parent !== id
			? resolve(parent, category, lock, seen)
			: undefined;
	};
	const defaults = children(document, 'DocumentSheet')[0];
	let inheritedRotationLocked = false;
	for (const category of ['LineStyle', 'FillStyle', 'TextStyle']) {
		const id =
			attribute(shape, category) ??
			attribute(defaults, category) ??
			(styles.has('0') ? '0' : undefined);
		if (id === undefined) continue;
		for (const lock of protectionLocks) {
			const cell = resolve(id, category, lock);
			if (lock === 'LockRotate' && cell && numeric(cell) === 1) inheritedRotationLocked = true;
		}
	}
	for (const lock of protectionLocks)
		if (attribute(local.get(lock), 'F') === 'Inh')
			fail('EDIT_PROTECTED_CELL', 'Local inherited protection cannot be overridden.');

	for (const lock of protectionLocks) {
		const cell = local.get(lock);
		if (!cell) continue;
		if (
			preserveRotation &&
			lock === 'LockRotate' &&
			(cell.hasAttribute('E') || ![0, 1].includes(numeric(cell)))
		)
			fail('EDIT_PROTECTED_CELL', 'Rotation protection must be a scalar boolean.');
		if (visioFormulaCachedValue('0', attribute(cell, 'U')).unit !== 'scalar')
			fail('EDIT_FORMULA_UNIT', 'Protection cells must use scalar units.');
		const formula = executableCellFormula(attribute(cell, 'F'));
		if (formula) {
			const analysis = analyzeVisioFormula(formula);
			if (analysis.dynamic || analysis.references.length || analysis.unsupportedFunctions.length)
				fail('EDIT_PROTECTED_CELL', 'Protection formula cannot be resolved safely.');
			// Cached protection must agree with its constant formula before trusting it.
			const result = evaluateVisioFormula(formula, () =>
				fail('EDIT_PROTECTED_CELL', 'Protection dependency is unsupported.'),
			);
			if (result.unit !== 'scalar' || result.value !== numeric(cell))
				fail('EDIT_PROTECTED_CELL', 'Protection cache is stale.');
		}
	}
	return local.has('LockRotate')
		? numeric(local.get('LockRotate'), 0) !== 0
		: inheritedRotationLocked;
}
export function resizeGeometry(
	shape: Element,
	roots: ReadonlyMap<string, Element>,
	edit: VisioGeometryEdit,
	check: () => void,
	lineEditShapes: ReadonlySet<Element> = new Set(),
	glueShapes: ReadonlySet<Element> = new Set(),
): void {
	const sections = children(shape, 'Section').filter((node) => attribute(node, 'N') === 'Geometry');
	let depends: ReturnType<typeof createVisioDependencyQuery> | undefined;
	if (!sections.length)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Resizing requires explicit supported local geometry.');
	for (const section of sections) {
		if (section.hasAttribute('Del'))
			fail('UNSUPPORTED_GEOMETRY_EDIT', 'Deleted geometry is unsupported.');
		for (const row of children(section, 'Row')) {
			const type = attribute(row, 'T');
			if (type === 'Ellipse') assertEllipseResizeRow(shape, row);
			// Relative curve rows are fractions of Width and Height; the arc ratio needs its own proof.
			if (type === 'RelEllipticalArcTo') assertRelativeArcResizeRow(shape, row);
			const controlled = ['Ellipse', 'RelCubBezTo', 'RelEllipticalArcTo'].includes(type ?? '');
			if (
				row.hasAttribute('Del') ||
				![
					'MoveTo',
					'LineTo',
					'RelMoveTo',
					'RelLineTo',
					'Ellipse',
					'RelCubBezTo',
					'RelEllipticalArcTo',
				].includes(type ?? '')
			)
				fail(
					'UNSUPPORTED_GEOMETRY_EDIT',
					'Only local line-based geometry, relative curves and canonical ellipses can be resized.',
				);
			for (const node of children(row, 'Cell')) {
				if (
					!(controlled ? ['X', 'Y', 'A', 'B', 'C', 'D'] : ['X', 'Y']).includes(
						attribute(node, 'N') ?? '',
					)
				)
					fail('UNSUPPORTED_GEOMETRY_EDIT', 'Unknown geometry row cells are unsupported.');
				if ((type ?? '').startsWith('Rel')) continue;
				const formula = executableCellFormula(attribute(node, 'F'));
				if (!formula && numeric(node) === 0) continue;
				if (!formula)
					fail(
						'UNSUPPORTED_GEOMETRY_EDIT',
						'Absolute geometry needs a provable dimension-dependent formula.',
					);
				const analysis = analyzeVisioFormula(formula);
				if (!analysis.references.length && numeric(node) === 0) continue;
				const cell = `Geometry${Number(attribute(section, 'IX') ?? '0') + 1}.${attribute(node, 'N')}${attribute(row, 'IX')}`;
				depends ??= createVisioDependencyQuery(roots, { check, lineEditShapes, glueShapes });
				if (
					!depends(
						{ pageId: edit.pageId, shapeId: edit.shapeId, cell },
						['Width', 'Height'].map((cell) => ({
							pageId: edit.pageId,
							shapeId: edit.shapeId,
							cell,
						})),
					)
				)
					fail(
						'UNSUPPORTED_GEOMETRY_EDIT',
						'Absolute geometry scaling must explicitly reference Width or Height.',
					);
			}
		}
	}
}

/** Shared unit admission for local leaf and group descendant transforms. */
export function assertLengthTransformCells(local: ReadonlyMap<string, Element>): void {
	for (const name of ['Width', 'Height', 'PinX', 'PinY', 'LocPinX', 'LocPinY']) {
		const cell = local.get(name),
			unit = attribute(cell, 'U');
		if (unit && visioFormulaCachedValue('0', unit).unit !== 'length')
			fail('EDIT_FORMULA_UNIT', 'Transform cells must use length units.');
	}
}
