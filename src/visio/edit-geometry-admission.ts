import { attribute, children } from './sheet.js';
import { fail } from './package-common.js';
import { analyzeVisioFormula, evaluateVisioFormula, visioFormulaCachedValue } from './formula.js';
import { visioCellDependsOn } from './edit-recalculate.js';
import type { VisioGeometryEdit } from './edit-commands.js';
import { executableCellFormula } from './cell-formula.js';

const locks = [
	'LockMoveX',
	'LockMoveY',
	'LockWidth',
	'LockHeight',
	'LockAspect',
	'LockDelete',
] as const;
export const cells = (shape: Element) =>
	new Map(children(shape, 'Cell').map((node) => [attribute(node, 'N') ?? '', node]));
export function numeric(node: Element | undefined, fallback?: number): number {
	if (!node && fallback !== undefined) return fallback;
	if (!node || node.hasAttribute('E'))
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'A usable local numeric transform cache is required.');
	const value = visioFormulaCachedValue(attribute(node, 'V') ?? '', attribute(node, 'U')).value;
	if (Math.abs(value) > 1e6)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Geometry cache exceeds coordinate limits.');
	return value;
}
export function editableCell(node: Element | undefined): void {
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
		analysis.references.length ||
		analysis.unsupportedFunctions.length
	)
		fail(
			'EDIT_PROTECTED_CELL',
			'Transform formulas, GUARD and SETATREF redirection cannot be overwritten.',
		);
	evaluateVisioFormula(formula, () =>
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Unexpected formula dependency.'),
	);
}
export function setCell(shape: Element, name: string, value: number, formula?: string): void {
	let node = cells(shape).get(name);
	if (!node) {
		node = shape.ownerDocument!.createElementNS(shape.namespaceURI, 'Cell');
		node.setAttribute('N', name);
		const before = Array.from(shape.childNodes).find(
			(child) => child.nodeType === 1 && (child as Element).localName !== 'Cell',
		);
		shape.insertBefore(node, before ?? null);
	}
	node.setAttribute('V', String(value));
	if (formula) node.setAttribute('F', formula);
	else if (attribute(node, 'F') !== 'No Formula') node.removeAttribute('F');
}
export function admitted(root: Element, shapeId: string): Element {
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
		shape.hasAttribute('Master') ||
		shape.hasAttribute('MasterShape') ||
		['1', 'true'].includes(attribute(shape, 'Del') ?? '') ||
		children(shape, 'Shapes').length ||
		children(shape, 'ForeignData').length ||
		children(shape, 'Rel').length ||
		(attribute(shape, 'Type') && attribute(shape, 'Type') !== 'Shape')
	)
		fail(
			'UNSUPPORTED_GEOMETRY_EDIT',
			'Master, group, foreign, deleted and non-shape operations are unsupported.',
		);
	const local = cells(shape);
	for (const name of ['Width', 'Height', 'PinX', 'PinY', 'LocPinX', 'LocPinY']) {
		const cell = local.get(name),
			unit = attribute(cell, 'U');
		if (unit && visioFormulaCachedValue('0', unit).unit !== 'length')
			fail('EDIT_FORMULA_UNIT', 'Transform cells must use length units.');
	}
	if (
		numeric(local.get('OneD'), 0) !== 0 ||
		['BeginX', 'BeginY', 'EndX', 'EndY'].some((name) => local.has(name))
	)
		fail(
			'UNSUPPORTED_GEOMETRY_EDIT',
			'Only local 2D shapes are admitted; line routing and glue are unsupported.',
		);
	if (!(numeric(local.get('Width')) > 0) || !(numeric(local.get('Height')) > 0))
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Positive local Width and Height caches are required.');
	return shape;
}
/** Only proven inactive inherited scalar protection is admitted; masters remain unsupported. */
export function protectedShape(shape: Element, document: Element): void {
	const local = cells(shape);
	const styles = new Map<string, Element>();
	for (const container of children(document, 'StyleSheets'))
		for (const node of children(container, 'StyleSheet')) {
			const id = attribute(node, 'ID');
			if (!id || styles.has(id)) fail('EDIT_PROTECTED_CELL', 'Ambiguous protection style IDs.');
			styles.set(id, node);
		}
	const inactive = (cell: Element): void => {
		try {
			if (
				cell.hasAttribute('E') ||
				visioFormulaCachedValue(attribute(cell, 'V') ?? '', attribute(cell, 'U')).unit !==
					'scalar' ||
				numeric(cell) !== 0
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
			if (result.unit !== 'scalar' || result.value !== 0)
				fail('EDIT_PROTECTED_CELL', 'Inherited protection formula is active or stale.');
		} catch (error) {
			fail(
				'EDIT_PROTECTED_CELL',
				`Inherited protection cannot be proven inactive: ${error instanceof Error ? error.message : 'invalid cache'}`,
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
			inactive(cell);
			return cell;
		}
		if (cell) {
			// An Inh zero cache is not evidence: a real parent must prove the effective protection value.
			inactive(cell);
			if (parent === undefined || parent === id)
				fail('EDIT_PROTECTED_CELL', 'Inherited protection has no provable parent.');
			const inherited = resolve(parent, category, lock, seen);
			if (!inherited)
				fail('EDIT_PROTECTED_CELL', 'Inherited protection has no explicit inactive ancestor.');
			return inherited;
		}
		return parent !== undefined && parent !== id
			? resolve(parent, category, lock, seen)
			: undefined;
	};
	const defaults = children(document, 'DocumentSheet')[0];
	for (const category of ['LineStyle', 'FillStyle', 'TextStyle']) {
		const id =
			attribute(shape, category) ??
			attribute(defaults, category) ??
			(styles.has('0') ? '0' : undefined);
		if (id === undefined) continue;
		for (const lock of locks) resolve(id, category, lock);
	}
	for (const lock of locks)
		if (attribute(local.get(lock), 'F') === 'Inh')
			fail('EDIT_PROTECTED_CELL', 'Local inherited protection cannot be overridden.');

	for (const lock of locks) {
		const cell = local.get(lock);
		if (!cell) continue;
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
}
export function resizeGeometry(
	shape: Element,
	roots: ReadonlyMap<string, Element>,
	edit: VisioGeometryEdit,
	check: () => void,
): void {
	const sections = children(shape, 'Section').filter((node) => attribute(node, 'N') === 'Geometry');
	if (!sections.length)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Resizing requires explicit supported local geometry.');
	for (const section of sections) {
		if (section.hasAttribute('Del'))
			fail('UNSUPPORTED_GEOMETRY_EDIT', 'Deleted geometry is unsupported.');
		for (const row of children(section, 'Row')) {
			const type = attribute(row, 'T');
			if (
				row.hasAttribute('Del') ||
				!['MoveTo', 'LineTo', 'RelMoveTo', 'RelLineTo'].includes(type ?? '')
			)
				fail('UNSUPPORTED_GEOMETRY_EDIT', 'This bundle resizes only local line-based geometry.');
			for (const node of children(row, 'Cell')) {
				if (!['X', 'Y'].includes(attribute(node, 'N') ?? ''))
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
				if (
					!visioCellDependsOn(
						roots,
						{ pageId: edit.pageId, shapeId: edit.shapeId, cell },
						['Width', 'Height'].map((cell) => ({
							pageId: edit.pageId,
							shapeId: edit.shapeId,
							cell,
						})),
						{ check },
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
