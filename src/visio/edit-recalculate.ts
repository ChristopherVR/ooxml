import { children, attribute } from './sheet.js';
import { fail } from './package-common.js';
import {
	analyzeVisioFormula,
	evaluateVisioFormula,
	visioFormulaCachedValue,
	type VisioFormulaValue,
} from './formula.js';
import {
	indexCells,
	key,
	bound,
	type VisioCellKey,
	type VisioRecalculationOptions,
} from './edit-recalculate-index.js';
export type { VisioCellKey, VisioRecalculationOptions } from './edit-recalculate-index.js';

/** Compute first, write caches only after the entire affected closure succeeds. F and U remain untouched. */
export function recalculateVisioCells(
	roots: ReadonlyMap<string, Element>,
	changed: readonly VisioCellKey[],
	options: VisioRecalculationOptions = {},
): readonly string[] {
	const cells = indexCells(roots, options);
	const reverse = new Map<string, string[]>();
	for (const [id, item] of cells)
		for (const dependency of item.dependencies) {
			const next = reverse.get(dependency) ?? [];
			next.push(id);
			reverse.set(dependency, next);
		}
	const affected = new Set(changed.map(key));
	const pending = [...affected];
	const maxAffected = bound(options.maxAffectedCells, 10_000);
	for (let i = 0; i < pending.length; i++) {
		options.check?.();
		for (const id of reverse.get(pending[i]!) ?? [])
			if (!affected.has(id)) {
				affected.add(id);
				pending.push(id);
			}
		if (affected.size > maxAffected)
			fail('LIMIT_FORMULA_AFFECTED', 'Affected ShapeSheet cell limit exceeded.');
	}
	const active = new Set<string>(),
		values = new Map<string, VisioFormulaValue>();
	let steps = bound(options.maxSteps, 100_000);
	const maxDepth = bound(options.maxDepth, 64);
	const evaluate = (id: string): VisioFormulaValue => {
		options.check?.();
		if (--steps < 0)
			fail('LIMIT_FORMULA_STEPS', 'ShapeSheet dependency evaluation limit exceeded.');
		const cached = values.get(id);
		if (cached) return cached;
		if (active.has(id)) fail('EDIT_FORMULA_CYCLE', 'Affected ShapeSheet dependency cycle.');
		if (active.size >= maxDepth)
			fail('LIMIT_FORMULA_DEPTH', 'ShapeSheet dependency depth limit exceeded.');
		const item = cells.get(id);
		if (!item || item.unsafe)
			fail(
				'EDIT_UNSUPPORTED_DEPENDENCY',
				'Affected formula references missing, inherited or grouped cells.',
			);
		if (item.node?.hasAttribute('E'))
			fail('EDIT_FORMULA_ERROR', 'Affected ShapeSheet cell has an error cache.');
		active.add(id);
		let result: VisioFormulaValue;
		if (item.formula) {
			const analysis = analyzeVisioFormula(item.formula);
			if (analysis.unsupportedFunctions.length)
				fail('EDIT_UNSUPPORTED_FORMULA', 'Affected formula uses unsupported functions.');
			result = evaluateVisioFormula(
				item.formula,
				(ref) =>
					evaluate(
						key({ pageId: item.pageId, shapeId: ref.shapeId ?? item.shapeId, cell: ref.cell }),
					),
				{
					onStep: () => {
						if (--steps < 0)
							fail('LIMIT_FORMULA_STEPS', 'ShapeSheet dependency evaluation limit exceeded.');
					},
				},
			);
		} else
			result = visioFormulaCachedValue(
				attribute(item.node, 'V') ?? '',
				attribute(item.node, 'U') ??
					(item.unit === 'length' ? 'DL' : item.unit === 'angle' ? 'DA' : undefined),
			);
		if (result.unit === 'scalar' && item.unit !== 'scalar') result = { ...result, unit: item.unit };
		const declared = attribute(item.node, 'U');
		const expected = declared ? visioFormulaCachedValue('0', declared).unit : item.unit;
		if (result.unit !== expected)
			fail('EDIT_FORMULA_UNIT', 'Affected formula cache has incompatible units.');
		if (!Number.isFinite(result.value) || Math.abs(result.value) > 1e9)
			fail('EDIT_FORMULA_VALUE', 'Affected formula cache exceeds supported magnitude.');
		active.delete(id);
		values.set(id, result);
		return result;
	};
	for (const id of affected) evaluate(id);
	const pages = new Set<string>();
	for (const id of affected) {
		const item = cells.get(id)!;
		if (item.node && item.formula) {
			const value = String(values.get(id)!.value);
			if (item.node.getAttribute('V') !== value) {
				item.node.setAttribute('V', value);
				pages.add(item.pageId);
			}
		}
	}
	return [...pages];
}

/** Deletion is only admitted when no static formula or Connect record refers to the shape. */
export function assertVisioShapeUnreferenced(
	roots: ReadonlyMap<string, Element>,
	pageId: string,
	shapeId: string,
	options: VisioRecalculationOptions = {},
): void {
	const cells = indexCells(roots, options);
	for (const cell of cells.values()) {
		if (cell.pageId !== pageId || cell.shapeId === shapeId) continue;
		if (
			cell.dependencies.some((id) => {
				const parts = JSON.parse(id) as string[];
				return parts[0] === pageId && parts[1] === shapeId;
			})
		)
			fail('EDIT_REFERENCED_DELETE', 'Shape is referenced by a ShapeSheet formula.');
	}
	const root = roots.get(pageId);
	if (root)
		for (const connections of children(root, 'Connects'))
			for (const connection of children(connections, 'Connect')) {
				if (['FromSheet', 'ToSheet'].some((name) => attribute(connection, name) === shapeId))
					fail('EDIT_REFERENCED_DELETE', 'Shape participates in a Connect record.');
			}
}

/** Static transitive proof for admission of dimension-dependent absolute geometry. */
export function visioCellDependsOn(
	roots: ReadonlyMap<string, Element>,
	source: VisioCellKey,
	targets: readonly VisioCellKey[],
	options: VisioRecalculationOptions = {},
): boolean {
	return createVisioDependencyQuery(roots, options)(source, targets);
}

/** Reuse a bounded index for package admission instead of rebuilding it per reference. */
export function createVisioDependencyQuery(
	roots: ReadonlyMap<string, Element>,
	options: VisioRecalculationOptions = {},
): (source: VisioCellKey, targets: readonly VisioCellKey[]) => boolean {
	const cells = indexCells(roots, options);
	return (source, targets) => {
		const wanted = new Set(targets.map(key));
		const seen = new Set<string>();
		const pending = [key(source)];
		const maxSteps = bound(options.maxSteps, 100_000);
		let found = false;
		while (pending.length) {
			options.check?.();
			const id = pending.pop()!;
			if (seen.has(id)) continue;
			seen.add(id);
			if (seen.size > maxSteps) fail('LIMIT_FORMULA_STEPS', 'Dependency analysis limit exceeded.');
			const item = cells.get(id);
			if (!item || item.unsafe || item.node?.hasAttribute('E')) return false;
			if (item.formula && analyzeVisioFormula(item.formula).unsupportedFunctions.length)
				return false;
			if (wanted.has(id)) found = true;
			else pending.push(...item.dependencies);
		}
		return found;
	};
}
