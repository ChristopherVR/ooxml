import { createVisioCellEvaluator } from './edit-recalculate-values';
import { children, attribute } from './sheet';
import { fail } from './package-common';
import { analyzeVisioFormula, visioFormulaCachedValue, type VisioFormulaValue } from './formula';
import { plainShapeText } from './edit-recalculate-text';
import { VisioPackageError } from './package-common';
import {
	indexCells,
	key,
	bound,
	type VisioCellKey,
	type VisioRecalculationOptions,
} from './edit-recalculate-index';
export type { VisioCellKey, VisioRecalculationOptions } from './edit-recalculate-index';

/** Compute first, write caches only after the entire affected closure succeeds. F and U remain untouched. */
export function recalculateVisioCells(
	roots: ReadonlyMap<string, Element>,
	changed: readonly VisioCellKey[],
	options: VisioRecalculationOptions = {},
): readonly string[] {
	const cells = indexCells(roots, { scope: 'affected', follow: true, ...options });
	const reverse = new Map<string, string[]>();
	for (const [id, item] of cells)
		for (const dependency of item.dependencies) {
			const next = reverse.get(dependency) ?? [];
			next.push(id);
			reverse.set(dependency, next);
		}
	const affected = new Set(changed.map(key));
	// A formula that may read cells it does not name is reached by every edit on its page.
	const editedPages = new Set(changed.map((cell) => cell.pageId));
	for (const [id, item] of cells) if (item.opaque && editedPages.has(item.pageId)) affected.add(id);
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
	const evaluate = createVisioCellEvaluator(cells, {
		follow: true,
		affected,
		textOf: (pageId, shapeId) => plainShapeText(roots.get(pageId), shapeId),
		...options,
	});
	const values = new Map<string, VisioFormulaValue>();
	for (const id of affected) {
		// The edited text is an input, not a cell with a formula or a cache.
		if (!cells.has(id) && /,"thetext"]$/.test(id)) continue;
		// A cell whose formula cannot run keeps its cache; evaluating it only reports it.
		if (cells.get(id)?.keepCache && (options.follow ?? true)) {
			try {
				evaluate(id);
			} catch {
				// A cache that is not a number is left as it is.
			}
			continue;
		}
		try {
			values.set(id, evaluate(id));
		} catch (error) {
			if (
				options.cycles !== 'keep' ||
				!(error instanceof VisioPackageError) ||
				error.code !== 'EDIT_FORMULA_CYCLE'
			)
				throw error;
			const cyclic = cells.get(id)!;
			options.diagnostics?.push({
				code: 'FORMULA_CYCLE',
				message: `Kept the cached value of ${cyclic.cell} on shape ${cyclic.shapeId || 'page'} (page ${cyclic.pageId}): it is part of, or depends on, a dependency cycle.`,
			});
		}
	}
	const pages = new Set<string>();
	for (const id of affected) {
		const item = cells.get(id);
		if (item?.node && item.formula && values.has(id)) {
			const value = String(values.get(id)!.value);
			if (item.node.getAttribute('V') !== value) {
				item.node.setAttribute('V', value);
				pages.add(item.pageId);
			}
		}
	}
	return [...pages];
}

/**
 * Retained cells must not refer to a removed sheet. Selected owners disappear together. With
 * `freeze`, a retained cell that reads a removed sheet is instead returned to be frozen: Visio
 * replaces the formula of such a cell by its last result as a constant (`Sheet.2!Width*0.25`
 * becomes `0.5 in`; evidence/native-delete-error-2026-10-29.json) and does not recompute it.
 * Only a plain evaluable cell qualifies; a protected, error-valued or unanalysable one refuses.
 */
export function assertVisioShapesUnreferenced(
	roots: ReadonlyMap<string, Element>,
	removed: ReadonlyMap<string, ReadonlySet<string>>,
	options: VisioRecalculationOptions & { freeze?: boolean } = {},
): Element[] {
	const cells = indexCells(roots, { scope: 'affected', ...options });
	const frozen: Element[] = [];
	for (const cell of cells.values()) {
		options.check?.();
		if (cell.opaque && removed.has(cell.pageId) && !removed.get(cell.pageId)!.has(cell.shapeId))
			fail(
				'EDIT_REFERENCED_DELETE',
				`${cell.cell} on shape ${cell.shapeId || 'page'} may read the removed shape through a dynamic dependency.`,
			);
		if (removed.get(cell.pageId)?.has(cell.shapeId)) continue;
		if (
			cell.dependencies.some((id) => {
				const parts = JSON.parse(id) as string[];
				return removed.get(parts[0]!)?.has(parts[1]!) ?? false;
			})
		) {
			if (
				!options.freeze ||
				!cell.node ||
				!cell.formula ||
				cell.unsafe ||
				cell.node.hasAttribute('E') ||
				analyzeVisioFormula(cell.formula, { literalIndirect: true }).guarded ||
				!Number.isFinite(
					visioFormulaCachedValue(
						cell.node.getAttribute('V') ?? '',
						cell.node.getAttribute('U') ?? undefined,
					).value,
				)
			)
				fail('EDIT_REFERENCED_DELETE', 'Shape is referenced by a ShapeSheet formula.');
			frozen.push(cell.node);
		}
	}
	for (const [pageId, ids] of removed)
		for (const connections of children(roots.get(pageId), 'Connects'))
			for (const connection of children(connections, 'Connect')) {
				options.check?.();
				if (['FromSheet', 'ToSheet'].some((name) => ids.has(attribute(connection, name) ?? '')))
					fail('EDIT_REFERENCED_DELETE', 'Shape participates in a Connect record.');
			}
	return frozen;
}

/** Make the cells returned by `assertVisioShapesUnreferenced` constants holding their last result. */
export const freezeVisioCells = (nodes: readonly Element[]): void => {
	for (const node of nodes) node.removeAttribute('F');
};

/** Scalar deletion retains its existing reference policy. */
export function assertVisioShapeUnreferenced(
	roots: ReadonlyMap<string, Element>,
	pageId: string,
	shapeId: string,
	options: VisioRecalculationOptions & { freeze?: boolean } = {},
): Element[] {
	return assertVisioShapesUnreferenced(roots, new Map([[pageId, new Set([shapeId])]]), options);
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
	const cells = indexCells(roots, { scope: 'affected', ...options });
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
			if (
				item.formula &&
				analyzeVisioFormula(item.formula, { literalIndirect: true }).unsupportedFunctions.length
			)
				return false;
			if (wanted.has(id)) found = true;
			else pending.push(...item.dependencies);
		}
		return found;
	};
}
