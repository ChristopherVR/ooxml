import { attribute, children } from './sheet';
import { fail } from './package-common';
import type { IndexedCell } from './edit-recalculate-index';

export interface VisioResizeWrite {
	/** Omitted for the top-level resize target. */
	shape?: Element;
	name: string;
	value: number;
	formula?: string;
}
export const resizeClose = (a: number, b: number) =>
	Math.abs(a - b) <= 1e-10 * Math.max(1, Math.abs(a), Math.abs(b));

/** Reject ambiguous spelling before trusting a transform or its protection caches. */
export function assertCanonicalResizeCells(shape: Element): void {
	const canonical = [
		'Width',
		'Height',
		'PinX',
		'PinY',
		'LocPinX',
		'LocPinY',
		'Angle',
		'FlipX',
		'FlipY',
		'LockMoveX',
		'LockMoveY',
		'LockWidth',
		'LockHeight',
		'LockAspect',
		'LockDelete',
		'ResizeMode',
	];
	for (const cell of children(shape, 'Cell')) {
		const name = attribute(cell, 'N') ?? '';
		if (
			canonical.some(
				(expected) => expected.toLowerCase() === name.toLowerCase() && expected !== name,
			)
		)
			fail(
				'EDIT_AMBIGUOUS_CELL',
				'Noncanonical transform and protection cell names are unsupported.',
			);
	}
}

/** Bounded dependency closure shared by leaf and group resize plans. */
export function resizeClosure(
	graph: ReadonlyMap<string, IndexedCell>,
	changed: readonly string[],
	check: () => void,
): Set<string> {
	const reverse = new Map<string, string[]>();
	for (const [id, item] of graph) {
		check();
		for (const dependency of item.dependencies) {
			const list = reverse.get(dependency) ?? [];
			list.push(id);
			reverse.set(dependency, list);
		}
	}
	const affected = new Set(changed),
		pending = [...changed];
	for (let index = 0; index < pending.length; index++) {
		check();
		for (const id of reverse.get(pending[index]!) ?? [])
			if (!affected.has(id)) {
				affected.add(id);
				pending.push(id);
			}
		if (affected.size > 10_000)
			fail('LIMIT_FORMULA_AFFECTED', 'Affected ShapeSheet cell limit exceeded.');
	}
	return affected;
}
