import type { VisioShapeSelection } from 'ooxml-core/visio/ui';

/**
 * The shape a pointer selects, as in Visio: the first click on a group member selects the
 * outermost group; clicking again inside a selected group subselects one level deeper.
 * Additive clicks always use the outermost group.
 */
export function pointerShapeTarget(
	node: EventTarget | null,
	selected: readonly Pick<VisioShapeSelection, 'id'>[],
	additive = false,
): SVGGElement | null {
	const chain: SVGGElement[] = [];
	let current = (node as Element | null)?.closest?.<SVGGElement>('[data-shape-id]') ?? null;
	while (current) {
		chain.push(current);
		current = current.parentElement?.closest<SVGGElement>('[data-shape-id]') ?? null;
	}
	if (chain.length < 2) return chain[0] ?? null;
	const outermost = chain[chain.length - 1]!;
	if (additive || selected.length !== 1) return outermost;
	const index = chain.findIndex((item) => item.dataset.shapeId === selected[0]!.id);
	return index < 0 ? outermost : chain[Math.max(0, index - 1)]!;
}
