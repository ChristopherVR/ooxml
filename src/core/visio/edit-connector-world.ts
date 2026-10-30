import { attribute, children } from './sheet';
import { fail } from './package-common';
import { transform } from './geometry';
import { cells, numeric } from './edit-geometry-cells';
import { effectiveCells, stencilTemplate } from './edit-stencil-connector';
import type { VisioMatrix } from './model';

/**
 * The shape named `shapeId` at any depth of the page, with the chain of groups above it (outermost
 * first, the shape last). Nothing is proven here; callers read the chain's cached cells.
 */
export function shapeChain(root: Element, shapeId: string): Element[] | undefined {
	const walk = (parent: Element, trail: Element[]): Element[] | undefined => {
		for (const container of children(parent, 'Shapes'))
			for (const shape of children(container, 'Shape')) {
				if (attribute(shape, 'ID') === shapeId) return [...trail, shape];
				const found = walk(shape, [...trail, shape]);
				if (found) return found;
			}
		return undefined;
	};
	return walk(root, []);
}

/** `parent` after `local`: maps a point of the shape through its group chain to the page. */
export function composeMatrix(parent: VisioMatrix, local: VisioMatrix): VisioMatrix {
	const [pa, pb, pc, pd, pe, pf] = parent,
		[la, lb, lc, ld, le, lf] = local;
	return [
		pa * la + pc * lb,
		pb * la + pd * lb,
		pa * lc + pc * ld,
		pb * lc + pd * ld,
		pa * le + pc * lf + pe,
		pb * le + pd * lf + pf,
	];
}

/** The cached size and local-to-parent transform of one shape of a chain (stencil-aware). */
export function localFrame(shape: Element): { width: number; height: number; matrix: VisioMatrix } {
	const local = stencilTemplate(shape) ? effectiveCells(shape) : cells(shape);
	const width = numeric(local.get('Width')),
		height = numeric(local.get('Height'));
	if (!(width > 0) || !(height > 0))
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Connectors glue to shapes with a positive size.');
	const flip = (name: string) => numeric(local.get(name), 0) !== 0;
	return {
		width,
		height,
		matrix: transform(
			numeric(local.get('PinX'), width / 2),
			numeric(local.get('PinY'), height / 2),
			numeric(local.get('LocPinX'), width / 2),
			numeric(local.get('LocPinY'), height / 2),
			numeric(local.get('Angle'), 0),
			flip('FlipX'),
			flip('FlipY'),
		),
	};
}

/**
 * The size and page transform of a shape at any depth: its own local-to-parent transform composed
 * with each group above it (PinX/PinY, LocPin, Angle and flips), so a connector glued to a group
 * member lands where the member is drawn.
 */
export function worldFrame(chain: readonly Element[]): {
	width: number;
	height: number;
	transform: VisioMatrix;
} {
	let world: VisioMatrix | undefined;
	let own = { width: 0, height: 0 };
	for (const shape of chain) {
		const frame = localFrame(shape);
		world = world ? composeMatrix(world, frame.matrix) : frame.matrix;
		own = frame;
	}
	if (!world) fail('EDIT_TARGET_NOT_FOUND', 'The glue target does not exist.');
	return { width: own.width, height: own.height, transform: world };
}
