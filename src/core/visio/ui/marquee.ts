import type { VisioPage, VisioShape } from '../model';
import type { VisioShapeSelection } from './contract';
import type { VisioPagePoint } from './shape-move';
import { hasVisibleShapeContent } from './selection';
export interface VisioPageBox {
	x: number;
	y: number;
	width: number;
	height: number;
}
export function visioMarqueeBox(
	start: VisioPagePoint,
	end: VisioPagePoint,
): VisioPageBox | undefined {
	if (![start.x, start.y, end.x, end.y].every(Number.isFinite)) return undefined;
	const box = {
		x: Math.min(start.x, end.x),
		y: Math.min(start.y, end.y),
		width: Math.abs(end.x - start.x),
		height: Math.abs(end.y - start.y),
	};
	return Object.values(box).every(Number.isFinite) ? box : undefined;
}
/** A transformed rectangular selection extent in physical, downward-positive page inches. */
export function visioShapePageBox(page: VisioPage, shape: VisioShape): VisioPageBox | undefined {
	if (
		![page.height, shape.width, shape.height, ...shape.transform].every(Number.isFinite) ||
		shape.width < 0 ||
		shape.height < 0
	)
		return undefined;
	const [a, b, c, d, e, f] = shape.transform;
	const corners = [
		[0, 0],
		[shape.width, 0],
		[0, shape.height],
		[shape.width, shape.height],
	] as const;
	const xs = corners.map(([x, y]) => a * x + c * y + e),
		ys = corners.map(([x, y]) => b * x + d * y + f);
	const left = Math.min(...xs),
		right = Math.max(...xs),
		bottom = Math.min(...ys),
		top = Math.max(...ys);
	const box = { x: left, y: page.height - top, width: right - left, height: top - bottom };
	return Object.values(box).every(Number.isFinite) ? box : undefined;
}
/** Full enclosure of visible current-page top-level shapes; background sources are excluded. */
export function visioMarqueeSelection(
	page: VisioPage,
	box: VisioPageBox,
	visible?: WeakMap<VisioShape, boolean>,
): VisioShapeSelection[] {
	if (
		![box.x, box.y, box.width, box.height].every(Number.isFinite) ||
		box.width <= 0 ||
		box.height <= 0
	)
		return [];
	const cache = new WeakMap<VisioShape, boolean>();
	return page.shapes.flatMap((shape) => {
		const extent = visioShapePageBox(page, shape);
		return extent &&
			hasVisibleShapeContent(shape, visible, cache) &&
			extent.x >= box.x &&
			extent.y >= box.y &&
			extent.x + extent.width <= box.x + box.width &&
			extent.y + extent.height <= box.y + box.height
			? [{ id: shape.id, name: shape.name, pageId: page.id }]
			: [];
	});
}
