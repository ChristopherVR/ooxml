import { computeSnap, type SnapBox, type SnapGuide } from '../../geometry/snap-guides';
import type { VisioGuideEdit } from '../edit-guide-commands';
import type { VisioPage, VisioShape } from '../model';
import { visioNextShapeId } from './shape-id';
import { visioMovementShape, type VisioPagePoint } from './shape-move';

/** A ruler guide in physical page inches: x for vertical guides, y-up y for horizontal ones. */
export interface VisioPageGuide {
	id: string;
	orientation: 'horizontal' | 'vertical';
	position: number;
}

/** Top-level horizontal and vertical guides of the page. Angled guides are left out. */
export function visioPageGuides(page: VisioPage): VisioPageGuide[] {
	const guides: VisioPageGuide[] = [];
	for (const shape of page.shapes) {
		if (!shape.visibility?.guide || !shape.rotation) continue;
		const { pinX, pinY, angle } = shape.rotation;
		const turn = ((angle % Math.PI) + Math.PI) % Math.PI;
		if (Math.abs(turn) < 1e-6 || Math.abs(turn - Math.PI) < 1e-6)
			guides.push({ id: shape.id, orientation: 'horizontal', position: pinY });
		else if (Math.abs(turn - Math.PI / 2) < 1e-6)
			guides.push({ id: shape.id, orientation: 'vertical', position: pinX });
	}
	return guides.filter((guide) => Number.isFinite(guide.position));
}

const drawing = (page: VisioPage, value: number) => {
	const ratio = page.drawingToPageScale ?? 1;
	if (!Number.isFinite(ratio) || ratio <= 0) throw new Error('Invalid page scale.');
	const result = value / ratio;
	if (!Number.isFinite(result)) throw new Error('Guide position exceeds finite limits.');
	return result;
};

/** A guide dragged out of a ruler: horizontal from the top ruler, vertical from the side one. */
export function visioGuideCreateCommand(
	page: VisioPage,
	orientation: 'horizontal' | 'vertical',
	position: number,
): VisioGuideEdit {
	return {
		type: 'create-guide',
		pageId: page.id,
		shapeId: visioNextShapeId(page),
		orientation,
		position: drawing(page, position),
	};
}
export function visioGuideMoveCommand(
	page: VisioPage,
	id: string,
	position: number,
): VisioGuideEdit {
	return { type: 'move-guide', pageId: page.id, shapeId: id, position: drawing(page, position) };
}

/** A shape's axis-aligned page box with y growing downward (y is the negated y-up top). */
export function visioShapeSnapBox(shape: VisioShape): SnapBox | undefined {
	const [a, b, c, d, e, f] = shape.transform;
	const corners = [
		[0, 0],
		[shape.width, 0],
		[0, shape.height],
		[shape.width, shape.height],
	] as const;
	const xs = corners.map(([x, y]) => a * x + c * y + e);
	const ys = corners.map(([x, y]) => b * x + d * y + f);
	const box = {
		x: Math.min(...xs),
		y: -Math.max(...ys),
		width: Math.max(...xs) - Math.min(...xs),
		height: Math.max(...ys) - Math.min(...ys),
	};
	return Object.values(box).every(Number.isFinite) ? box : undefined;
}

export interface VisioSnapOptions {
	/** Dynamic Grid: snap to other shapes' edges and centres. */
	shapes: boolean;
	/** Snap to the page's ruler guides. */
	guides: boolean;
	/** Page inches; a snap happens within this distance. */
	threshold: number;
}
export interface VisioSnapResult {
	/** Adjusted drag delta (x right, y up), as `visioPageDragDelta` returns it. */
	delta: VisioPagePoint;
	/** Alignment hints in page inches with y growing downward, for an overlay. */
	lines: SnapGuide[];
}

/** Snap the union box of a moving selection to shapes and guides (closest per axis). */
export function visioSnapMoveDelta(
	page: VisioPage,
	ids: readonly string[],
	delta: VisioPagePoint,
	options: VisioSnapOptions,
): VisioSnapResult {
	const unchanged = { delta, lines: [] };
	if ((!options.shapes && !options.guides) || !(options.threshold > 0)) return unchanged;
	const moving = new Set(ids);
	const boxes = ids
		.map((id) => visioMovementShape(page, id))
		.map((shape) => shape && visioShapeSnapBox(shape));
	if (!boxes.length || boxes.some((box) => !box)) return unchanged;
	const left = Math.min(...boxes.map((box) => box!.x));
	const top = Math.min(...boxes.map((box) => box!.y));
	const right = Math.max(...boxes.map((box) => box!.x + box!.width));
	const bottom = Math.max(...boxes.map((box) => box!.y + box!.height));
	const box = { x: left + delta.x, y: top - delta.y, width: right - left, height: bottom - top };
	const others: SnapBox[] = [];
	if (options.shapes)
		for (const shape of page.shapes) {
			if (moving.has(shape.id) || shape.hidden || shape.visibility?.guide) continue;
			const other = visioShapeSnapBox(shape);
			if (other && (other.width > 0 || other.height > 0)) others.push(other);
		}
	if (options.guides)
		for (const guide of visioPageGuides(page))
			others.push(
				guide.orientation === 'vertical'
					? { x: guide.position, y: -1e6, width: 0, height: 4e6 }
					: { x: -1e6, y: -guide.position, width: 4e6, height: 0 },
			);
	const snapped = computeSnap(box, others, options.threshold);
	return {
		delta: { x: delta.x + snapped.x - box.x, y: delta.y - (snapped.y - box.y) },
		lines: snapped.guides,
	};
}
