import { placementForWrap } from './drawing-wrap';
import type { InlineImage, PicturePlacement } from './inline-model';
import type { ShapeSpec } from './shape-model';

export type DrawingPosition = `${'top' | 'center' | 'bottom'}-${'left' | 'center' | 'right'}`;
export type DrawingAlignment = 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom';
export type DrawingAlignTarget = 'page' | 'margin' | 'selected';
export type DrawingTextDirection = 'horizontal' | 'clockwise' | 'counterclockwise';

export interface DrawingBox {
	x: number;
	y: number;
	width: number;
	height: number;
}

/** The nine Position presets align to the page margins and use square wrapping. */
export function drawingPosition(image: InlineImage, position: DrawingPosition): PicturePlacement {
	const [vertical, horizontal] = position.split('-');
	const {
		offsetXPx: _x,
		offsetYPx: _y,
		offsetXPercent: _px,
		offsetYPercent: _py,
		...placement
	} = placementForWrap(image.anchored ? image.placement : undefined, 'square');
	return {
		...placement,
		relativeFrom: 'margin',
		relativeFromV: 'margin',
		align: horizontal === 'left' || horizontal === 'right' ? horizontal : 'center',
		alignV: vertical === 'top' || vertical === 'bottom' ? vertical : 'center',
	};
}

/** Changes the vertical reference frame while retaining the on-page position. */
export function drawingMovement(
	placement: PicturePlacement,
	mode: 'move' | 'fix',
	pageY: number,
	paragraphY: number,
): PicturePlacement {
	const { alignV: _align, ...rest } = placement;
	return {
		...rest,
		relativeFromV: mode === 'fix' ? 'page' : 'paragraph',
		offsetYPx: pageY - (mode === 'move' ? paragraphY : 0),
	};
}

/** Aligns one axis to a page or margin, preserving the reference and position of the other axis. */
export function alignDrawingToPage(
	placement: PicturePlacement,
	target: 'page' | 'margin',
	alignment: DrawingAlignment,
): PicturePlacement {
	if (alignment === 'left' || alignment === 'center' || alignment === 'right') {
		const { offsetXPx: _offset, offsetXPercent: _percent, ...rest } = placement;
		return { ...rest, relativeFrom: target, align: alignment };
	}
	const { offsetYPx: _offset, offsetYPercent: _percent, ...rest } = placement;
	return { ...rest, relativeFromV: target, alignV: alignment === 'middle' ? 'center' : alignment };
}

export function drawingBounds(boxes: readonly DrawingBox[]): DrawingBox | undefined {
	const first = boxes[0];
	if (!first) return undefined;
	let { x, y } = first;
	let right = x + first.width;
	let bottom = y + first.height;
	for (const box of boxes) {
		x = Math.min(x, box.x);
		y = Math.min(y, box.y);
		right = Math.max(right, box.x + box.width);
		bottom = Math.max(bottom, box.y + box.height);
	}
	return { x, y, width: right - x, height: bottom - y };
}

export function alignDrawingBox(
	box: DrawingBox,
	target: DrawingBox,
	alignment: DrawingAlignment,
): DrawingBox {
	switch (alignment) {
		case 'left':
			return { ...box, x: target.x };
		case 'center':
			return { ...box, x: target.x + (target.width - box.width) / 2 };
		case 'right':
			return { ...box, x: target.x + target.width - box.width };
		case 'top':
			return { ...box, y: target.y };
		case 'middle':
			return { ...box, y: target.y + (target.height - box.height) / 2 };
		case 'bottom':
			return { ...box, y: target.y + target.height - box.height };
	}
}

/** Gives the objects equal gaps inside the target along one axis, retaining their order. */
export function distributeDrawingBoxes(
	boxes: readonly DrawingBox[],
	target: DrawingBox,
	axis: 'x' | 'y',
): DrawingBox[] {
	const size = axis === 'x' ? 'width' : 'height';
	const ordered = boxes
		.map((box, index) => ({ box, index }))
		.sort((a, b) => a.box[axis] - b.box[axis]);
	const result = [...boxes];
	const gap =
		(target[size] - boxes.reduce((sum, box) => sum + box[size], 0)) / Math.max(1, boxes.length - 1);
	let offset = target[axis];
	for (const { box, index } of ordered) {
		result[index] = { ...box, [axis]: offset };
		offset += box[size] + gap;
	}
	return result;
}

export function placeDrawingBox(
	placement: PicturePlacement,
	box: DrawingBox,
	origins?: Partial<Pick<DrawingBox, 'x' | 'y'>>,
): PicturePlacement {
	const next = { ...placement };
	if (!origins || origins.x !== undefined) {
		delete next.align;
		delete next.offsetXPercent;
		if (!origins) next.relativeFrom = 'page';
		next.offsetXPx = box.x - (origins?.x ?? 0);
	}
	if (!origins || origins.y !== undefined) {
		delete next.alignV;
		delete next.offsetYPercent;
		if (!origins) next.relativeFromV = 'page';
		next.offsetYPx = box.y - (origins?.y ?? 0);
	}
	return next;
}

export function groupDrawingShapes(
	members: readonly { box: DrawingBox; shape: ShapeSpec }[],
): { box: DrawingBox; shape: ShapeSpec } | undefined {
	const box = drawingBounds(members.map((member) => member.box));
	if (!box || members.length < 2) return undefined;
	return {
		box,
		shape: {
			preset: 'rect',
			editable: false,
			textEditable: false,
			groupCanUngroup: true,
			groupExtent: { xPx: 0, yPx: 0, widthPx: box.width, heightPx: box.height },
			groupChildren: members.map(({ box: item, shape }) => ({
				xPx: item.x - box.x,
				yPx: item.y - box.y,
				widthPx: item.width,
				heightPx: item.height,
				shape,
			})),
		},
	};
}

export function ungroupDrawingShapes(
	box: DrawingBox,
	shape: ShapeSpec,
): {
	box: DrawingBox;
	shape: ShapeSpec;
	identity?: NonNullable<ShapeSpec['groupChildren']>[number]['identity'];
}[] {
	const extent = shape.groupExtent;
	if (!shape.groupCanUngroup || !extent || shape.rotation || shape.flipH || shape.flipV) return [];
	const sx = box.width / extent.widthPx;
	const sy = box.height / extent.heightPx;
	return (shape.groupChildren ?? []).map((child) => ({
		box: {
			x: box.x + (child.xPx - extent.xPx) * sx,
			y: box.y + (child.yPx - extent.yPx) * sy,
			width: child.widthPx * sx,
			height: child.heightPx * sy,
		},
		shape: child.shape,
		...(child.identity ? { identity: child.identity } : {}),
	}));
}

/** Computes the next stable order of one text layer. */
export function drawingStackOrder(
	heights: readonly number[],
	current: number,
	action: 'forward' | 'backward' | 'first' | 'last',
): number[] | undefined {
	const ordered = heights
		.map((height, index) => ({ height, index }))
		.sort((a, b) => a.height - b.height);
	const from = ordered.findIndex(({ index }) => index === current);
	const moved = ordered[from];
	if (!moved) return undefined;
	const to =
		action === 'first'
			? ordered.length - 1
			: action === 'last'
				? 0
				: action === 'forward'
					? from + 1
					: from - 1;
	if (from === to || to < 0 || to >= ordered.length) return undefined;
	ordered.splice(from, 1);
	ordered.splice(to, 0, moved);
	const result = [...heights];
	ordered.forEach(({ index }, order) => {
		result[index] = 251658240 + order;
	});
	return result;
}

export interface DrawingTransform {
	rotation?: number;
	flipH?: boolean;
	flipV?: boolean;
}

export function rotateDrawing(
	transform: DrawingTransform,
	action: 'right' | 'left' | 'flipH' | 'flipV' | number,
): DrawingTransform {
	if (action === 'flipH' || action === 'flipV')
		return { ...transform, [action]: !transform[action] };
	const degrees =
		typeof action === 'number'
			? action
			: (transform.rotation ?? 0) + (action === 'right' ? 90 : -90);
	return Number.isFinite(degrees)
		? { ...transform, rotation: ((degrees % 360) + 360) % 360 }
		: transform;
}
