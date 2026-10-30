import type { VisioEdit } from '../edit-commands';
import type { VisioPage, VisioShape, VisioSwimlaneFrameOptions } from '../model';
import { VISIO_SWIMLANE_MIN, type VisioSwimlaneDirection } from '../edit-swimlane-commands';
import { stackSwimlanes, turnSwimlanes, type StackBox, type StackLane } from '../swimlane-stack';
import { visioNextShapeId } from './shape-id';
import { visioShapePageBounds } from './diagram-parts';
import { visioMoveCommands, visioMovementShape } from './shape-move';
import { visioPageEditToDrawing } from './page-edit';

/** A cross-functional flowchart on a page: its frame, options and lanes in stacking order. */
export interface VisioSwimlaneInfo {
	frame: VisioShape;
	options: VisioSwimlaneFrameOptions;
	lanes: VisioShape[];
}

const topLevel = (page: VisioPage, id: string) => page.shapes.find((shape) => shape.id === id);
const membersOf = (shape: VisioShape | undefined): string[] =>
	shape?.structure?.type === 'container' ? shape.structure.memberIds : [];
const scale = (page: VisioPage) => {
	const ratio = page.drawingToPageScale ?? 1;
	return Number.isFinite(ratio) && ratio > 0 ? ratio : 1;
};
/** A shape's box in drawing inches. */
function drawingBox(page: VisioPage, shape: VisioShape): StackBox {
	const ratio = scale(page);
	const box = visioShapePageBounds(shape);
	return {
		minX: box.minX / ratio,
		maxX: box.maxX / ratio,
		minY: box.minY / ratio,
		maxY: box.maxY / ratio,
	};
}

/**
 * The ids of a frame's lanes: through its Swimlane List, or directly for a flowchart saved as
 * plain containers (no lists yet).
 */
function laneIdsOf(page: VisioPage, frame: VisioShape, phases = false): string[] {
	return membersOf(frame).flatMap((id) => {
		const structure = topLevel(page, id)?.structure;
		if (structure?.type !== 'list') return [id];
		return structure.kind === 'swimlane' || (phases && structure.kind === 'phase')
			? structure.memberIds
			: [];
	});
}
/** The flowchart a frame, a lane or any shape inside a lane belongs to. */
export function visioSwimlaneInfo(page: VisioPage, shapeId: string): VisioSwimlaneInfo | undefined {
	const frame = page.shapes.find((shape) => {
		const role = shape.structure?.type === 'container' ? shape.structure.swimlane : undefined;
		if (role?.role !== 'frame') return false;
		if (shape.id === shapeId) return true;
		return laneIdsOf(page, shape, true).some((laneId) => {
			const lane = topLevel(page, laneId);
			return laneId === shapeId || membersOf(lane).includes(shapeId);
		});
	});
	const role = frame?.structure?.type === 'container' ? frame.structure.swimlane : undefined;
	if (!frame || role?.role !== 'frame') return undefined;
	const horizontal = role.options.direction === 'horizontal';
	const lanes = laneIdsOf(page, frame)
		.map((id) => topLevel(page, id))
		.filter(
			(lane): lane is VisioShape =>
				lane?.structure?.type === 'container' && lane.structure.swimlane?.role === 'lane',
		);
	const key = (lane: VisioShape) => {
		const box = visioShapePageBounds(lane);
		return horizontal ? -box.maxY : box.minX;
	};
	return { frame, options: role.options, lanes: lanes.sort((a, b) => key(a) - key(b)) };
}
/** Whether a selected shape is the frame or a lane (what the contextual tab keys on). */
export const visioIsSwimlaneShape = (page: VisioPage, shapeId: string): boolean =>
	visioSwimlaneInfo(page, shapeId) !== undefined;

/** Direct members of a lane that are not lanes themselves; nested ones move with their container. */
function contentIds(page: VisioPage, lane: VisioShape): string[] {
	return membersOf(lane).filter((id) => topLevel(page, id));
}
/** Every shape inside a lane, nested containers included. */
function contentClosure(page: VisioPage, lane: VisioShape): string[] {
	const seen = new Set<string>();
	const queue = [...contentIds(page, lane)];
	while (queue.length && seen.size < 5000) {
		const id = queue.pop()!;
		if (seen.has(id) || id === lane.id) continue;
		seen.add(id);
		queue.push(...membersOf(topLevel(page, id)).filter((member) => topLevel(page, member)));
	}
	return [...seen];
}

/** `move-shape` edits that carry each lane's shapes to where the new stacking puts the lane. */
function carryContents(
	page: VisioPage,
	stacked: readonly { id: string; dx: number; dy: number }[],
): VisioEdit[] | undefined {
	const ratio = scale(page);
	const edits: VisioEdit[] = [];
	for (const entry of stacked) {
		const lane = topLevel(page, entry.id);
		if (!lane || (Math.abs(entry.dx) < 1e-9 && Math.abs(entry.dy) < 1e-9)) continue;
		const ids = contentIds(page, lane);
		if (!ids.length) continue;
		const moves = visioMoveCommands(page, ids, { x: entry.dx * ratio, y: entry.dy * ratio });
		if (!moves) return undefined;
		edits.push(...moves);
	}
	return edits;
}
const stackInput = (page: VisioPage, lanes: readonly VisioShape[]): StackLane[] =>
	lanes.map((lane) => ({ id: lane.id, box: drawingBox(page, lane) }));

/** Room kept between a flowchart that grew past the page edge and the new edge, page inches. */
const PAGE_GROWTH_MARGIN = 0.25;
/**
 * The page grows with the frame: a `set-page-size` that makes room for `frame` (drawing inches)
 * when it reaches past the page's right or top edge, else nothing. The page never shrinks and the
 * origin stays, so every shape keeps its place. Applied in the same step as the flowchart edit.
 */
function growPage(page: VisioPage, frame: StackBox): VisioEdit[] {
	const ratio = scale(page);
	const fit = (needed: number, have: number) =>
		needed + 1e-6 > have ? Math.ceil((needed + PAGE_GROWTH_MARGIN) * 4) / 4 : have;
	const width = fit(frame.maxX * ratio, page.width);
	const height = fit(frame.maxY * ratio, page.height);
	return width > page.width || height > page.height
		? [{ type: 'set-page-size', pageId: page.id, width, height }]
		: [];
}
/** `edits` led by the page growth the turned or resized flowchart (its new `frame`) needs. */
const withGrowth = (page: VisioPage, frame: StackBox, edits: VisioEdit[]): VisioEdit[] => [
	...growPage(page, frame),
	...edits,
];

/** A move of one shape by a drawing-inch offset, or undefined when the editor cannot move it. */
function moveBy(page: VisioPage, id: string, dx: number, dy: number): VisioEdit | undefined {
	const shape = visioMovementShape(page, id);
	if (!shape?.rotation) return undefined;
	const ratio = scale(page);
	return visioPageEditToDrawing(page, {
		type: 'move-shape',
		pageId: page.id,
		shapeId: id,
		x: shape.rotation.pinX + dx * ratio,
		y: shape.rotation.pinY + dy * ratio,
	});
}

/** Insert Swimlane: a lane before or after the selected one, shifting the lanes after it. */
export function visioInsertSwimlane(
	page: VisioPage,
	laneId: string,
	position: 'before' | 'after',
	heading: string,
): VisioEdit[] | undefined {
	const info = visioSwimlaneInfo(page, laneId);
	const reference = info?.lanes.find((lane) => lane.id === laneId);
	if (!info || !reference) return undefined;
	let shapeId: string;
	try {
		shapeId = visioNextShapeId(page);
	} catch {
		return undefined;
	}
	const order = [...info.lanes];
	const index = order.indexOf(reference) + (position === 'after' ? 1 : 0);
	const inputs = stackInput(page, order);
	inputs.splice(index, 0, { id: shapeId, box: drawingBox(page, reference) });
	const stacking = stackSwimlanes(
		info.options.direction,
		drawingBox(page, info.frame),
		info.options.showTitle,
		inputs,
	);
	const moves = carryContents(page, stacking.lanes);
	if (!moves) return undefined;
	return withGrowth(page, stacking.frame, [
		{
			type: 'insert-swimlane',
			pageId: page.id,
			frameId: info.frame.id,
			shapeId,
			relativeTo: laneId,
			position,
			heading,
			contents: 'leave',
		},
		...moves,
	]);
}

/**
 * Delete Swimlane: the lane and the shapes in it go; the lanes after it close up. Deleting
 * stencil shapes needs a transaction of its own, so this is two batches to apply in order.
 */
export function visioDeleteSwimlane(page: VisioPage, laneId: string): VisioEdit[][] | undefined {
	const info = visioSwimlaneInfo(page, laneId);
	const lane = info?.lanes.find((item) => item.id === laneId);
	if (!info || !lane || info.lanes.length < 2) return undefined;
	const remaining = info.lanes.filter((item) => item !== lane);
	const stacking = stackSwimlanes(
		info.options.direction,
		drawingBox(page, info.frame),
		info.options.showTitle,
		stackInput(page, remaining),
	);
	const moves = carryContents(page, stacking.lanes);
	if (!moves) return undefined;
	// One batch: the edit deletes the lane's shapes with it (stencil instances included) and
	// `moves` close the gap, so undo steps back the whole command.
	return [
		[
			{
				type: 'delete-swimlane',
				pageId: page.id,
				frameId: info.frame.id,
				laneId,
				contents: 'leave',
			},
			...moves,
		],
	];
}

/** Resize a lane along the stacking axis (page inches); the lanes after it move and the frame follows. */
export function visioResizeSwimlane(
	page: VisioPage,
	laneId: string,
	size: number,
): VisioEdit[] | undefined {
	const info = visioSwimlaneInfo(page, laneId);
	const lane = info?.lanes.find((item) => item.id === laneId);
	if (!info || !lane || !Number.isFinite(size)) return undefined;
	const drawingSize = Math.max(VISIO_SWIMLANE_MIN, size / scale(page));
	const horizontal = info.options.direction === 'horizontal';
	const inputs = stackInput(page, info.lanes).map((entry) => {
		if (entry.id !== laneId) return entry;
		const box = entry.box;
		return {
			id: entry.id,
			box: horizontal
				? { ...box, minY: box.maxY - drawingSize }
				: { ...box, maxX: box.minX + drawingSize },
		};
	});
	const stacking = stackSwimlanes(
		info.options.direction,
		drawingBox(page, info.frame),
		info.options.showTitle,
		inputs,
	);
	const moves = carryContents(page, stacking.lanes);
	if (!moves) return undefined;
	return withGrowth(page, stacking.frame, [
		{
			type: 'resize-swimlane',
			pageId: page.id,
			frameId: info.frame.id,
			laneId,
			size: drawingSize,
			contents: 'leave',
		},
		...moves,
	]);
}

/** Insert Separator: a phase bar across the lanes, at the middle of the flowchart by default. */
export function visioInsertSeparator(
	page: VisioPage,
	shapeId: string,
	at?: number,
): VisioEdit[] | undefined {
	const info = visioSwimlaneInfo(page, shapeId);
	if (!info) return undefined;
	let id: string;
	try {
		id = visioNextShapeId(page);
	} catch {
		return undefined;
	}
	return [
		{
			type: 'insert-separator',
			pageId: page.id,
			frameId: info.frame.id,
			shapeId: id,
			...(at === undefined ? {} : { at: at / scale(page) }),
		},
	];
}

export interface VisioSwimlaneChange {
	showTitle?: boolean;
	showSeparators?: boolean;
	rotateLabels?: boolean;
	style?: number;
	/** Direction > Horizontal or Vertical; turning is a command of its own (no other option). */
	direction?: VisioSwimlaneDirection;
}

/**
 * Direction: turn the flowchart on its side. Lanes keep their size along the stacking axis,
 * every shape in a lane goes to the mirrored place in the turned lane, the phase bars turn with
 * the flow, and the page grows when the turned frame no longer fits.
 */
function turnEdits(page: VisioPage, info: VisioSwimlaneInfo, direction: VisioSwimlaneDirection) {
	const turn = turnSwimlanes(
		info.options.direction,
		drawingBox(page, info.frame),
		info.options.showTitle,
		stackInput(page, info.lanes),
	);
	const edits: VisioEdit[] = [];
	for (const [index, lane] of info.lanes.entries()) {
		for (const id of contentClosure(page, lane)) {
			const shape = topLevel(page, id);
			if (!shape) continue;
			const box = drawingBox(page, shape);
			const centre = { x: (box.minX + box.maxX) / 2, y: (box.minY + box.maxY) / 2 };
			const target = turn.point(turn.lanes[index]!, centre.x, centre.y);
			const dx = target.x - centre.x;
			const dy = target.y - centre.y;
			if (Math.abs(dx) < 1e-9 && Math.abs(dy) < 1e-9) continue;
			const move = moveBy(page, id, dx, dy);
			if (!move) return undefined;
			edits.push(move);
		}
	}
	return withGrowth(page, turn.stacking.frame, [
		{
			type: 'set-swimlane-options',
			pageId: page.id,
			frameId: info.frame.id,
			direction,
			contents: 'leave',
		},
		...edits,
	]);
}
/** The Design and Arrange options; hiding or showing the title bar shifts every lane. */
export function visioSwimlaneOptions(
	page: VisioPage,
	shapeId: string,
	change: VisioSwimlaneChange,
): VisioEdit[] | undefined {
	const info = visioSwimlaneInfo(page, shapeId);
	if (!info) return undefined;
	if (change.direction !== undefined) {
		if (Object.keys(change).length > 1) return undefined;
		return change.direction === info.options.direction
			? []
			: turnEdits(page, info, change.direction);
	}
	let moves: VisioEdit[] = [];
	if (change.showTitle !== undefined && change.showTitle !== info.options.showTitle) {
		const frame = drawingBox(page, info.frame);
		const stacking = stackSwimlanes(
			info.options.direction,
			frame,
			change.showTitle,
			stackInput(page, info.lanes),
		);
		const carried = carryContents(page, stacking.lanes);
		if (!carried) return undefined;
		moves = carried;
	}
	return [
		{
			type: 'set-swimlane-options',
			pageId: page.id,
			frameId: info.frame.id,
			...change,
			contents: 'leave',
		},
		...moves,
	];
}

/** File > New and Insert: a flowchart of `laneCount` lanes centred on the page. */
export function visioCreateSwimlanes(
	page: VisioPage,
	direction: VisioSwimlaneDirection,
	laneCount = 3,
): VisioEdit[] | undefined {
	if (!Number.isInteger(laneCount) || laneCount < 1 || laneCount > 20) return undefined;
	let first: number;
	try {
		first = Number(visioNextShapeId(page));
	} catch {
		return undefined;
	}
	if (first + laneCount + 2 >= 4294967295) return undefined;
	const ratio = scale(page);
	const width = Math.min(page.width * 0.85, 9) / ratio;
	const height = Math.min(page.height * 0.7, 6) / ratio;
	return [
		{
			type: 'create-swimlanes',
			pageId: page.id,
			frameId: String(first),
			laneIds: Array.from({ length: laneCount }, (_, index) => String(first + 1 + index)),
			headings: Array.from({ length: laneCount }, (_, index) => `Function ${index + 1}`),
			box: { x: page.width / 2 / ratio, y: page.height / 2 / ratio, width, height },
			direction,
			title: 'Title',
			listId: String(first + laneCount + 1),
			phaseListId: String(first + laneCount + 2),
		},
	];
}
