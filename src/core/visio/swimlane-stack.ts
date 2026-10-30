import { VISIO_SWIMLANE_TITLE, type VisioSwimlaneDirection } from './edit-swimlane-commands';

/** An axis-aligned box in drawing inches (y up). */
export interface StackBox {
	minX: number;
	minY: number;
	maxX: number;
	maxY: number;
}
export interface StackLane {
	id: string;
	box: StackBox;
}
export interface StackedLane {
	id: string;
	/** Where the lane goes. */
	target: StackBox;
	/** How far its top-left corner (and so its shapes) moves. */
	dx: number;
	dy: number;
}
export interface Stacking {
	lanes: StackedLane[];
	/** The frame that spans the stacked lanes under its title bar; its top-left corner stays. */
	frame: StackBox;
	/** Where the lanes end along the stacking axis (their bottom or right edge). */
	end: number;
	/** The lanes' top edge (below the title). */
	top: number;
}

/** A lane's old box and the box the turned flowchart gives it. */
export interface TurnedLane {
	id: string;
	from: StackBox;
	to: StackBox;
}
export interface Turn {
	/** The direction the flowchart now has. */
	direction: VisioSwimlaneDirection;
	lanes: TurnedLane[];
	stacking: Stacking;
	/** Where a point of an old lane lands in its new lane: the lane's corner stays, the axes swap. */
	point(lane: TurnedLane, x: number, y: number): { x: number; y: number };
	/**
	 * The centre line of a phase separator after the turn: `centre` is its old x (horizontal
	 * flowcharts) or y (vertical ones), the result the new y or x.
	 */
	separator(centre: number): number;
}

/**
 * Turn a flowchart on its side (Direction > Horizontal or Vertical): lanes keep their size along
 * the stacking axis and the flow axis becomes the lane axis, so a shape one inch from a lane's
 * left edge and half an inch below its top lands one inch below the top and half an inch from the
 * left of the turned lane. The frame's top-left corner stays. `lanes` is in stacking order.
 */
export function turnSwimlanes(
	from: VisioSwimlaneDirection,
	frame: StackBox,
	showTitle: boolean,
	lanes: readonly StackLane[],
): Turn {
	const horizontal = from === 'horizontal';
	const direction: VisioSwimlaneDirection = horizontal ? 'vertical' : 'horizontal';
	const cross = horizontal
		? frame.maxX - frame.minX
		: Math.max(0, ...lanes.map(({ box }) => box.maxY - box.minY));
	const sizes = lanes.map(({ id, box }) => ({
		id,
		size: horizontal ? box.maxY - box.minY : box.maxX - box.minX,
	}));
	const oldTop = frame.maxY - (showTitle ? VISIO_SWIMLANE_TITLE : 0);
	const stacking = stackSwimlanes(
		direction,
		{ minX: frame.minX, maxX: frame.minX + cross, minY: frame.maxY - 1, maxY: frame.maxY },
		showTitle,
		sizes.map(({ id, size }) => ({
			id,
			box: horizontal
				? { minX: 0, maxX: size, minY: 0, maxY: cross }
				: { minX: 0, maxX: cross, minY: 0, maxY: size },
		})),
	);
	const turned = stacking.lanes.map((entry, index) => ({
		id: entry.id,
		from: lanes[index]!.box,
		to: entry.target,
	}));
	return {
		direction,
		lanes: turned,
		stacking,
		point: (lane, x, y) => ({
			x: lane.to.minX + (lane.from.maxY - y),
			y: lane.to.maxY - (x - lane.from.minX),
		}),
		separator: (centre) =>
			horizontal ? stacking.top - (centre - frame.minX) : frame.minX + (oldTop - centre),
	};
}

/**
 * Pure layout of a cross-functional flowchart: lanes edge to edge under the title bar, each
 * keeping its size along the stacking axis and taking the frame's width (horizontal flows) or
 * the tallest lane's height (vertical flows). `lanes` is in stacking order.
 */
export function stackSwimlanes(
	direction: VisioSwimlaneDirection,
	frame: StackBox,
	showTitle: boolean,
	lanes: readonly StackLane[],
	/** The title area above the lanes; the editor's own bar unless the chart has another (Visio's). */
	titleSize: number = VISIO_SWIMLANE_TITLE,
): Stacking {
	const horizontal = direction === 'horizontal';
	const top = frame.maxY - (showTitle ? titleSize : 0);
	const cross = horizontal
		? frame.maxX - frame.minX
		: Math.max(0, ...lanes.map(({ box }) => box.maxY - box.minY));
	let cursor = horizontal ? top : frame.minX;
	const stacked = lanes.map(({ id, box }): StackedLane => {
		const size = horizontal ? box.maxY - box.minY : box.maxX - box.minX;
		const target: StackBox = horizontal
			? { minX: frame.minX, maxX: frame.minX + cross, maxY: cursor, minY: cursor - size }
			: { minX: cursor, maxX: cursor + size, maxY: top, minY: top - cross };
		cursor += horizontal ? -size : size;
		return { id, target, dx: target.minX - box.minX, dy: target.maxY - box.maxY };
	});
	const end = horizontal ? cursor : top - cross;
	return {
		lanes: stacked,
		top,
		end,
		frame: horizontal
			? { minX: frame.minX, maxX: frame.minX + cross, maxY: frame.maxY, minY: end }
			: { minX: frame.minX, maxX: cursor, maxY: frame.maxY, minY: end },
	};
}
