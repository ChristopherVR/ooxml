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
): Stacking {
	const horizontal = direction === 'horizontal';
	const top = frame.maxY - (showTitle ? VISIO_SWIMLANE_TITLE : 0);
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
