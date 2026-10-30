import type { VisioShape } from './model';

/** A shape's page-space bounds (the transform maps its local box to the page). */
function bounds(shape: VisioShape) {
	const [a, b, c, d, e, f] = shape.transform;
	const xs: number[] = [];
	const ys: number[] = [];
	for (const [x, y] of [
		[0, 0],
		[shape.width, 0],
		[0, shape.height],
		[shape.width, shape.height],
	] as const) {
		xs.push(a * x + c * y + e);
		ys.push(b * x + d * y + f);
	}
	return {
		minX: Math.min(...xs),
		maxX: Math.max(...xs),
		minY: Math.min(...ys),
		maxY: Math.max(...ys),
	};
}

/**
 * Visio puts a shape that lies in a swimlane into that lane's members when it opens the drawing.
 * Do the same to the loaded model: every top-level, two-dimensional shape that has no structure
 * of its own and whose centre lies in a lane of a cross-functional flowchart (reached through
 * the frame's Swimlane List or directly) joins that lane's `memberIds`, after the members it
 * already has. Lists, lane order and everything else stay as saved, and so do the saved parts:
 * this only makes the model say what Visio shows. A shape that is a member of another lane (or
 * of any container) is left where it is.
 */
export function adoptSwimlaneContents(shapes: VisioShape[]): void {
	const byId = new Map(shapes.map((shape) => [shape.id, shape]));
	const lanes: VisioShape[] = [];
	for (const frame of shapes) {
		const structure = frame.structure;
		if (structure?.type !== 'container' || structure.swimlane?.role !== 'frame') continue;
		for (const id of structure.memberIds) {
			const member = byId.get(id);
			const inner = member?.structure;
			if (inner?.type === 'list' && inner.kind === 'swimlane') {
				for (const laneId of inner.memberIds) {
					const lane = byId.get(laneId);
					if (lane?.structure?.type === 'container' && lane.structure.swimlane?.role === 'lane')
						lanes.push(lane);
				}
			} else if (
				inner?.type === 'container' &&
				inner.swimlane?.role === 'lane' &&
				member &&
				!lanes.includes(member)
			)
				lanes.push(member);
		}
	}
	if (!lanes.length) return;
	const claimed = new Set<string>();
	for (const shape of shapes)
		if (shape.structure?.type === 'container' || shape.structure?.type === 'list')
			for (const id of shape.structure.memberIds) claimed.add(id);
	const laneBoxes = lanes.map((lane) => ({ lane, box: bounds(lane) }));
	for (const shape of shapes) {
		if (shape.structure || shape.kind === 'connector' || shape.hidden || claimed.has(shape.id))
			continue;
		const box = bounds(shape);
		const x = (box.minX + box.maxX) / 2;
		const y = (box.minY + box.maxY) / 2;
		const home = laneBoxes.find(
			({ box: lane }) => x >= lane.minX && x <= lane.maxX && y >= lane.minY && y <= lane.maxY,
		);
		const structure = home?.lane.structure;
		if (structure?.type !== 'container') continue;
		structure.memberIds = [...structure.memberIds, shape.id];
		claimed.add(shape.id);
	}
}
