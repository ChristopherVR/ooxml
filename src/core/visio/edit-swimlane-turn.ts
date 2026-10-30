import { boundsOf } from './edit-container-fit';
import { containerClosure, shapeIdOf } from './edit-container-members';
import { turnSwimlanes } from './swimlane-stack';
import { VISIO_SEPARATOR_THICKNESS, type VisioSwimlaneDirection } from './edit-swimlane-commands';
import {
	applyFrameTitle,
	applyLaneGeometry,
	applyLaneLabel,
	applySeparatorLabel,
	setUserRow,
} from './edit-swimlane-shapes';
import {
	assertMovable,
	ensureLists,
	lanesOf,
	optionsOf,
	place,
	relayout,
	separatorsOf,
	shift,
	type Scope,
} from './edit-swimlane-layout';
import { fail } from './package-common';

/**
 * Direction > Horizontal or Vertical on an existing flowchart. Lanes keep their size along the
 * stacking axis, headings and dividers follow, phase bars turn with the flow and the shapes in a
 * lane are carried to the mirrored place (unless the caller moves them: contents 'leave').
 */
export function turnFlowchart(
	scope: Scope,
	frame: Element,
	direction: VisioSwimlaneDirection,
): void {
	const before = optionsOf(frame);
	if (before.direction === direction) return;
	ensureLists(scope, frame);
	const frameBox = boundsOf([frame]);
	const lanes = lanesOf(scope, frame, before.direction);
	if (!frameBox || !lanes.length) fail('UNSUPPORTED_GEOMETRY_EDIT', 'The flowchart has no lanes.');
	const turn = turnSwimlanes(
		before.direction,
		frameBox,
		before.title,
		lanes.map((lane) => ({ id: shapeIdOf(lane), box: boundsOf([lane])! })),
	);
	if (scope.contents === 'move') {
		const carried = lanes.map((lane) => [...containerClosure(scope.root, lane)]);
		for (const ids of carried) assertMovable(scope.root, ids);
		const byId = new Map(
			[...scope.root.getElementsByTagName('Shape')].map((shape) => [shapeIdOf(shape), shape]),
		);
		lanes.forEach((_, index) => {
			const lane = turn.lanes[index]!;
			for (const id of carried[index]!) {
				const box = byId.has(id) ? boundsOf([byId.get(id)!]) : undefined;
				if (!box) continue;
				const target = turn.point(lane, (box.minX + box.maxX) / 2, (box.minY + box.maxY) / 2);
				shift(
					scope,
					[id],
					target.x - (box.minX + box.maxX) / 2,
					target.y - (box.minY + box.maxY) / 2,
				);
			}
		});
	}
	const bars = separatorsOf(scope, frame);
	const barCentres = bars.map((bar) => {
		const box = boundsOf([bar])!;
		return before.direction === 'horizontal'
			? (box.minX + box.maxX) / 2
			: (box.minY + box.maxY) / 2;
	});
	setUserRow(frame, 'CFFVertical', direction === 'vertical' ? '1' : '0', '');
	const { swim, phase } = ensureLists(scope, frame);
	setUserRow(swim, 'msvSDListDirection', direction === 'horizontal' ? '2' : '0', '');
	setUserRow(phase, 'msvSDListDirection', direction === 'horizontal' ? '0' : '2', '');
	lanes.forEach((lane, index) => {
		const target = turn.lanes[index]!.to;
		place(scope, lane, target);
		applyLaneGeometry(lane, direction, {
			x: 0,
			y: 0,
			width: target.maxX - target.minX,
			height: target.maxY - target.minY,
		});
		applyLaneLabel(lane, direction, before.rotate);
	});
	const t = VISIO_SEPARATOR_THICKNESS;
	const area = turn.stacking.frame;
	bars.forEach((bar, index) => {
		const at = turn.separator(barCentres[index]!);
		place(
			scope,
			bar,
			direction === 'horizontal'
				? { minX: at - t / 2, maxX: at + t / 2, minY: turn.stacking.end, maxY: turn.stacking.top }
				: { minX: area.minX, maxX: area.maxX, minY: at - t / 2, maxY: at + t / 2 },
		);
		applySeparatorLabel(bar, direction, before.separators);
	});
	place(scope, frame, area);
	applyFrameTitle(frame, before.title);
	relayout(scope, frame);
}
