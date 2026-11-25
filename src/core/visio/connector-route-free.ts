import {
	relevantObstacles,
	routeCrossings,
	VISIO_ROUTE_CLEARANCE as clearance,
} from './connector-route-avoid';
import type { VisioRouteBox, VisioRoutePoint } from './connector-route';

/**
 * The path recorded from Visio 16 for an aligned free Dynamic connector while a single
 * placeable shape moves across it. It turns just before the obstacle, passes its nearer side,
 * and turns back at the far end. A tie goes above a horizontal run or left of a vertical run.
 * When the obstacle clears, the running segment returns to the default offset on the other side.
 */
export function freeConnectorObstaclePath(
	begin: VisioRoutePoint,
	end: VisioRoutePoint,
	obstacles: readonly VisioRouteBox[],
	cleared: boolean,
): VisioRoutePoint[] | undefined {
	const vertical = Math.abs(begin.x - end.x) < 1e-9;
	if (!vertical && Math.abs(begin.y - end.y) >= 1e-9) return undefined;
	const nearby = relevantObstacles(begin, end, obstacles);
	const blocked = nearby.filter((box) => routeCrossings([begin, end], [box]));
	const toPage = (along: number, across: number) =>
		vertical ? { x: -across, y: along } : { x: along, y: across };
	const start = vertical ? begin.y : begin.x;
	const finish = vertical ? end.y : end.x;
	const run = vertical ? -begin.x : begin.y;
	if (start === finish || blocked.length > 1) return undefined;
	let turn = start;
	let across = run - clearance;
	const box = blocked[0];
	if (box) {
		const low = (vertical ? -box.maxX : box.minY) - clearance;
		const high = (vertical ? -box.minX : box.maxY) + clearance;
		across = run - low < high - run ? low : high;
		const min = vertical ? box.minY : box.minX;
		const max = vertical ? box.maxY : box.maxX;
		turn = start < finish ? min - clearance : max + clearance;
		if (!(Math.min(start, finish) < turn && turn < Math.max(start, finish))) return undefined;
	} else if (!cleared) return undefined;
	const path = [
		begin,
		...(turn === start ? [] : [toPage(turn, run)]),
		toPage(turn, across),
		toPage(finish, across),
		end,
	];
	return routeCrossings(path, nearby) ? undefined : path;
}
