import type { VisioRouteBox, VisioRoutePoint } from './connector-route';

/**
 * Which routed connectors Visio spaces apart when they run on top of each other: the page's
 * LineAdjustFrom (Page Setup > Layout and Routing > Separate), with the spacing between lines
 * (LineToLineX for vertical runs, LineToLineY for horizontal runs, drawing inches).
 * https://learn.microsoft.com/office/client-developer/visio/lineadjustfrom-cell-page-layout-section
 */
export interface VisioConnectorSpacing {
	/** 0 unrelated lines, 1 all lines, 2 no lines, 3 the routing style's default. */
	separate: 'unrelated' | 'all' | 'none';
	lineToLineX: number;
	lineToLineY: number;
}

/** Visio's defaults for a page that stores no layout cells: Separate unrelated lines, 1/8 in. */
export const VISIO_CONNECTOR_SPACING_DEFAULTS: Readonly<VisioConnectorSpacing> = Object.freeze({
	separate: 'unrelated',
	lineToLineX: 0.125,
	lineToLineY: 0.125,
});

/**
 * The spacing policy of a page from its LineAdjustFrom value. The routing style default (3) is
 * read as Separate unrelated lines, the default of the right-angle (flowchart) style; other
 * routing styles' defaults are not modelled.
 */
export function visioConnectorSpacing(
	lineAdjustFrom: number | undefined,
	lineToLineX = VISIO_CONNECTOR_SPACING_DEFAULTS.lineToLineX,
	lineToLineY = VISIO_CONNECTOR_SPACING_DEFAULTS.lineToLineY,
): VisioConnectorSpacing {
	const separate = lineAdjustFrom === 1 ? 'all' : lineAdjustFrom === 2 ? 'none' : 'unrelated';
	const clamp = (value: number) => (Number.isFinite(value) && value > 0 ? Math.min(value, 10) : 0);
	return { separate, lineToLineX: clamp(lineToLineX), lineToLineY: clamp(lineToLineY) };
}

/** Another connector already on the page: its route and the shapes its ends are glued to. */
export interface VisioSpacedRoute {
	points: readonly VisioRoutePoint[];
	targets: ReadonlySet<string>;
}

const EPSILON = 1e-6;
/** The most line spacings a run moves away from where it was routed. */
const MAX_STEPS = 6;

type Segment = { a: VisioRoutePoint; b: VisioRoutePoint };
const horizontal = ({ a, b }: Segment) =>
	Math.abs(a.y - b.y) < EPSILON && Math.abs(a.x - b.x) > EPSILON;
const vertical = ({ a, b }: Segment) =>
	Math.abs(a.x - b.x) < EPSILON && Math.abs(a.y - b.y) > EPSILON;

/** True when two axis-aligned runs lie on one line and share a length, not just a point. */
function coincide(one: Segment, other: Segment, tolerance: number): boolean {
	if (horizontal(one) && horizontal(other)) {
		if (Math.abs(one.a.y - other.a.y) > tolerance) return false;
		const low = Math.max(Math.min(one.a.x, one.b.x), Math.min(other.a.x, other.b.x));
		const high = Math.min(Math.max(one.a.x, one.b.x), Math.max(other.a.x, other.b.x));
		return high - low > EPSILON;
	}
	if (vertical(one) && vertical(other)) {
		if (Math.abs(one.a.x - other.a.x) > tolerance) return false;
		const low = Math.max(Math.min(one.a.y, one.b.y), Math.min(other.a.y, other.b.y));
		const high = Math.min(Math.max(one.a.y, one.b.y), Math.max(other.a.y, other.b.y));
		return high - low > EPSILON;
	}
	return false;
}

const crossesBox = ({ a, b }: Segment, box: VisioRouteBox) =>
	Math.max(a.x, b.x) > box.minX + EPSILON &&
	Math.min(a.x, b.x) < box.maxX - EPSILON &&
	Math.max(a.y, b.y) > box.minY + EPSILON &&
	Math.min(a.y, b.y) < box.maxY - EPSILON;

const segments = (points: readonly VisioRoutePoint[]): Segment[] =>
	points.slice(1).map((b, index) => ({ a: points[index]!, b }));

/**
 * Space a right-angle route apart from the runs of other connectors it lies on, as Visio does for
 * parallel dynamic connectors: each interior run (one that does not leave or enter a shape) that
 * coincides with another connector's run is moved sideways by whole line spacings (nearest
 * first, alternating sides) to the first place that is free, keeps both neighbouring runs
 * pointing the way they did and does not cross an obstacle. The ends never move, so glue is kept.
 * Related connectors (sharing a glued shape) are left to overlap unless the page separates all
 * lines. Runs that cannot be moved this way (a route of one or two runs) stay where they are.
 */
export function spaceVisioRoute(
	points: readonly VisioRoutePoint[],
	targets: ReadonlySet<string>,
	others: readonly VisioSpacedRoute[],
	spacing: VisioConnectorSpacing,
	obstacles: readonly VisioRouteBox[] = [],
): VisioRoutePoint[] {
	const result = points.map((point) => ({ ...point }));
	if (spacing.separate === 'none' || result.length < 4) return result;
	const blockers = others
		.filter(
			(other) =>
				spacing.separate === 'all' || ![...other.targets].some((target) => targets.has(target)),
		)
		.flatMap((other) => segments(other.points));
	if (!blockers.length) return result;
	const taken = (segment: Segment, gap: number) =>
		blockers.some((blocker) => coincide(segment, blocker, gap / 4));
	for (let index = 1; index + 2 < result.length; index++) {
		const a = result[index]!,
			b = result[index + 1]!;
		const run = { a, b };
		const flat = horizontal(run);
		if (!flat && !vertical(run)) continue;
		const gap = flat ? spacing.lineToLineY : spacing.lineToLineX;
		if (gap <= 0 || !taken(run, gap)) continue;
		const axis = flat ? 'y' : 'x';
		const before = result[index - 1]![axis],
			after = result[index + 2]![axis],
			at = a[axis];
		const sign = (value: number) => (Math.abs(value) < EPSILON ? 0 : Math.sign(value));
		for (let step = 1; step <= 2 * MAX_STEPS; step++) {
			const offset = Math.ceil(step / 2) * gap * (step % 2 ? 1 : -1);
			const moved = at + offset;
			if (sign(moved - before) !== sign(at - before) || sign(after - moved) !== sign(after - at))
				continue;
			if (Math.abs(moved - before) < EPSILON && sign(at - before) !== 0) continue;
			const candidate: Segment = {
				a: { ...a, [axis]: moved },
				b: { ...b, [axis]: moved },
			};
			const neighbours: Segment[] = [
				{ a: result[index - 1]!, b: candidate.a },
				{ a: candidate.b, b: result[index + 2]! },
			];
			if (taken(candidate, gap)) continue;
			if (
				obstacles.some((box) =>
					[candidate, ...neighbours].some((segment) => crossesBox(segment, box)),
				)
			)
				continue;
			a[axis] = moved;
			b[axis] = moved;
			break;
		}
	}
	return result;
}
