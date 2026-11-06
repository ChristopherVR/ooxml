/**
 * Where a dashed CLOSED outline starts.
 *
 * A dash pattern begins at the first point of the path it is applied to, so
 * which point that is shows in the picture: it fixes the phase of every dash
 * round the shape and where the odd, shortened dash of the closing seam lands.
 * PowerPoint's own exports (measured on 18 presets, rotated and flipped, with
 * an asymmetric custom dash so the direction shows too) start a closed outline
 * at the LAST vertex of the preset path and run through the closing segment
 * first. In ECMA-376's path order that is:
 *
 *  - `rect`: the bottom-left corner, up the left edge (the path is
 *    top-left, top-right, bottom-right, bottom-left);
 *  - `triangle`: the bottom-right corner, along the base to the left;
 *  - `roundRect`, `diamond`, `pentagon`, `hexagon`, the arrows, `plus`,
 *    `parallelogram`, `trapezoid`, `octagon`, `star5`, callouts and custom
 *    geometry: likewise the last vertex, forward through the closing edge.
 *
 * A path whose last point is its first (`heart`, `cloud`: the figure closes
 * itself) is not rotated. Ellipses are the exception: `ellipse` and
 * `flowChartConnector` start at the left-middle point and run COUNTER-clockwise.
 * Flips and rotations carry the same start with the shape: PowerPoint applies
 * them to the path, as the SVG overlay's transform does.
 *
 * Only a single closed sub-path of absolute `M L C Q A Z` is rewritten; open
 * paths and anything else come back unchanged.
 *
 * Pure and framework-agnostic.
 *
 * @module render/dash-path-start
 */

type Point = readonly [number, number];

interface Segment {
	cmd: 'L' | 'C' | 'Q' | 'A';
	args: number[];
	end: Point;
}

interface ParsedFigure {
	start: Point;
	segments: Segment[];
	closed: boolean;
}

const ARGS_PER_COMMAND: Readonly<Record<string, number>> = { M: 2, L: 2, C: 6, Q: 4, A: 7 };

/** Presets drawn from a full ellipse: PowerPoint starts them on the left, counter-clockwise. */
const COUNTER_CLOCKWISE_PRESETS = new Set(['ellipse', 'oval', 'flowchartconnector']);

function parseFigure(d: string): ParsedFigure | undefined {
	const tokens = d.match(/[A-Za-z]|-?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?/giu);
	if (!tokens) return undefined;
	let start: Point | undefined;
	const segments: Segment[] = [];
	let closed = false;
	let cmd = '';
	let i = 0;
	while (i < tokens.length) {
		const token = tokens[i] as string;
		if (/^[A-Za-z]$/u.test(token)) {
			cmd = token;
			i += 1;
			if (cmd === 'Z') {
				closed = true;
				if (i < tokens.length) return undefined;
				continue;
			}
		}
		const count = ARGS_PER_COMMAND[cmd];
		if (count === undefined || closed) return undefined;
		const args = tokens.slice(i, i + count).map(Number);
		i += count;
		if (args.length < count || args.some((n) => !Number.isFinite(n))) return undefined;
		const end: Point = [args[count - 2] as number, args[count - 1] as number];
		if (cmd === 'M') {
			if (start) return undefined;
			start = end;
		} else {
			if (!start) return undefined;
			segments.push({ cmd: cmd as Segment['cmd'], args, end });
		}
	}
	return start && closed && segments.length > 0 ? { start, segments, closed } : undefined;
}

const num = (n: number): string => String(Number(n.toFixed(5)));

function writeSegment(segment: Segment): string {
	return `${segment.cmd} ${segment.args.map(num).join(' ')}`;
}

/** The same figure walked backwards, starting at its first point. */
function reverseSegments(figure: ParsedFigure): string[] {
	const points: Point[] = [figure.start, ...figure.segments.map((s) => s.end)];
	const out: string[] = [];
	for (let i = figure.segments.length - 1; i >= 0; i -= 1) {
		const segment = figure.segments[i] as Segment;
		const to = points[i] as Point;
		const a = segment.args;
		if (segment.cmd === 'L') {
			out.push(`L ${num(to[0])} ${num(to[1])}`);
		} else if (segment.cmd === 'C') {
			const c = a as [number, number, number, number, number, number];
			out.push(`C ${[c[2], c[3], c[0], c[1], to[0], to[1]].map(num).join(' ')}`);
		} else if (segment.cmd === 'Q') {
			out.push(`Q ${[a[0], a[1], to[0], to[1]].map((n) => num(n as number)).join(' ')}`);
		} else {
			const sweep = a[4] === 0 ? 1 : 0;
			out.push(
				`A ${[a[0], a[1], a[2], a[3], sweep, to[0], to[1]].map((n) => num(n as number)).join(' ')}`,
			);
		}
	}
	return out;
}

/**
 * The outline path data to dash so the pattern starts where PowerPoint starts
 * it. `shapeType` is the preset name as authored (not `getShapeType`, which
 * folds unknown presets into `rect`), or empty for custom geometry.
 */
export function dashStartPathData(d: string, shapeType?: string): string {
	const figure = parseFigure(d);
	if (!figure) return d;
	const lastEnd = (figure.segments[figure.segments.length - 1] as Segment).end;
	const selfClosing =
		Math.abs(lastEnd[0] - figure.start[0]) < 1e-3 && Math.abs(lastEnd[1] - figure.start[1]) < 1e-3;
	if (selfClosing) {
		if (!COUNTER_CLOCKWISE_PRESETS.has((shapeType ?? '').toLowerCase())) return d;
		return `M ${num(figure.start[0])} ${num(figure.start[1])} ${reverseSegments(figure).join(' ')} Z`;
	}
	// Start at the last vertex, run the closing edge to the first point, then the figure.
	return [
		`M ${num(lastEnd[0])} ${num(lastEnd[1])}`,
		`L ${num(figure.start[0])} ${num(figure.start[1])}`,
		...figure.segments.map(writeSegment),
		'Z',
	].join(' ');
}
