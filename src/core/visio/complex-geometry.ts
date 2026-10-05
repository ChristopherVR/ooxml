import { normalizedNurbs } from './nurbs.js';
import { number, type Report, type Row } from './sheet.js';

// Literal data only, never a ShapeSheet expression evaluator. Syntax and units:
// https://learn.microsoft.com/en-us/office/client-developer/visio/polyline-function
// https://learn.microsoft.com/en-us/office/client-developer/visio/nurbs-function
// https://learn.microsoft.com/en-us/office/vba/api/visio.page.drawnurbs
const MAX_TEXT = 65_536;
const MAX_POINTS = 256;
// Keep digit runs separated by a required decimal point to avoid quadratic backtracking.
const numeric = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i;
type Point = { x: number; y: number };
type Control = Point & { weight: number };
const bounded = (n: number) => Number.isFinite(n) && Math.abs(n) <= 1e9;
const validPoint = (p: Point) => bounded(p.x) && bounded(p.y);
const command = (p: Point) => `L ${p.x === 0 ? 0 : p.x} ${p.y === 0 ? 0 : p.y}`;

function literalData(row: Row, report: Report): number[] | undefined {
	const polyline = row.type === 'PolylineTo' || row.type === 'PolyLineTo';
	const cell = row.cells.get(polyline ? 'A' : 'E');
	const text = cell?.formula ?? cell?.value;
	if (!text || text.length > MAX_TEXT) {
		report(
			'invalid-geometry',
			'Missing or oversized literal geometry data; its geometry section was omitted.',
		);
		return;
	}
	const name = polyline ? 'POLYLINE' : 'NURBS';
	const match = new RegExp(`^\\s*${name}\\s*\\(([^()]*)\\)\\s*$`, 'i').exec(text);
	const invalid = () =>
		report(
			'unsupported-geometry-formula',
			`${name} accepts only comma-separated numeric literals; references, units, expressions and nested functions are unsupported.`,
		);
	if (!match) {
		invalid();
		return;
	}
	// The text bound applies before regex/split allocation; limit arguments as well.
	const tokens = match[1]!.split(',');
	if (tokens.length > 4 + MAX_POINTS * 4) {
		report(
			'geometry-limit',
			'Literal geometry exceeds the bounded control-point count; its geometry section was omitted.',
		);
		return;
	}
	const values: number[] = [];
	for (const token of tokens) {
		const trimmed = token.trim();
		if (!numeric.test(trimmed)) {
			invalid();
			return;
		}
		const value = Number(trimmed);
		if (!bounded(value)) {
			report(
				'invalid-geometry',
				'Literal geometry contains an out-of-range number; its geometry section was omitted.',
			);
			return;
		}
		values.push(value);
	}
	return values;
}

/** Handles literal PolyLineTo/NURBSTo rows. Every emitted segment consumes the caller's budget. */
export function geometryRow(
	row: Row,
	startX: number,
	startY: number,
	width: number,
	height: number,
	report: Report,
	consume: () => void,
	consumeWork: (units: number) => void = () => {},
): string | undefined {
	if (row.type !== 'PolylineTo' && row.type !== 'PolyLineTo' && row.type !== 'NURBSTo') return;
	const values = literalData(row, report);
	if (!values) return;
	const invalid = (message: string) => {
		report('invalid-geometry', `${message}; its geometry section was omitted.`);
		return undefined;
	};
	const read = (name: string) =>
		(row.cells.get(name)?.value?.length ?? 0) > 128 ? NaN : number(row.cells, name, NaN, report);
	const end = { x: read('X'), y: read('Y') };
	if (
		!validPoint(end) ||
		!validPoint({ x: startX, y: startY }) ||
		!bounded(width) ||
		!bounded(height)
	)
		return invalid('Invalid geometry coordinates or dimensions');
	const polyline = row.type === 'PolylineTo' || row.type === 'PolyLineTo';
	const offset = polyline ? 2 : 4,
		stride = polyline ? 2 : 4;
	if (values.length < offset + stride || (values.length - offset) % stride)
		return invalid('Malformed literal geometry argument count');
	const count = (values.length - offset) / stride;
	if (count + (polyline ? 0 : 2) > MAX_POINTS)
		return invalid('Literal geometry exceeds 256 control points');
	const xType = values[offset - 2]!,
		yType = values[offset - 1]!;
	if ((xType !== 0 && xType !== 1) || (yType !== 0 && yType !== 1))
		return invalid('Geometry coordinate types must be 0 (relative) or 1 (local)');
	const points: Point[] = [];
	for (let i = offset; i < values.length; i += stride) {
		const p = {
			x: values[i]! * (xType === 0 ? width : 1),
			y: values[i + 1]! * (yType === 0 ? height : 1),
		};
		if (!validPoint(p)) return invalid('Scaled geometry coordinates are out of range');
		points.push(p);
	}
	if (polyline) {
		// The row's cached X/Y is the final vertex; formula vertices precede it.
		if (points[points.length - 1]!.x !== end.x || points[points.length - 1]!.y !== end.y)
			points.push(end);
		if (points.length > MAX_POINTS) return invalid('Literal geometry exceeds 256 vertices');
		return points
			.map((p) => {
				consume();
				return command(p);
			})
			.join(' ');
	}
	const degree = values[1]!;
	if (!Number.isInteger(degree) || degree < 1 || degree > 25 || points.length + 2 <= degree)
		return invalid('NURBS requires degree 1 to 25 and more control points than its degree');
	const knots = [read('C')],
		controls: Control[] = [{ x: startX, y: startY, weight: read('D') }];
	for (let i = 0; i < points.length; i++) {
		knots.push(values[offset + i * stride + 2]!);
		controls.push({ ...points[i]!, weight: values[offset + i * stride + 3]! });
	}
	controls.push({ ...end, weight: read('B') });
	knots.push(read('A'), values[0]!);
	return normalizedNurbs(controls, knots, degree, report, consume, 'NURBSTo', consumeWork);
}
