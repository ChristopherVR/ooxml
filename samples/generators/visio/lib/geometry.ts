/**
 * Geometry sections written once and used twice: with formulas in a master or a drawn shape (so it
 * can be resized) and as cached overrides in an instance whose size differs from its master.
 */
import { cell, num, row, section } from './xml';

/** A coordinate: a plain number is a fraction of Width or Height; a pair adds inches to it. */
export type Co = number | readonly [fraction: number, inches: number];

export type Segment =
	| { t: 'M' | 'L'; x: Co; y: Co }
	/** Elliptical arc to (x, y) through (ax, ay); the ellipse's axes are parallel to the shape's. */
	| { t: 'E'; x: Co; y: Co; ax: Co; ay: Co; rx: Co; ry: Co }
	/** A whole ellipse centred on (x, y). */
	| { t: 'O'; x: Co; y: Co; rx: Co; ry: Co }
	/** Cubic Bezier in fractions of the shape: end point, then the two control points. */
	| { t: 'C'; x: number; y: number; x1: number; y1: number; x2: number; y2: number };

export interface Path {
	segments: readonly Segment[];
	noFill?: boolean;
	noLine?: boolean;
}

export interface Size {
	w: number;
	h: number;
}

const parts = (co: Co): readonly [number, number] => (typeof co === 'number' ? [co, 0] : co);

export function along(co: Co, length: number): number {
	const [fraction, inches] = parts(co);
	return fraction * length + inches;
}

export function coFormula(co: Co, dimension: string): string {
	const [fraction, inches] = parts(co);
	const scaled = `${dimension}*${num(fraction)}`;
	if (inches === 0) return scaled;
	return `${scaled}${inches < 0 ? '-' : '+'}${num(Math.abs(inches))}`;
}

type Mode = 'formula' | 'inherited';

const isFraction = (co: Co): boolean => typeof co === 'number' || co[1] === 0;
const fraction = (co: Co): number => (typeof co === 'number' ? co : co[0]);

/** A whole ellipse as two half arcs, so it can be written with relative rows. */
function arcs(segment: Extract<Segment, { t: 'O' }>): Segment[] {
	const [cx, cy, rx, ry] = [segment.x, segment.y, segment.rx, segment.ry].map(fraction) as [
		number,
		number,
		number,
		number,
	];
	return [
		{ t: 'M', x: cx - rx, y: cy },
		{ t: 'E', x: cx + rx, y: cy, ax: cx, ay: cy + ry, rx, ry },
		{ t: 'E', x: cx - rx, y: cy, ax: cx, ay: cy - ry, rx, ry },
	];
}

/**
 * A path whose coordinates are all fractions of the shape is written with relative rows: they
 * need no formulas and stay right when an instance is resized.
 */
function relativeRows(segments: readonly Segment[], size: Size, mode: Mode): string[] {
	const rows: string[] = [];
	const expanded = segments.flatMap((segment) => (segment.t === 'O' ? arcs(segment) : [segment]));
	expanded.forEach((segment, at) => {
		const index = at + 1;
		if (segment.t === 'E') {
			// D is the ratio of the ellipse's axes in inches, so it alone follows the shape's size.
			const [rx, ry] = [fraction(segment.rx), fraction(segment.ry)];
			const ratio = cell(
				'D',
				(rx * size.w) / (ry * size.h),
				mode === 'inherited' ? 'Inh' : `Width*${num(rx)}/(Height*${num(ry)})`,
			);
			if (mode === 'inherited')
				return rows.push(row({ T: 'RelEllipticalArcTo', IX: index }, [ratio]));
			return rows.push(
				row({ T: 'RelEllipticalArcTo', IX: index }, [
					cell('X', fraction(segment.x)),
					cell('Y', fraction(segment.y)),
					cell('A', fraction(segment.ax)),
					cell('B', fraction(segment.ay)),
					cell('C', 0),
					ratio,
				]),
			);
		}
		if (mode === 'inherited') return undefined;
		if (segment.t === 'C')
			return rows.push(
				row({ T: 'RelCubBezTo', IX: index }, [
					cell('X', segment.x),
					cell('Y', segment.y),
					cell('A', segment.x1),
					cell('B', segment.y1),
					cell('C', segment.x2),
					cell('D', segment.y2),
				]),
			);
		if (segment.t === 'O') return undefined;
		return rows.push(
			row({ T: segment.t === 'M' ? 'RelMoveTo' : 'RelLineTo', IX: index }, [
				cell('X', fraction(segment.x)),
				cell('Y', fraction(segment.y)),
			]),
		);
	});
	return rows;
}

function isRelative(path: Path): boolean {
	return path.segments.every((segment) => {
		if (segment.t === 'C') return true;
		const cos: Co[] = [segment.x, segment.y];
		if (segment.t === 'E') cos.push(segment.ax, segment.ay);
		if (segment.t === 'E' || segment.t === 'O') cos.push(segment.rx, segment.ry);
		return cos.every(isFraction);
	});
}

function coordinate(name: string, co: Co, size: Size, axis: 'x' | 'y', mode: Mode): string {
	const value = along(co, axis === 'x' ? size.w : size.h);
	const f = mode === 'inherited' ? 'Inh' : coFormula(co, axis === 'x' ? 'Width' : 'Height');
	return cell(name, value, f);
}

/** A row in inches, with formulas: for paths that mix fractions and fixed offsets. */
function absoluteRow(segment: Segment, index: number, size: Size, mode: Mode): string {
	const xy = (co: Co, axis: 'x' | 'y', name: string) => coordinate(name, co, size, axis, mode);
	if (segment.t === 'M' || segment.t === 'L')
		return row({ T: segment.t === 'M' ? 'MoveTo' : 'LineTo', IX: index }, [
			xy(segment.x, 'x', 'X'),
			xy(segment.y, 'y', 'Y'),
		]);
	if (segment.t === 'E') {
		const ratio = along(segment.rx, size.w) / along(segment.ry, size.h);
		const ratioFormula =
			mode === 'inherited'
				? 'Inh'
				: `(${coFormula(segment.rx, 'Width')})/(${coFormula(segment.ry, 'Height')})`;
		return row({ T: 'EllipticalArcTo', IX: index }, [
			xy(segment.x, 'x', 'X'),
			xy(segment.y, 'y', 'Y'),
			xy(segment.ax, 'x', 'A'),
			xy(segment.ay, 'y', 'B'),
			cell('C', 0),
			cell('D', ratio, ratioFormula),
		]);
	}
	if (segment.t === 'O') {
		const cx = along(segment.x, size.w);
		const cy = along(segment.y, size.h);
		const f = (text: string) => (mode === 'inherited' ? 'Inh' : text);
		const fx = coFormula(segment.x, 'Width');
		const fy = coFormula(segment.y, 'Height');
		return row({ T: 'Ellipse', IX: index }, [
			cell('X', cx, f(fx)),
			cell('Y', cy, f(fy)),
			cell('A', cx + along(segment.rx, size.w), f(`${fx}+(${coFormula(segment.rx, 'Width')})`)),
			cell('B', cy, f(fy)),
			cell('C', cx, f(fx)),
			cell('D', cy + along(segment.ry, size.h), f(`${fy}+(${coFormula(segment.ry, 'Height')})`)),
		]);
	}
	throw new Error('A Bezier segment needs a path made of fractions only.');
}

/** The Geometry sections of a shape drawn at `size`. */
export function geometry(paths: readonly Path[], size: Size, mode: Mode = 'formula'): string {
	return paths
		.map((path, index) => {
			const rows = isRelative(path)
				? relativeRows(path.segments, size, mode)
				: path.segments.map((segment, at) => absoluteRow(segment, at + 1, size, mode));
			const flags =
				mode === 'inherited'
					? []
					: [
							cell('NoFill', path.noFill ? 1 : 0),
							cell('NoLine', path.noLine ? 1 : 0),
							cell('NoShow', 0),
							cell('NoSnap', 0),
							cell('NoQuickDrag', 0),
						];
			return section('Geometry', rows, { ix: index, cells: flags });
		})
		.join('');
}

type Corner = readonly [Co, Co];

/** A closed polygon through the given corners. */
export function polygon(points: readonly Corner[], options: Partial<Path> = {}): Path {
	const first = points[0];
	if (!first) throw new Error('A polygon needs points.');
	const segments: Segment[] = points.map(([x, y], index) => ({ t: index ? 'L' : 'M', x, y }));
	segments.push({ t: 'L', x: first[0], y: first[1] });
	return { segments, ...options };
}

/** An open polyline: a stroke with no fill. */
export function polyline(points: readonly Corner[]): Path {
	return {
		segments: points.map(([x, y], index) => ({ t: index ? 'L' : 'M', x, y })),
		noFill: true,
	};
}

export const rectangle = (
	left: Co = 0,
	bottom: Co = 0,
	right: Co = 1,
	top: Co = 1,
	options: Partial<Path> = {},
): Path =>
	polygon(
		[
			[left, bottom],
			[right, bottom],
			[right, top],
			[left, top],
		],
		options,
	);

export const ellipse = (x: Co, y: Co, rx: Co, ry: Co, options: Partial<Path> = {}): Path => ({
	segments: [{ t: 'O', x, y, rx, ry }],
	...options,
});
