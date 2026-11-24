/**
 * Glyph outlines from a TrueType `glyf` table (simple and composite glyphs, OpenType spec
 * "glyf - Glyph Data"), as move, line, quadratic and close commands in font units, y up. Hinting
 * instructions are ignored, as PDF viewers ignore them. DOM-free; malformed data yields the
 * commands read so far, never an exception.
 */
import type { TrueTypeFont } from './parse';

export type GlyphPathCommand =
	| { t: 'M'; x: number; y: number }
	| { t: 'L'; x: number; y: number }
	| { t: 'Q'; cx: number; cy: number; x: number; y: number }
	| { t: 'Z' };

interface Point {
	x: number;
	y: number;
	on: boolean;
}

const MAX_POINTS = 0xffff;
const MAX_DEPTH = 8;

function readSimple(view: DataView, contours: number): Point[][] {
	const size = view.byteLength;
	const u16 = (at: number) => (at + 2 <= size ? view.getUint16(at) : 0);
	const ends: number[] = [];
	for (let index = 0; index < contours; index++) ends.push(u16(10 + index * 2));
	const count = (ends[ends.length - 1] ?? -1) + 1;
	if (count <= 0 || count > MAX_POINTS) return [];
	let at = 10 + contours * 2;
	at += 2 + u16(at);
	const flags = new Uint8Array(count);
	for (let index = 0; index < count && at < size;) {
		const flag = view.getUint8(at++);
		flags[index++] = flag;
		if (flag & 8 && at < size) {
			let repeat = view.getUint8(at++);
			while (repeat-- > 0 && index < count) flags[index++] = flag;
		}
	}
	const coordinates = (short: number, same: number): number[] => {
		const out: number[] = [];
		let value = 0;
		for (let index = 0; index < count; index++) {
			const flag = flags[index]!;
			if (flag & short) {
				const delta = at < size ? view.getUint8(at) : 0;
				at++;
				value += flag & same ? delta : -delta;
			} else if (!(flag & same)) {
				value += at + 2 <= size ? view.getInt16(at) : 0;
				at += 2;
			}
			out.push(value);
		}
		return out;
	};
	const xs = coordinates(2, 16);
	const ys = coordinates(4, 32);
	const result: Point[][] = [];
	let start = 0;
	for (const end of ends) {
		if (end < start || end >= count) break;
		const contour: Point[] = [];
		for (let index = start; index <= end; index++)
			contour.push({ x: xs[index]!, y: ys[index]!, on: (flags[index]! & 1) === 1 });
		result.push(contour);
		start = end + 1;
	}
	return result;
}

function readComposite(font: TrueTypeFont, view: DataView, depth: number): Point[][] {
	const size = view.byteLength;
	const result: Point[][] = [];
	let at = 10;
	for (let guard = 0; guard < 256 && at + 4 <= size; guard++) {
		const flags = view.getUint16(at);
		const glyph = view.getUint16(at + 2);
		at += 4;
		let dx = 0;
		let dy = 0;
		if (flags & 1) {
			if (at + 4 > size) break;
			if (flags & 2) {
				dx = view.getInt16(at);
				dy = view.getInt16(at + 2);
			}
			at += 4;
		} else {
			if (at + 2 > size) break;
			if (flags & 2) {
				dx = view.getInt8(at);
				dy = view.getInt8(at + 1);
			}
			at += 2;
		}
		const f2dot14 = (offset: number) => (offset + 2 <= size ? view.getInt16(offset) / 16384 : 0);
		let [a, b, c, d] = [1, 0, 0, 1];
		if (flags & 8) {
			a = d = f2dot14(at);
			at += 2;
		} else if (flags & 0x40) {
			a = f2dot14(at);
			d = f2dot14(at + 2);
			at += 4;
		} else if (flags & 0x80) {
			[a, b, c, d] = [f2dot14(at), f2dot14(at + 2), f2dot14(at + 4), f2dot14(at + 6)];
			at += 8;
		}
		for (const contour of glyphContours(font, glyph, depth + 1))
			result.push(
				contour.map((p) => ({ x: a * p.x + c * p.y + dx, y: b * p.x + d * p.y + dy, on: p.on })),
			);
		if (!(flags & 0x20)) break;
	}
	return result;
}

function glyphContours(font: TrueTypeFont, glyph: number, depth = 0): Point[][] {
	if (depth > MAX_DEPTH) return [];
	const bytes = font.outline(glyph);
	if (bytes.length < 10) return [];
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const contours = view.getInt16(0);
	return contours >= 0 ? readSimple(view, contours) : readComposite(font, view, depth);
}

/** One contour as commands: implied on-curve points sit midway between two off-curve points. */
function contourCommands(points: readonly Point[], out: GlyphPathCommand[]): void {
	const n = points.length;
	if (!n) return;
	const mid = (p: Point, q: Point): Point => ({ x: (p.x + q.x) / 2, y: (p.y + q.y) / 2, on: true });
	const first = points[0]!;
	const last = points[n - 1]!;
	const start = first.on ? first : last.on ? last : mid(last, first);
	out.push({ t: 'M', x: start.x, y: start.y });
	let control: Point | undefined;
	// The start point is written by the move; the walk ends by returning to it.
	const begin = first.on ? 1 : 0;
	const stop = !first.on && last.on ? n - 1 : n;
	const visit = (p: Point) => {
		if (p.on) {
			out.push(
				control
					? { t: 'Q', cx: control.x, cy: control.y, x: p.x, y: p.y }
					: { t: 'L', x: p.x, y: p.y },
			);
			control = undefined;
		} else if (control) {
			const m = mid(control, p);
			out.push({ t: 'Q', cx: control.x, cy: control.y, x: m.x, y: m.y });
			control = p;
		} else control = p;
	};
	for (let index = begin; index < stop; index++) visit(points[index]!);
	visit(start);
	out.push({ t: 'Z' });
}

/** The outline of `glyph` in font units (y up), as path commands; empty for a blank glyph. */
export function glyphOutlinePath(font: TrueTypeFont, glyph: number): GlyphPathCommand[] {
	const out: GlyphPathCommand[] = [];
	for (const contour of glyphContours(font, glyph)) contourCommands(contour, out);
	return out;
}
