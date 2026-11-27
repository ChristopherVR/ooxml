/**
 * Glyph outlines of an embedded TrueType font as PDF path segments: for text the PDF draws as
 * shapes (WordArt bent along a warp, the silhouette a soft text shadow is blurred from). Glyphs
 * are placed as `PdfVectorPage.embeddedText` places them, so an outline lies exactly under the
 * text it stands for. DOM-free.
 */
import { glyphOutlinePath } from '../../text/index';
import type { PdfEmbeddedFont } from './pdf-embedded-font';
import type { PathSegment } from './pdf-path';
import type { PdfTextSpacing } from './pdf-vector';

export interface PlacedGlyph {
	char: string;
	glyph: number;
	/** Start of the glyph along the baseline from the run's origin, in text units (px). */
	x: number;
	/** Advance to the next glyph, spacing and scaling included. */
	advance: number;
}

/** Maps a point of the run (along the baseline, below it, y down) to the page. */
export type GlyphPointMap = (u: number, v: number) => readonly [number, number];

const glyphOf = (font: PdfEmbeddedFont, char: string): number => {
	const code = char.codePointAt(0)!;
	return font.glyphOf(code === 0xa0 || code < 32 ? 32 : code);
};

/** Where each glyph of `text` goes, as `embeddedText` writes it with `spacing`. */
export function placeGlyphs(
	font: PdfEmbeddedFont,
	size: number,
	text: string,
	spacing: PdfTextSpacing = {},
): PlacedGlyph[] {
	const scale = (spacing.scale ?? 100) / 100;
	const em = size / font.font.unitsPerEm;
	const out: PlacedGlyph[] = [];
	let x = 0;
	for (const char of text) {
		const glyph = glyphOf(font, char);
		const word = char === ' ' ? (spacing.word ?? 0) : 0;
		const advance = (font.font.advance(glyph) * em + (spacing.character ?? 0) + word) * scale;
		out.push({ char, glyph, x, advance });
		x += advance;
	}
	return out;
}

/**
 * One glyph's outline mapped by `map` from run units (u along the baseline from the glyph's
 * origin plus `origin`, v down from the baseline). With `bend`, lines and curves are cut into
 * short pieces first, so a curved map bends them instead of only moving their ends.
 */
export function glyphSegments(
	font: PdfEmbeddedFont,
	glyph: number,
	size: number,
	origin: number,
	map: GlyphPointMap,
	options: { scale?: number; bend?: boolean } = {},
): PathSegment[] {
	const em = size / font.font.unitsPerEm;
	const sx = em * (options.scale ?? 1);
	const local = (x: number, y: number) => map(origin + x * sx, -y * em);
	const piece = font.font.unitsPerEm * 0.08;
	const out: PathSegment[] = [];
	let cx = 0;
	let cy = 0;
	const quad = (qx: number, qy: number, x: number, y: number) => {
		const [x0, y0] = local(cx, cy);
		const [x1, y1] = local(qx, qy);
		const [x3, y3] = local(x, y);
		out.push({
			t: 'C',
			x1: x0 + (2 / 3) * (x1 - x0),
			y1: y0 + (2 / 3) * (y1 - y0),
			x2: x3 + (2 / 3) * (x1 - x3),
			y2: y3 + (2 / 3) * (y1 - y3),
			x: x3,
			y: y3,
		});
		cx = x;
		cy = y;
	};
	for (const command of glyphOutlinePath(font.font, glyph)) {
		if (command.t === 'Z') out.push({ t: 'Z' });
		else if (command.t === 'M') {
			const [x, y] = local(command.x, command.y);
			out.push({ t: 'M', x, y });
			[cx, cy] = [command.x, command.y];
		} else if (command.t === 'L') {
			const steps = options.bend
				? Math.min(32, Math.max(1, Math.ceil(Math.hypot(command.x - cx, command.y - cy) / piece)))
				: 1;
			const [fx, fy] = [cx, cy];
			for (let index = 1; index <= steps; index++) {
				const t = index / steps;
				const [x, y] = local(fx + (command.x - fx) * t, fy + (command.y - fy) * t);
				out.push({ t: 'L', x, y });
			}
			[cx, cy] = [command.x, command.y];
		} else if (options.bend) {
			// Two halves (de Casteljau) keep a bent curve close to the bent outline.
			const mx = (cx + 2 * command.cx + command.x) / 4;
			const my = (cy + 2 * command.cy + command.y) / 4;
			const [ax, ay] = [(cx + command.cx) / 2, (cy + command.cy) / 2];
			const [bx, by] = [(command.cx + command.x) / 2, (command.cy + command.y) / 2];
			quad(ax, ay, mx, my);
			quad(bx, by, command.x, command.y);
		} else quad(command.cx, command.cy, command.x, command.y);
	}
	return out;
}

/** The outlines of a straight run of `text` with its origin (baseline start) at x, y (y down). */
export function runOutlines(
	font: PdfEmbeddedFont,
	size: number,
	x: number,
	y: number,
	text: string,
	spacing: PdfTextSpacing = {},
): PathSegment[] {
	const scale = (spacing.scale ?? 100) / 100;
	return placeGlyphs(font, size, text, spacing).flatMap((placed) =>
		glyphSegments(font, placed.glyph, size, placed.x, (u, v) => [x + u, y + v], { scale }),
	);
}
