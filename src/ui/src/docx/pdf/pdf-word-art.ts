/**
 * WordArt for the Word PDF export, drawn as Print Layout draws it (`renderWordArtSvg`, through the
 * shared `wordArtLayout`): each line centred on its warp path in bold Arial, with its fill or
 * two-stop gradient, outline and soft shadow. With an embedded font the glyph outlines are bent
 * along the warp and filled and stroked as paths, with the text written invisibly over them so it
 * still selects, searches and copies; without one each letter is real text in a standard font,
 * placed and turned along the warp. DOM-free.
 */
import type { TextBoxContent } from 'ooxml-core/docx';
import {
	WARP_PATH_GENERATORS,
	WORD_ART_FONT,
	WORD_ART_SIZE_PT,
	wordArtLayout,
	type WordArtLayout,
} from 'ooxml-core/drawingml';
import { drawBlurredSilhouette, type SilhouettePart } from '../../export/pdf-blur-mask';
import type { PdfEmbeddedFont } from '../../export/pdf-embedded-font';
import { encodeWinAnsi, pdfBaseFont, pdfLiteral, pdfTextWidth } from '../../export/pdf-fonts';
import { glyphSegments, placeGlyphs } from '../../export/pdf-glyph-outlines';
import { pathOperators, type PathSegment } from '../../export/pdf-path';
import type { PathPoint } from '../../export/pdf-path-sampler';
import type { PdfVectorPage } from '../../export/pdf-vector';
import type { DocxPdfReportBuilder } from './pdf-report';
import type { DocxPdfText } from './pdf-text';
import {
	ASCENT,
	cellBounds,
	DESCENT,
	paintGradient,
	placement,
	rgb,
	wordArtLines,
	type LinePlacement,
	type Slot,
} from './pdf-word-art-geometry';
import { wordArtSoftCopies } from './pdf-text-effects';

const STANDARD_NOTE =
	'WordArt without an embedded font is written in a standard PDF font: each letter is placed and turned along the warp rather than bent, and its shadow is drawn as soft offset copies.';

/** Whether Print Layout draws this text box as WordArt (a warp or a gradient) rather than lines. */
export function isDrawnAsWordArt(content: TextBoxContent): boolean {
	const style = content.wordArt;
	if (!style) return false;
	if (!style.gradient && !(style.warp && Object.hasOwn(WARP_PATH_GENERATORS, style.warp.preset)))
		return false;
	return wordArtLines(content).some(Boolean);
}

function outlineLines(page: PdfVectorPage, font: PdfEmbeddedFont, layout: WordArtLayout): void {
	const size = layout.size;
	const em = size / font.font.unitsPerEm;
	const drawn: { segments: PathSegment[]; bounds?: [number, number, number, number] }[] = [];
	const glyphs: { slot: Slot; frame: PathPoint }[] = [];
	for (const line of layout.lines) {
		const placed = placeGlyphs(font, size, line.text);
		const total = placed.reduce((sum, glyph) => sum + glyph.advance, 0);
		const place = placement(line, total);
		if (!place) continue;
		const slots = placed.map((glyph) => ({ ...glyph, code: glyph.char }));
		const segments: PathSegment[] = [];
		for (const slot of slots) {
			const frame = place.frame(slot.x + slot.advance / 2);
			if (!frame) continue;
			glyphs.push({ slot, frame });
			segments.push(
				...glyphSegments(font, slot.glyph, size, slot.x, place.map, { bend: place.bend }),
			);
		}
		const bounds = cellBounds(slots, place, font.font.ascent * em, -font.font.descent * em);
		drawn.push({ segments, ...(bounds ? { bounds } : {}) });
	}
	const outline = layout.outline;
	const glow = layout.glow;
	const shadow = layout.shadow;
	const reflection = layout.reflection;
	if (reflection) {
		const bottom = Math.max(
			0,
			...drawn.map(({ bounds }) => (bounds ? bounds[3] : 0)),
		);
		const top = Math.min(
			bottom,
			...drawn.map(({ bounds }) => (bounds ? bounds[1] : bottom)),
		);
		const span = Math.max(1, bottom - top);
		const axisY = bottom + reflection.distance / 2;
		const bands = 4;
		const fill = rgb(layout.fill);
		for (let band = 0; band < bands; band++) {
			const t0 = band / bands;
			const t1 = (band + 1) / bands;
			const alpha =
				reflection.startOpacity +
				(reflection.endOpacity - reflection.startOpacity) * ((t0 + t1) / 2);
			if (alpha <= 0.005) continue;
			const bandTop = bottom + reflection.distance + t0 * span * Math.abs(reflection.scaleY);
			const bandHeight = (t1 - t0) * span * Math.abs(reflection.scaleY);
			page.save();
			page.rect(-1e5, bandTop, 2e5, bandHeight + 0.5);
			page.op('W n');
			page.alpha(alpha);
			page.concat([1, 0, 0, reflection.scaleY, 0, axisY * (1 - reflection.scaleY)]);
			page.fillColor(fill.r, fill.g, fill.b);
			for (const { segments } of drawn) {
				if (!segments.length) continue;
				for (const op of pathOperators(segments)) page.op(op);
				page.op('f');
			}
			page.restore();
		}
	}
	if (glow) {
		const parts: SilhouettePart[] = drawn.map(({ segments }) => ({
			segments,
			fill: 'nonzero',
			matrix: [1, 0, 0, 1, 0, 0],
			...(outline ? { stroke: outline.width } : {}),
		}));
		drawBlurredSilhouette(page, parts, glow.radius / 2, rgb(glow.color), glow.opacity);
	}
	if (shadow) {
		const parts: SilhouettePart[] = drawn.map(({ segments }) => ({
			segments,
			fill: 'nonzero',
			matrix: [1, 0, 0, 1, shadow.dx, shadow.dy],
			...(outline ? { stroke: outline.width } : {}),
		}));
		drawBlurredSilhouette(page, parts, shadow.blur / 2, rgb(shadow.color), shadow.opacity);
	}
	for (const { segments, bounds } of drawn) {
		if (!segments.length) continue;
		const ops = pathOperators(segments);
		// `paint-order: stroke fill`: the outline goes first, and the fill covers its inner half.
		if (outline) {
			const color = rgb(outline.color);
			page.save();
			page.strokeColor(color.r, color.g, color.b);
			page.lineStyle(outline.width, 0, 0);
			for (const op of ops) page.op(op);
			page.op('S');
			page.restore();
		}
		page.save();
		const gradient = layout.gradient && bounds;
		const fill = rgb(layout.fill);
		if (!gradient) page.fillColor(fill.r, fill.g, fill.b);
		for (const op of ops) page.op(op);
		if (gradient) {
			page.op('W n');
			paintGradient(page, layout, gradient);
		} else page.op('f');
		page.restore();
	}
	// The text itself, invisible over its outlines, glyph by glyph along the warp.
	for (const { slot, frame } of glyphs) {
		const cos = Math.cos(frame.angle);
		const sin = Math.sin(frame.angle);
		page.save();
		page.concat([cos, sin, -sin, cos, frame.x, frame.y]);
		page.embeddedText(font, size, -slot.advance / 2, 0, slot.char, { invisible: true });
		page.restore();
	}
}

function standardLines(
	page: PdfVectorPage,
	layout: WordArtLayout,
	report: DocxPdfReportBuilder,
): void {
	const base = pdfBaseFont(WORD_ART_FONT, true, false);
	const name = page.font(base);
	const size = layout.size;
	const lines: { slots: Slot[]; place: LinePlacement }[] = [];
	for (const line of layout.lines) {
		const slots: Slot[] = [];
		let x = 0;
		for (const char of line.text) {
			const normalized = char === '\t' || char === '\n' || char === '\r' ? ' ' : char;
			const encoded = encodeWinAnsi(normalized);
			if (encoded === undefined) report.missing(char);
			const code = encoded ?? '?';
			const advance = pdfTextWidth(code, base) * size;
			slots.push({ char, code, glyph: 0, x, advance });
			x += advance;
		}
		const place = placement(line, x);
		if (place) lines.push({ slots, place });
	}
	/** One text object setting every visible letter of a line with its own matrix. */
	const text = (slots: readonly Slot[], place: LinePlacement, mode: number, dx = 0, dy = 0) => {
		let ops = `BT /${name} ${size} Tf ${mode} Tr`;
		for (const slot of slots) {
			const frame = place.frame(slot.x + slot.advance / 2);
			if (!frame || !slot.code.trim()) continue;
			const cos = Math.cos(frame.angle);
			const sin = Math.sin(frame.angle);
			const ox = frame.x - (cos * slot.advance) / 2 + dx;
			const oy = frame.y - (sin * slot.advance) / 2 + dy;
			const m = [cos, sin, sin, -cos, ox, oy].map((v) => Number(v.toFixed(4)));
			ops += ` ${m.join(' ')} Tm ${pdfLiteral(slot.code)} Tj`;
		}
		page.op(`${ops} ET`);
	};
	const glow = layout.glow;
	const shadow = layout.shadow;
	const reflection = layout.reflection;
	if (reflection) {
		const bottom = Math.max(
			0,
			...lines.map(({ slots, place }) => {
				const b = cellBounds(slots, place, ASCENT * size, DESCENT * size);
				return b ? b[3] : 0;
			}),
		);
		const top = Math.min(
			bottom,
			...lines.map(({ slots, place }) => {
				const b = cellBounds(slots, place, ASCENT * size, DESCENT * size);
				return b ? b[1] : bottom;
			}),
		);
		const span = Math.max(1, bottom - top);
		const axisY = bottom + reflection.distance / 2;
		const fill = rgb(layout.fill);
		const bands = 4;
		page.op('/Span <</ActualText ()>> BDC');
		for (let band = 0; band < bands; band++) {
			const t0 = band / bands;
			const t1 = (band + 1) / bands;
			const alpha =
				reflection.startOpacity +
				(reflection.endOpacity - reflection.startOpacity) * ((t0 + t1) / 2);
			if (alpha <= 0.005) continue;
			const bandTop = bottom + reflection.distance + t0 * span * Math.abs(reflection.scaleY);
			const bandHeight = (t1 - t0) * span * Math.abs(reflection.scaleY);
			page.save();
			page.rect(-1e5, bandTop, 2e5, bandHeight + 0.5);
			page.op('W n');
			page.alpha(alpha);
			page.concat([1, 0, 0, reflection.scaleY, 0, axisY * (1 - reflection.scaleY)]);
			page.fillColor(fill.r, fill.g, fill.b);
			for (const { slots, place } of lines) text(slots, place, 0);
			page.restore();
		}
		page.endMarked();
	}
	if (glow) {
		const color = rgb(glow.color);
		wordArtSoftCopies(
			page,
			{ color: glow.color, dx: 0, dy: 0, blur: glow.radius, opacity: glow.opacity },
			(dx, dy) => {
				page.fillColor(color.r, color.g, color.b);
				for (const { slots, place } of lines) text(slots, place, 0, dx, dy);
			},
		);
	}
	if (shadow) {
		const color = rgb(shadow.color);
		wordArtSoftCopies(page, shadow, (dx, dy) => {
			page.fillColor(color.r, color.g, color.b);
			for (const { slots, place } of lines) text(slots, place, 0, dx, dy);
		});
	}
	const outline = layout.outline;
	const fill = rgb(layout.fill);
	for (const { slots, place } of lines) {
		page.save();
		page.fillColor(fill.r, fill.g, fill.b);
		if (outline) {
			const color = rgb(outline.color);
			page.strokeColor(color.r, color.g, color.b);
			page.lineStyle(outline.width, 0, 0);
		}
		const gradient = !!layout.gradient;
		// A gradient fills the letters' clip; with an outline, mode 6 strokes them too.
		text(slots, place, gradient ? (outline ? 6 : 7) : outline ? 2 : 0);
		const bounds = gradient && cellBounds(slots, place, ASCENT * size, DESCENT * size);
		if (bounds) paintGradient(page, layout, bounds);
		page.restore();
	}
}

/**
 * Draws a WordArt text box in its `width` x `height` box at (x, y), clipped to the box as Print
 * Layout's SVG is. Returns false when there is nothing to draw.
 */
export function paintWordArt(
	page: PdfVectorPage,
	text: DocxPdfText,
	report: DocxPdfReportBuilder,
	content: TextBoxContent,
	x: number,
	y: number,
	width: number,
	height: number,
): boolean {
	const style = content.wordArt;
	if (!style) return false;
	const paragraphs = content.blocks.flatMap((block) => (block.type === 'paragraph' ? [block] : []));
	const layout = wordArtLayout({
		style,
		lines: wordArtLines(content),
		width,
		height,
		fontPx: (paragraphs[0]?.runs[0]?.fontSize ?? WORD_ART_SIZE_PT) * (96 / 72),
	});
	if (!layout) return false;
	page.save();
	page.rect(x, y, width, height);
	page.op('W n');
	page.concat([1, 0, 0, 1, x, y]);
	const all = layout.lines.map((line) => line.text).join('');
	const font = text.fonts.pick(WORD_ART_FONT, true, false, all);
	if (font) {
		if (font.family.toLowerCase() === WORD_ART_FONT.toLowerCase()) report.embedded(font.family);
		else report.substituted(WORD_ART_FONT, font.family);
		outlineLines(page, font, layout);
	} else {
		report.substituted(WORD_ART_FONT, pdfBaseFont(WORD_ART_FONT, true, false).replace(/-.*$/, ''));
		if (layout.preset || layout.shadow) report.note(STANDARD_NOTE);
		standardLines(page, layout, report);
	}
	page.restore();
	return true;
}
