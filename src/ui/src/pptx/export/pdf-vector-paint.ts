/**
 * Paint a slide display list (`pdf-vector-display-list.ts`) onto one PDF page of the shared
 * vector writer: paths, gradient shadings, pictures, inline SVG through the shared converter,
 * and text as real PDF text. Text uses a supplied (embedded, subset) font of the run's family
 * when there is one, else a standard PDF font; each fragment is scaled horizontally to its
 * measured width, so positions follow the slide even where the glyph shapes differ. Text no font
 * can write (drawn as a picture by the collector) is kept as an invisible text layer in a glyphless
 * font (`pdf-text-layer-font.ts`), so it stays searchable. Pure.
 */
import type { PdfFontSet } from '../../export/pdf-embedded-font';
import { pdfTextWidth } from '../../export/pdf-fonts';
import { pathOperators, parsePath } from '../../export/pdf-path';
import { drawSvgSubtree } from '../../export/pdf-svg';
import { PdfTextLayerFont } from '../../export/pdf-text-layer-font';
import { PdfVectorPage } from '../../export/pdf-vector';
import { PdfDiagnosticSet, type PdfExportDiagnostic } from './pdf-export-diagnostics';
import type { DisplayItem, PdfClip, TextItem } from './pdf-vector-display-list';
import { primaryFamily, resolvePdfTextFont } from './pdf-text-font';

export {
	canWritePdfText,
	primaryFamily,
	resolvePdfTextFont,
	type PdfTextFont,
} from './pdf-text-font';

export interface PaintSlideOptions {
	/** The stage size in CSS px. */
	canvasWidth: number;
	canvasHeight: number;
	/** The page size in points. */
	pageWidth: number;
	pageHeight: number;
	fonts?: PdfFontSet;
}

export interface PaintedSlide {
	page: PdfVectorPage;
	diagnostics: PdfExportDiagnostic[];
}

/** Concatenate the clip's path in its own space and intersect the clip with it. */
function clipTo(page: PdfVectorPage, clip: PdfClip): void {
	const segments = parsePath(clip.d);
	if (!segments) return;
	const [a, b, c, d, e, f] = clip.matrix;
	const det = a * d - b * c;
	if (!segments.length || Math.abs(det) < 1e-12) {
		// An empty or degenerate clip shows nothing.
		page.rect(0, 0, 0, 0);
		page.op('W n');
		return;
	}
	page.concat(clip.matrix);
	for (const operator of pathOperators(segments)) page.op(operator);
	page.op(clip.evenOdd ? 'W* n' : 'W n');
	// Undo the clip's matrix so the item's own matrix applies from stage space.
	page.concat([d / det, -b / det, -c / det, a / det, (c * f - d * e) / det, (b * e - a * f) / det]);
}

/** The fonts one page paints with: the supplied set and the invisible text layer's font. */
interface PaintFonts {
	set: PdfFontSet | undefined;
	textLayer: PdfTextLayerFont;
}

/** `item` as invisible text in the glyphless text-layer font, fitted to its measured width. */
function paintTextLayer(page: PdfVectorPage, item: TextItem, font: PdfTextLayerFont): void {
	const natural = font.width(item.text) * item.size;
	const scale =
		natural > 0 && item.width > 0 ? Math.min(400, Math.max(25, (item.width / natural) * 100)) : 100;
	page.embeddedText(font, item.size, item.x, item.baseline, item.text, {
		scale,
		invisible: true,
	});
}

function paintText(
	page: PdfVectorPage,
	item: TextItem,
	fonts: PaintFonts,
	notes: PdfDiagnosticSet,
): void {
	const font = resolvePdfTextFont(item.text, item.family, item.bold, item.italic, fonts.set);
	if (!font) {
		// No font has these characters. The collector draws such text as a picture with this
		// invisible layer over it; visible text that still reaches here is not drawn.
		if (!item.invisible) notes.add('unsearchableText');
		paintTextLayer(page, item, fonts.textLayer);
		return;
	}
	// An empty family is the default font; the message says so in the viewer's language.
	const family = primaryFamily(item.family) || '';
	if (font.kind === 'standard' && !item.invisible)
		notes.add('fontStandard', { family, substitute: font.baseFont.split('-')[0] ?? '' });
	else if (font.kind === 'embedded' && !font.named && !item.invisible)
		notes.add('fontSubstituted', { family, substitute: font.font.family });
	const natural =
		(font.kind === 'embedded'
			? font.font.width(item.text)
			: pdfTextWidth(font.encoded, font.baseFont)) * item.size;
	const scale =
		natural > 0 && item.width > 0 ? Math.min(400, Math.max(25, (item.width / natural) * 100)) : 100;
	const spacing = { scale, ...(item.invisible ? { invisible: true } : {}) };
	page.fillColor(item.color.r, item.color.g, item.color.b);
	const alpha = item.alpha * item.color.a;
	if (alpha < 1) page.alpha(alpha);
	if (font.kind === 'embedded')
		page.embeddedText(font.font, item.size, item.x, item.baseline, item.text, spacing);
	else page.text(font.baseFont, item.size, item.x, item.baseline, font.encoded, spacing);
	if (item.invisible) return;
	const thickness = Math.max(0.5, item.size * 0.06);
	if (item.underline) {
		page.rect(item.x, item.baseline + item.size * 0.12, item.width, thickness);
		page.op('f');
	}
	if (item.strike) {
		page.rect(item.x, item.baseline - item.size * 0.3, item.width, thickness);
		page.op('f');
	}
}

function paintItem(
	page: PdfVectorPage,
	item: DisplayItem,
	fonts: PaintFonts,
	notes: PdfDiagnosticSet,
): void {
	page.save();
	for (const clip of item.clips) clipTo(page, clip);
	page.concat(item.matrix);
	switch (item.kind) {
		case 'path': {
			const segments = parsePath(item.d);
			if (!segments?.length) break;
			const evenOdd = item.evenOdd ? '*' : '';
			if (item.shading) {
				page.save();
				const alpha = item.alpha * item.shading.opacity;
				if (alpha < 1) page.alpha(alpha, 1);
				for (const operator of pathOperators(segments)) page.op(operator);
				page.op(`W${evenOdd} n`);
				for (const matrix of item.shading.matrices) page.concat(matrix);
				page.paintShading(item.shading.shading);
				page.restore();
			}
			const fill = item.fill && item.fill.a > 0 ? item.fill : undefined;
			const stroke = item.stroke && item.stroke.color.a > 0 && item.stroke.width > 0;
			if (!fill && !stroke) break;
			const fillAlpha = fill ? item.alpha * fill.a : 1;
			const strokeAlpha = item.stroke ? item.alpha * item.stroke.color.a : 1;
			if (fillAlpha < 1 || strokeAlpha < 1) page.alpha(fillAlpha, strokeAlpha);
			if (fill) page.fillColor(fill.r, fill.g, fill.b);
			if (stroke && item.stroke) {
				const s = item.stroke;
				page.strokeColor(s.color.r, s.color.g, s.color.b);
				page.lineStyle(s.width, s.cap, s.join, s.dash);
			}
			for (const operator of pathOperators(segments)) page.op(operator);
			page.op(fill && stroke ? `B${evenOdd}` : fill ? `f${evenOdd}` : 'S');
			break;
		}
		case 'image':
			if (!(item.width > 0 && item.height > 0)) break;
			if (item.alpha < 1) page.alpha(item.alpha);
			page.concat([item.width, 0, 0, -item.height, item.x, item.y + item.height]);
			page.drawImage(item.image);
			break;
		case 'text':
			paintText(page, item, fonts, notes);
			break;
		case 'svg':
			try {
				for (const note of drawSvgSubtree(page, item.svg, item.images, fonts.set, item.alpha))
					notes.add(note);
			} catch {
				notes.add('drawingPartMissing');
			}
			break;
	}
	page.restore();
}

/** Paint `items` in order on a new page; the diagnostics say what was approximated. */
export function paintSlide(
	items: readonly DisplayItem[],
	options: PaintSlideOptions,
): PaintedSlide {
	const page = new PdfVectorPage(options.pageWidth, options.pageHeight);
	const notes = new PdfDiagnosticSet();
	// Stage px, y down, to PDF points, origin bottom left.
	const sx = options.pageWidth / options.canvasWidth;
	const sy = options.pageHeight / options.canvasHeight;
	page.concat([sx, 0, 0, -sy, 0, options.pageHeight]);
	// The stage clips everything to the slide.
	page.rect(0, 0, options.canvasWidth, options.canvasHeight);
	page.op('W n');
	const fonts = {
		set: options.fonts,
		textLayer: options.fonts?.textLayer ?? new PdfTextLayerFont(),
	};
	for (const item of items) paintItem(page, item, fonts, notes);
	return { page, diagnostics: notes.toArray() };
}
