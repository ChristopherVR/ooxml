/**
 * WordArt text effects on a text box's laid-out lines, as Print Layout paints them
 * (`paintWordArtEffects`): a 1 px outline (text render mode 2, fill then stroke) and the soft
 * outer shadow. The shadow is blurred from the glyph outlines of the embedded fonts the lines are
 * written in; text in a standard PDF font has no outlines to read, so its shadow is a ring of
 * faint offset copies, marked with an empty `ActualText` so the copies do not extract. DOM-free.
 */
import type { LayoutBlockBox, LayoutFragment, LayoutLine } from 'ooxml-core/docx/layout';
import type {
	WordArtGlowGeometry,
	WordArtReflectionGeometry,
	WordArtShadowGeometry,
} from 'ooxml-core/drawingml';
import { drawBlurredSilhouette, type SilhouettePart } from '../../export/pdf-blur-mask';
import type { PdfCidFont } from '../../export/pdf-cid-font';
import { PdfEmbeddedFont } from '../../export/pdf-embedded-font';
import { runOutlines } from '../../export/pdf-glyph-outlines';
import { PdfVectorPage, type PdfTextSpacing } from '../../export/pdf-vector';
import { paintBlock } from './pdf-blocks';
import type { DocxPdfPictures } from './pdf-pictures';
import { cssRgba, DocxPdfText } from './pdf-text';

const SHADOW_NOTE =
	'Text shadows in a standard PDF font are drawn as faint offset copies, since the font has no outlines to blur; supply the font file for a blurred shadow.';
const COPIES = 8;

type RecordedRun =
	| {
			kind: 'embedded';
			font: PdfCidFont;
			size: number;
			x: number;
			y: number;
			text: string;
			spacing: PdfTextSpacing;
	  }
	| {
			kind: 'standard';
			font: string;
			size: number;
			x: number;
			y: number;
			text: string;
			spacing: PdfTextSpacing;
	  };

/** A page that keeps the text runs drawn on it, for painting their shadows. */
class TextRecorder extends PdfVectorPage {
	readonly runs: RecordedRun[] = [];
	constructor() {
		super(1, 1);
	}
	override text(
		font: string,
		size: number,
		x: number,
		y: number,
		text: string,
		spacing: PdfTextSpacing = {},
	): void {
		this.runs.push({ kind: 'standard', font, size, x, y, text, spacing });
	}
	override embeddedText(
		font: PdfCidFont,
		size: number,
		x: number,
		y: number,
		text: string,
		spacing: PdfTextSpacing = {},
	): void {
		this.runs.push({ kind: 'embedded', font, size, x, y, text, spacing });
	}
}

/**
 * Draws `paint(dx, dy)` as a ring of faint copies around the shadow offset, together about as
 * dark as the shadow in the middle and fading over the blur: a soft shadow for text that cannot
 * be blurred. The copies carry an empty `ActualText`, so they never extract as text.
 */
export function wordArtSoftCopies(
	page: PdfVectorPage,
	shadow: WordArtShadowGeometry,
	paint: (dx: number, dy: number) => void,
): void {
	const sigma = shadow.blur / 2;
	const offsets: [number, number][] = [[shadow.dx, shadow.dy]];
	if (sigma > 0)
		for (let index = 0; index < COPIES; index++) {
			const angle = (index / COPIES) * 2 * Math.PI;
			offsets.push([shadow.dx + Math.cos(angle) * sigma, shadow.dy + Math.sin(angle) * sigma]);
		}
	const each = 1 - Math.pow(1 - Math.min(0.999, shadow.opacity), 1 / offsets.length);
	page.op('/Span <</ActualText ()>> BDC');
	page.save();
	page.alpha(each);
	for (const [dx, dy] of offsets) paint(dx, dy);
	page.restore();
	page.endMarked();
}

/**
 * Paints the shadow of the text `blocks` draw (in the current space) beneath them. Returns
 * whether some text had no outlines and was shadowed with copies.
 */
export function paintTextShadow(
	page: PdfVectorPage,
	text: DocxPdfText,
	pictures: DocxPdfPictures,
	blocks: readonly LayoutBlockBox[],
	shadow: WordArtShadowGeometry,
): boolean {
	const recorder = new TextRecorder();
	// Pictures in the lines draw onto the recorder too, where they are dropped.
	const painter = { page: recorder, text, pictures };
	for (const block of blocks) paintBlock(painter, block, 0, 0);
	const color = cssRgba(`#${shadow.color}`) ?? { r: 0, g: 0, b: 0, a: 1 };
	const parts: SilhouettePart[] = [];
	const standard: Extract<RecordedRun, { kind: 'standard' }>[] = [];
	for (const run of recorder.runs) {
		if (run.kind === 'standard') {
			if (run.text.trim()) standard.push(run);
		} else if (run.font instanceof PdfEmbeddedFont)
			parts.push({
				segments: runOutlines(run.font, run.size, run.x, run.y, run.text, run.spacing),
				fill: 'nonzero',
				matrix: [1, 0, 0, 1, shadow.dx, shadow.dy],
			});
	}
	if (parts.length) drawBlurredSilhouette(page, parts, shadow.blur / 2, color, shadow.opacity);
	if (!standard.length) return false;
	text.report.note(SHADOW_NOTE);
	wordArtSoftCopies(page, shadow, (dx, dy) => {
		page.fillColor(color.r, color.g, color.b);
		for (const run of standard)
			page.text(run.font, run.size, run.x + dx, run.y + dy, run.text, run.spacing);
	});
	return true;
}

/**
 * Paints the glow of the text `blocks` draw (in the current space) beneath them.
 */
export function paintTextGlow(
	page: PdfVectorPage,
	text: DocxPdfText,
	pictures: DocxPdfPictures,
	blocks: readonly LayoutBlockBox[],
	glow: WordArtGlowGeometry,
): boolean {
	return paintTextShadow(page, text, pictures, blocks, {
		color: glow.color,
		dx: 0,
		dy: 0,
		blur: glow.radius,
		opacity: glow.opacity,
	});
}

/**
 * Paints a vertically flipped, fading reflection of the text `blocks` draw beneath them.
 * Marked with an empty `ActualText` so reflected text never extracts.
 */
export function paintTextReflection(
	page: PdfVectorPage,
	text: DocxPdfText,
	pictures: DocxPdfPictures,
	blocks: readonly LayoutBlockBox[],
	reflection: WordArtReflectionGeometry,
): void {
	const recorder = new TextRecorder();
	const painter = { page: recorder, text, pictures };
	for (const block of blocks) paintBlock(painter, block, 0, 0);
	if (!recorder.runs.length) return;
	const maxY = Math.max(...recorder.runs.map((run) => run.y));
	const minY = Math.min(...recorder.runs.map((run) => run.y - run.size));
	const span = Math.max(1, maxY - minY);
	const axisY = maxY + reflection.distance / 2;
	const bands = 4;
	page.op('/Span <</ActualText ()>> BDC');
	for (let band = 0; band < bands; band++) {
		const t0 = band / bands;
		const t1 = (band + 1) / bands;
		const alpha =
			reflection.startOpacity + (reflection.endOpacity - reflection.startOpacity) * ((t0 + t1) / 2);
		if (alpha <= 0.005) continue;
		const top = maxY + reflection.distance + t0 * span * Math.abs(reflection.scaleY);
		const height = (t1 - t0) * span * Math.abs(reflection.scaleY);
		page.save();
		page.rect(-1e5, top, 2e5, height + 0.5);
		page.op('W n');
		page.alpha(alpha);
		page.concat([1, 0, 0, reflection.scaleY, 0, axisY * (1 - reflection.scaleY)]);
		for (const run of recorder.runs) {
			if (run.kind === 'embedded')
				page.embeddedText(run.font, run.size, run.x, run.y, run.text, run.spacing);
			else page.text(run.font, run.size, run.x, run.y, run.text, run.spacing);
		}
		page.restore();
	}
	page.endMarked();
}

/** Text drawn with an outline: fill, then a stroke of `width` in `color` (render mode 2). */
export class DocxPdfOutlineText extends DocxPdfText {
	constructor(
		base: DocxPdfText,
		readonly outline: { color: string; width: number },
	) {
		super(base.fonts, base.measurer, base.report);
	}
	override draw(
		page: PdfVectorPage,
		fragment: LayoutFragment,
		x: number,
		line: LayoutLine,
		lineTop: number,
		gapPx = 0,
	): void {
		const color = cssRgba(`#${this.outline.color}`);
		if (!color) {
			super.draw(page, fragment, x, line, lineTop, gapPx);
			return;
		}
		page.save();
		page.strokeColor(color.r, color.g, color.b);
		page.lineStyle(this.outline.width, 0, 0);
		// Text state is part of the graphics state, so the restore puts mode 0 back.
		page.op('2 Tr');
		super.draw(page, fragment, x, line, lineTop, gapPx);
		page.restore();
	}
}
