import type { PdfVectorPage } from './pdf-vector';
import type { PdfEmbeddedFont, PdfFontSet } from './pdf-embedded-font';
import { encodeWinAnsi, pdfBaseFont, pdfTextWidth } from './pdf-fonts';
import { childElements } from './pdf-svg-dom';
import { inherited, number, parseColor, PdfFallback } from './pdf-svg-style';

interface Run {
	text: string;
	font: string;
	embedded?: PdfEmbeddedFont;
	size: number;
	color: { r: number; g: number; b: number; a: number };
	alpha: number;
	shift: number;
	underline: boolean;
	strike: boolean;
	letterSpacing: number;
	width: number;
}

const MAX_RUNS = 100_000;

/** Maps a `font-family` value to the family the PDF matches fonts by (default: as written). */
export type SvgFontFamilyMapper = (value: string) => string;

const asWritten: SvgFontFamilyMapper = (value) => value;

/** Baseline offset, in em, for a `dominant-baseline` other than the alphabetic default. */
function baselineShift(value: string | null): number {
	switch (value) {
		case 'middle':
		case 'central':
			return 0.35;
		case 'hanging':
		case 'text-before-edge':
		case 'text-top':
			return 0.8;
		case 'text-after-edge':
		case 'ideographic':
		case 'text-bottom':
			return -0.2;
		default:
			return 0;
	}
}

/**
 * Draw one SVG `<text>` as real PDF text, row by row (the renderer's `tspan` rows carry x, y
 * and anchor; their `tspan` children are the styled runs). The page space is y-down, so each
 * run uses a flipped text matrix. Positions inside a row use approximate standard-font widths.
 */
export function drawSvgText(
	page: PdfVectorPage,
	text: Element,
	alpha: number,
	fonts?: PdfFontSet,
	familyOf: SvgFontFamilyMapper = asWritten,
): void {
	const rows = childElements(text).filter((child) => child.localName === 'tspan');
	if (!rows.length) {
		if (text.textContent) drawRow(page, text, [text], alpha, fonts, familyOf);
		return;
	}
	let count = 0;
	for (const row of rows) {
		const spans = childElements(row).filter((child) => child.localName === 'tspan');
		count += spans.length || 1;
		if (count > MAX_RUNS) throw new PdfFallback('too much text');
		drawRow(page, row, spans.length ? spans : [row], alpha, fonts, familyOf);
	}
}

function drawRow(
	page: PdfVectorPage,
	row: Element,
	spans: Element[],
	alpha: number,
	fonts: PdfFontSet | undefined,
	familyOf: SvgFontFamilyMapper,
): void {
	let shift = 0;
	let total = 0;
	let chars = 0;
	let spaces = 0;
	const runs: Run[] = [];
	for (const span of spans) {
		const raw = span.textContent ?? '';
		shift += number(span.getAttribute('dy'));
		if (!raw) continue;
		const color = parseColor(inherited(span, 'fill') ?? '#000000');
		if (!color) continue;
		const size = number(inherited(span, 'font-size'), 12);
		const weight = inherited(span, 'font-weight') ?? '';
		const family = familyOf(inherited(span, 'font-family') ?? '');
		const bold = weight === 'bold' || number(weight) >= 600;
		const italic = (inherited(span, 'font-style') ?? '') === 'italic';
		const font = pdfBaseFont(family, bold, italic);
		const letterSpacing = number(span.getAttribute('letter-spacing'));
		const decoration = inherited(span, 'text-decoration') ?? '';
		// A supplied font of the requested family keeps the document's typeface; any supplied
		// font that covers the characters rescues text outside WinAnsi (and, for PDF/A, any text).
		const winAnsi = encodeWinAnsi(raw);
		const named = fonts?.prefers(family) ? fonts.pick(family, bold, italic, raw) : undefined;
		const embedded =
			named ?? (winAnsi === undefined ? fonts?.pick(family, bold, italic, raw) : undefined);
		if (winAnsi === undefined && !embedded)
			throw new PdfFallback('text outside the WinAnsi character set');
		const encoded = embedded ? raw : winAnsi!;
		const width =
			(embedded ? embedded.width(encoded) : pdfTextWidth(encoded, font)) * size +
			letterSpacing * [...encoded].length;
		total += width;
		chars += [...encoded].length;
		spaces += (encoded.match(/ /g) ?? []).length;
		runs.push({
			text: encoded,
			font,
			...(embedded ? { embedded } : {}),
			size,
			color,
			alpha: alpha * color.a * number(inherited(span, 'fill-opacity'), 1),
			shift,
			underline: decoration.includes('underline'),
			strike: decoration.includes('line-through'),
			letterSpacing,
			width,
		});
	}
	if (!runs.length) return;
	const wordSpacing = number(row.getAttribute('word-spacing'));
	const length = number(row.getAttribute('textLength'), 0);
	const distribute = length > 0 && chars > 0 ? (length - total) / chars : 0;
	total += wordSpacing * spaces + distribute * chars;
	const anchor = row.getAttribute('text-anchor');
	let x = number(row.getAttribute('x'), 0);
	const y =
		number(row.getAttribute('y'), 0) +
		baselineShift(inherited(row, 'dominant-baseline')) * runs[0]!.size;
	if (anchor === 'middle') x -= total / 2;
	else if (anchor === 'end') x -= total;
	for (const run of runs) {
		const baseline = y + run.shift;
		page.save();
		page.fillColor(run.color.r, run.color.g, run.color.b);
		if (run.alpha < 1) page.alpha(run.alpha);
		const spacing = { word: wordSpacing, character: run.letterSpacing + distribute };
		if (run.embedded) page.embeddedText(run.embedded, run.size, x, baseline, run.text, spacing);
		else page.text(run.font, run.size, x, baseline, run.text, spacing);
		const advance =
			run.width +
			wordSpacing * (run.text.match(/ /g) ?? []).length +
			distribute * [...run.text].length;
		const thickness = run.size * 0.06;
		const bars = [
			run.underline && baseline + run.size * 0.12,
			run.strike && baseline - run.size * 0.3,
		];
		for (const top of bars) {
			if (top === false) continue;
			page.rect(x, top, advance, thickness);
			page.op('f');
		}
		page.restore();
		x += advance;
	}
}
