/**
 * Where a PDF draws each glyph, for comparing the PDF export with Word's Save as PDF to the glyph
 * (issue #382). It follows `cm`, `q`/`Q`, the text matrix and the text state operators and steps
 * the pen by each font's own advances (`Widths`, `W`) and the `TJ` offsets, as a viewer does, so
 * it reads Word's PDFs and the export's the same way. Test support only; it reads the fixtures.
 */
import { PdfDocument } from '../../../../pdf/document';
import { PdfFont } from '../../../../pdf/fonts';
import { EOF, PdfKeyword, PdfLexer } from '../../../../pdf/lexer';
import {
	PdfName,
	PdfStream,
	PdfString,
	isNumber,
	type PdfValue,
} from '../../../../pdf/objects';

/** One line of text: its glyphs left to right, in PDF points (y up from the page's bottom). */
export interface PdfGlyphLine {
	page: number;
	/** The baseline. */
	y: number;
	/** The size the line's first glyph is drawn at (`Tf` times the text and page matrices). */
	size: number;
	/** The font's PostScript name without its subset tag. */
	font: string;
	text: string;
	/** The left edge (origin) of every glyph that is not white space. */
	x: number[];
}

type Matrix = [number, number, number, number, number, number];

const multiply = (a: Matrix, b: Matrix): Matrix => [
	a[0] * b[0] + a[1] * b[2],
	a[0] * b[1] + a[1] * b[3],
	a[2] * b[0] + a[3] * b[2],
	a[2] * b[1] + a[3] * b[3],
	a[4] * b[0] + a[5] * b[2] + b[4],
	a[4] * b[1] + a[5] * b[3] + b[5],
];

const translate = (x: number, y: number): Matrix => [1, 0, 0, 1, x, y];

async function contentBytes(document: PdfDocument, value: PdfValue | undefined) {
	const resolved = await document.resolve(value);
	const parts: Uint8Array[] = [];
	for (const item of Array.isArray(resolved) ? resolved : [resolved]) {
		const stream = await document.resolve(item);
		if (stream instanceof PdfStream) parts.push(await document.streamBytes(stream));
	}
	return parts;
}

interface Glyph {
	page: number;
	char: string;
	x: number;
	y: number;
	size: number;
	font: string;
}

/** Every glyph a page draws with its origin and drawn size. */
async function pageGlyphs(document: PdfDocument, index: number): Promise<Glyph[]> {
	const node = document.pages[index]!;
	const fontDicts = await document.dict(node.resources.get('Font'));
	const fonts = new Map<string, PdfFont>();
	const glyphs: Glyph[] = [];
	for (const bytes of await contentBytes(document, node.dict.get('Contents'))) {
		const lexer = new PdfLexer(bytes);
		const stack: { ctm: Matrix; charSpacing: number; wordSpacing: number; scale: number }[] = [];
		let ctm: Matrix = [1, 0, 0, 1, 0, 0];
		let tm: Matrix = [1, 0, 0, 1, 0, 0];
		let line: Matrix = tm;
		let font: PdfFont | undefined;
		let fontName = '';
		let size = 0;
		let charSpacing = 0;
		let wordSpacing = 0;
		let scale = 1;
		let operands: (PdfValue | '[')[] = [];
		const show = (items: readonly (PdfValue | '[')[]) => {
			for (const item of items) {
				if (isNumber(item)) {
					tm = multiply(translate((-item / 1000) * size * scale, 0), tm);
					continue;
				}
				if (!(item instanceof PdfString) || !font) continue;
				for (const glyph of font.glyphs(item.bytes)) {
					const at = multiply(tm, ctm);
					glyphs.push({
						page: index,
						char: glyph.text,
						x: at[4],
						y: at[5],
						size: size * Math.hypot(at[0], at[1]),
						font: fontName,
					});
					const space = glyph.bytes === 1 && glyph.code === 32 ? wordSpacing : 0;
					tm = multiply(translate((glyph.width * size + charSpacing + space) * scale, 0), tm);
				}
			}
		};
		for (;;) {
			const token = lexer.token();
			if (token === EOF) break;
			if (token === ']') continue;
			if (!(token instanceof PdfKeyword)) {
				if (token === '[' || token === '<<' || token === '>>' || token === '{' || token === '}')
					continue;
				operands.push(token);
				continue;
			}
			const n = operands.filter(isNumber);
			switch (token.word) {
				// The text state's spacing and scale are part of the graphics state that q and Q keep.
				case 'q':
					stack.push({ ctm, charSpacing, wordSpacing, scale });
					break;
				case 'Q': {
					const saved = stack.pop();
					if (saved) ({ ctm, charSpacing, wordSpacing, scale } = saved);
					break;
				}
				case 'cm':
					ctm = multiply(n as Matrix, ctm);
					break;
				case 'BT':
					tm = line = [1, 0, 0, 1, 0, 0];
					break;
				case 'Tf': {
					const name = operands[0];
					size = n[0] ?? size;
					if (name instanceof PdfName) {
						let loaded = fonts.get(name.name);
						const dict = await document.dict(fontDicts?.get(name.name));
						if (!loaded && dict) {
							loaded = await PdfFont.load(document, dict);
							fonts.set(name.name, loaded);
						}
						font = loaded;
						fontName = loaded?.info.name ?? name.name;
					}
					break;
				}
				case 'Tc':
					charSpacing = n[0] ?? 0;
					break;
				case 'Tw':
					wordSpacing = n[0] ?? 0;
					break;
				case 'Tz':
					scale = (n[0] ?? 100) / 100;
					break;
				case 'Tm':
					tm = line = n as Matrix;
					break;
				case 'Td':
				case 'TD':
					tm = line = multiply(translate(n[0] ?? 0, n[1] ?? 0), line);
					break;
				case 'TJ':
				case 'Tj':
					show(operands);
					break;
			}
			operands = [];
		}
	}
	return glyphs;
}

/** The lines of every page of a PDF, top to bottom: glyphs grouped by baseline. */
export async function pdfGlyphLines(bytes: Uint8Array): Promise<PdfGlyphLine[]> {
	const document = await PdfDocument.open(bytes);
	const lines: PdfGlyphLine[] = [];
	for (let index = 0; index < document.pages.length; index++) {
		const byBaseline = new Map<number, Glyph[]>();
		for (const glyph of await pageGlyphs(document, index)) {
			if (!glyph.char.trim()) continue;
			const key = Math.round(glyph.y * 100);
			byBaseline.set(key, [...(byBaseline.get(key) ?? []), glyph]);
		}
		for (const [, glyphs] of [...byBaseline].sort(([a], [b]) => b - a)) {
			glyphs.sort((a, b) => a.x - b.x);
			const first = glyphs[0]!;
			lines.push({
				page: index,
				y: first.y,
				size: first.size,
				font: first.font,
				text: glyphs.map((glyph) => glyph.char).join(''),
				x: glyphs.map((glyph) => glyph.x),
			});
		}
	}
	return lines;
}
