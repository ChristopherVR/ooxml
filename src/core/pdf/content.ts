/**
 * The content stream interpreter (ISO 32000-1 sections 8 and 9): runs a page's operators with
 * the graphics and text state and collects what the page shows as text spans, pictures and
 * rules, in reading coordinates. Shading, clipping and painting details are ignored.
 */
import type { PdfDocument } from './document';
import { PdfFont } from './fonts';
import { readImage } from './images';
import { EOF, isWhite, PdfKeyword, PdfLexer } from './lexer';
import {
	isName,
	isNumber,
	PdfName,
	PdfStream,
	PdfString,
	type PdfDict,
	type PdfValue,
} from './objects';
import type { PdfPageImage, PdfPageRule, PdfTextSpan } from './page-content';
import { indexOf } from './xref';

export type Matrix = [number, number, number, number, number, number];

export const multiply = (m: Matrix, n: Matrix): Matrix => [
	m[0] * n[0] + m[1] * n[2],
	m[0] * n[1] + m[1] * n[3],
	m[2] * n[0] + m[3] * n[2],
	m[2] * n[1] + m[3] * n[3],
	m[4] * n[0] + m[5] * n[2] + n[4],
	m[4] * n[1] + m[5] * n[3] + n[5],
];
export const apply = (m: Matrix, x: number, y: number): [number, number] => [
	m[0] * x + m[2] * y + m[4],
	m[1] * x + m[3] * y + m[5],
];

interface State {
	ctm: Matrix;
	fill: string;
	stroke: string;
	fillComponents: number;
	strokeComponents: number;
	lineWidth: number;
	charSpacing: number;
	wordSpacing: number;
	scale: number;
	leading: number;
	font: PdfFont | undefined;
	fontSize: number;
	render: number;
	rise: number;
}

const hex = (value: number) =>
	Math.round(Math.min(1, Math.max(0, value)) * 255)
		.toString(16)
		.padStart(2, '0');
function colorOf(values: number[]): string {
	if (values.length === 1) return `#${hex(values[0]!).repeat(3)}`;
	if (values.length === 3) return `#${hex(values[0]!)}${hex(values[1]!)}${hex(values[2]!)}`;
	if (values.length === 4) {
		const k = 1 - values[3]!;
		return `#${hex((1 - values[0]!) * k)}${hex((1 - values[1]!) * k)}${hex((1 - values[2]!) * k)}`;
	}
	return '#000000';
}

const MAX_RULES = 20_000;
const MAX_FORM_DEPTH = 12;

export interface ContentResult {
	spans: PdfTextSpan[];
	images: PdfPageImage[];
	rules: PdfPageRule[];
	notes: Set<string>;
}

export class ContentReader {
	readonly #fonts = new Map<PdfValue, PdfFont>();
	readonly result: ContentResult = { spans: [], images: [], rules: [], notes: new Set() };

	constructor(
		readonly document: PdfDocument,
		/** User space to reading space (crop box and page rotation). */
		readonly page: Matrix,
		readonly withImages: boolean,
	) {}

	async #font(resources: PdfDict, name: string): Promise<PdfFont | undefined> {
		const fonts = await this.document.dict(resources.get('Font'));
		const ref = fonts?.get(name);
		if (ref === undefined) return undefined;
		const cached = this.#fonts.get(ref);
		if (cached) return cached;
		const dict = await this.document.dict(ref);
		if (!dict) return undefined;
		const font = await PdfFont.load(this.document, dict);
		if (font.vertical) this.result.notes.add('Vertical text is read as horizontal text.');
		this.#fonts.set(ref, font);
		return font;
	}

	async run(
		data: Uint8Array,
		resources: PdfDict,
		ctm: Matrix,
		depth = 0,
		forms = new Set<PdfStream>(),
	): Promise<void> {
		const lexer = new PdfLexer(data);
		const stack: State[] = [];
		let state: State = {
			ctm,
			fill: '#000000',
			stroke: '#000000',
			fillComponents: 1,
			strokeComponents: 1,
			lineWidth: 1,
			charSpacing: 0,
			wordSpacing: 0,
			scale: 1,
			leading: 0,
			font: undefined,
			fontSize: 0,
			render: 0,
			rise: 0,
		};
		let tm: Matrix = [1, 0, 0, 1, 0, 0];
		let tlm: Matrix = [1, 0, 0, 1, 0, 0];
		let path: { kind: 'rect' | 'line'; points: number[] }[] = [];
		let pathComplex = false;
		let current: [number, number] | undefined;
		const operands: PdfValue[] = [];
		const num = (index: number) => {
			const value = operands[index];
			return isNumber(value) ? value : 0;
		};
		const show = (string: PdfString | PdfValue[]) => {
			const font = state.font;
			if (!font) return;
			const items = string instanceof PdfString ? [string] : string;
			const size = state.fontSize;
			const th = state.scale;
			let span: { text: string; start: Matrix; advance: number } | undefined;
			const flush = () => {
				if (span && span.text) this.#emit(span.text, span.start, span.advance, state, size);
				span = undefined;
			};
			for (const item of items) {
				if (isNumber(item)) {
					const shift = (-item / 1000) * size * th;
					if (Math.abs(item) > 200) flush();
					else if (span) span.advance += shift;
					tm = multiply([1, 0, 0, 1, shift, 0], tm);
					continue;
				}
				if (!(item instanceof PdfString)) continue;
				for (const glyph of font.glyphs(item.bytes)) {
					span ??= { text: '', start: tm, advance: 0 };
					const advance =
						(glyph.width * size +
							state.charSpacing +
							(glyph.bytes === 1 && glyph.code === 32 ? state.wordSpacing : 0)) *
						th;
					span.text += glyph.text;
					span.advance += advance;
					tm = multiply([1, 0, 0, 1, advance, 0], tm);
				}
			}
			flush();
		};
		const paint = (stroked: boolean, filled: boolean) => {
			if (!pathComplex && this.result.rules.length < MAX_RULES)
				for (const item of path) {
					const [x0, y0, x1, y1] = item.points as [number, number, number, number];
					const a = apply(multiply(state.ctm, this.page), x0, y0);
					const b = apply(multiply(state.ctm, this.page), x1, y1);
					const scale = Math.hypot(state.ctm[0], state.ctm[1]) || 1;
					this.result.rules.push({
						x0: Math.min(a[0], b[0]),
						y0: Math.min(a[1], b[1]),
						x1: Math.max(a[0], b[0]),
						y1: Math.max(a[1], b[1]),
						kind: item.kind,
						stroked,
						filled,
						color: filled ? state.fill : state.stroke,
						lineWidth: state.lineWidth * scale,
					});
				}
			path = [];
			pathComplex = false;
			current = undefined;
		};
		let guard = 0;
		for (;;) {
			if (++guard > 5_000_000) {
				this.result.notes.add('The page content is too long; the rest was not read.');
				break;
			}
			const token = lexer.object();
			if (token === EOF) break;
			if (!(token instanceof PdfKeyword)) {
				operands.push(token);
				continue;
			}
			const op = token.word;
			switch (op) {
				case 'q':
					if (stack.length < 64) stack.push({ ...state });
					break;
				case 'Q':
					state = stack.pop() ?? state;
					break;
				case 'cm':
					state.ctm = multiply([num(0), num(1), num(2), num(3), num(4), num(5)], state.ctm);
					break;
				case 'w':
					state.lineWidth = num(0);
					break;
				case 'g':
				case 'rg':
				case 'k':
					state.fill = colorOf(operands.filter(isNumber));
					break;
				case 'G':
				case 'RG':
				case 'K':
					state.stroke = colorOf(operands.filter(isNumber));
					break;
				case 'cs':
				case 'CS': {
					const name = operands[0];
					const components = isName(name)
						? name.name === 'DeviceRGB'
							? 3
							: name.name === 'DeviceCMYK'
								? 4
								: 1
						: 1;
					if (op === 'cs') state.fillComponents = components;
					else state.strokeComponents = components;
					break;
				}
				case 'sc':
				case 'scn':
				case 'SC':
				case 'SCN': {
					const values = operands.filter(isNumber);
					if (values.length === 1 || values.length === 3 || values.length === 4) {
						if (op.startsWith('s')) state.fill = colorOf(values);
						else state.stroke = colorOf(values);
					}
					break;
				}
				case 'BT':
					tm = [1, 0, 0, 1, 0, 0];
					tlm = tm;
					break;
				case 'Tc':
					state.charSpacing = num(0);
					break;
				case 'Tw':
					state.wordSpacing = num(0);
					break;
				case 'Tz':
					state.scale = num(0) / 100;
					break;
				case 'TL':
					state.leading = num(0);
					break;
				case 'Tr':
					state.render = num(0);
					break;
				case 'Ts':
					state.rise = num(0);
					break;
				case 'Tf': {
					const name = operands[0];
					state.fontSize = num(1);
					state.font = name instanceof PdfName ? await this.#font(resources, name.name) : undefined;
					break;
				}
				case 'Td':
				case 'TD':
					if (op === 'TD') state.leading = -num(1);
					tlm = multiply([1, 0, 0, 1, num(0), num(1)], tlm);
					tm = tlm;
					break;
				case 'Tm':
					tlm = [num(0), num(1), num(2), num(3), num(4), num(5)];
					tm = tlm;
					break;
				case 'T*':
					tlm = multiply([1, 0, 0, 1, 0, -state.leading], tlm);
					tm = tlm;
					break;
				case 'Tj':
				case 'TJ':
					if (operands[0] instanceof PdfString || Array.isArray(operands[0]))
						show(operands[0] as PdfString | PdfValue[]);
					break;
				case "'":
				case '"': {
					if (op === '"') {
						state.wordSpacing = num(0);
						state.charSpacing = num(1);
					}
					tlm = multiply([1, 0, 0, 1, 0, -state.leading], tlm);
					tm = tlm;
					const string = operands[operands.length - 1];
					if (string instanceof PdfString) show(string);
					break;
				}
				case 'm':
					current = [num(0), num(1)];
					break;
				case 'l':
					if (current)
						path.push({ kind: 'line', points: [current[0], current[1], num(0), num(1)] });
					current = [num(0), num(1)];
					break;
				case 're':
					path.push({ kind: 'rect', points: [num(0), num(1), num(0) + num(2), num(1) + num(3)] });
					break;
				case 'c':
				case 'v':
				case 'y':
					pathComplex = true;
					break;
				case 'S':
				case 's':
					paint(true, false);
					break;
				case 'f':
				case 'F':
				case 'f*':
					paint(false, true);
					break;
				case 'B':
				case 'B*':
				case 'b':
				case 'b*':
					paint(true, true);
					break;
				case 'n':
					path = [];
					pathComplex = false;
					break;
				case 'Do': {
					const name = operands[0];
					if (name instanceof PdfName)
						await this.#xobject(resources, name.name, state, depth, forms);
					break;
				}
				case 'BI':
					this.#inlineImage(lexer, state);
					break;
			}
			operands.length = 0;
		}
	}

	#emit(text: string, start: Matrix, advance: number, state: State, size: number): void {
		if (!text.trim() && !text.includes(' ')) return;
		const text2page = multiply(start, multiply(state.ctm, this.page));
		const trm = multiply([size * state.scale, 0, 0, size, 0, 0], text2page);
		// The baseline at zero rise; the rise (super- and subscripts) is kept apart.
		const [x, y] = apply(text2page, 0, 0);
		const [ex, ey] = apply(text2page, advance, 0);
		const [rx, ry] = apply(text2page, 0, state.rise);
		const span: PdfTextSpan = {
			text,
			x,
			y,
			width: Math.hypot(ex - x, ey - y),
			fontSize: Math.hypot(trm[2], trm[3]),
			font: state.font!.info,
			color: state.render === 1 ? state.stroke : state.fill,
			angle: Math.atan2(-trm[1], trm[0]) || 0,
			rise: Math.hypot(rx - x, ry - y) * Math.sign(state.rise),
		};
		if (state.render === 3 || state.render === 7) span.invisible = true;
		this.result.spans.push(span);
	}

	async #xobject(
		resources: PdfDict,
		name: string,
		state: State,
		depth: number,
		forms: Set<PdfStream>,
	): Promise<void> {
		const xobjects = await this.document.dict(resources.get('XObject'));
		const stream = await this.document.resolve(xobjects?.get(name));
		if (!(stream instanceof PdfStream)) return;
		const subtype = stream.dict.get('Subtype');
		if (isName(subtype, 'Form')) {
			if (depth >= MAX_FORM_DEPTH || forms.has(stream)) return;
			const matrixValue = await this.document.resolve(stream.dict.get('Matrix'));
			const matrix =
				Array.isArray(matrixValue) && matrixValue.length === 6 && matrixValue.every(isNumber)
					? (matrixValue as Matrix)
					: ([1, 0, 0, 1, 0, 0] as Matrix);
			const formResources = (await this.document.dict(stream.dict.get('Resources'))) ?? resources;
			let data: Uint8Array;
			try {
				data = await this.document.streamBytes(stream);
			} catch {
				this.result.notes.add('A drawing group of the page could not be read.');
				return;
			}
			forms.add(stream);
			await this.run(data, formResources, multiply(matrix, state.ctm), depth + 1, forms);
			forms.delete(stream);
			return;
		}
		if (!isName(subtype, 'Image') || !this.withImages) return;
		const box = this.#unitBox(state.ctm);
		try {
			this.result.images.push(await readImage(this.document, stream, box));
		} catch {
			this.result.images.push({
				...box,
				pixelWidth: 0,
				pixelHeight: 0,
				format: 'unsupported',
				reason: 'it could not be decoded',
			});
		}
	}

	/** Where the unit square lands on the page. */
	#unitBox(ctm: Matrix): { x: number; y: number; width: number; height: number } {
		const toPage = multiply(ctm, this.page);
		const corners = [
			apply(toPage, 0, 0),
			apply(toPage, 1, 0),
			apply(toPage, 0, 1),
			apply(toPage, 1, 1),
		];
		const xs = corners.map((point) => point[0]);
		const ys = corners.map((point) => point[1]);
		const x = Math.min(...xs);
		const y = Math.min(...ys);
		return { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y };
	}

	/** Skips an inline image (`BI ... ID data EI`), noting where it was. */
	#inlineImage(lexer: PdfLexer, state: State): void {
		const dict: PdfDict = new Map();
		for (;;) {
			const key = lexer.object();
			if (key === EOF) return;
			if (key instanceof PdfKeyword) {
				if (key.word === 'ID') break;
				continue;
			}
			if (key instanceof PdfName) {
				const value = lexer.object();
				if (value !== EOF && !(value instanceof PdfKeyword)) dict.set(key.name, value);
			}
		}
		const start = lexer.pos + 1;
		let end = start;
		for (;;) {
			end = indexOf(lexer.bytes, 'EI', end);
			if (end < 0) {
				lexer.pos = lexer.bytes.length;
				return;
			}
			if (
				isWhite(lexer.bytes[end - 1]) &&
				(end + 2 >= lexer.bytes.length || isWhite(lexer.bytes[end + 2]))
			)
				break;
			end += 2;
		}
		lexer.pos = end + 2;
		const box = this.#unitBox(state.ctm);
		if (this.withImages && box.width > 4 && box.height > 4)
			this.result.notes.add(
				'Pictures stored inside the page content (inline images) were left out.',
			);
	}
}
