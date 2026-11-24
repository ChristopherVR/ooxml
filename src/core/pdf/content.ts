/**
 * The content stream interpreter (ISO 32000-1 sections 8 and 9): runs a page's operators with
 * the graphics and text state and collects what the page shows as text spans, pictures and
 * rules and paths, in reading coordinates. Shadings and pattern fills are not drawn; clipping only
 * marks what it cuts.
 */
import { inlineImage, unitBox, xobjectImage } from './content-images';
import { PathBuilder, type ClipBox } from './content-paths';
import type { PdfDocument } from './document';
import { PdfFont } from './fonts';
import { EOF, PdfKeyword, PdfLexer } from './lexer';
import { currentMark, markedEntry, type MarkedEntry } from './marked';
import { apply, multiply, type Matrix } from './matrix';
import {
	isName,
	isNumber,
	PdfName,
	PdfStream,
	PdfString,
	type PdfDict,
	type PdfValue,
} from './objects';
import type { PdfPageImage, PdfPagePath, PdfPageRule, PdfTextSpan } from './page-content';

export type { Matrix };
export { apply, multiply };

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
	fillAlpha: number;
	strokeAlpha: number;
	dash: number[];
	fillPattern: boolean;
	strokePattern: boolean;
	clip: ClipBox | undefined;
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

/** The clip after intersecting `clip` with a new clipping path's bounds. */
function intersect(clip: ClipBox | undefined, next: ClipBox | undefined): ClipBox | undefined {
	if (!next) return clip;
	if (!clip) return next;
	return {
		x0: Math.max(clip.x0, next.x0),
		y0: Math.max(clip.y0, next.y0),
		x1: Math.min(clip.x1, next.x1),
		y1: Math.min(clip.y1, next.y1),
		rect: clip.rect && next.rect,
	};
}

const MAX_RULES = 20_000;
const MAX_PATHS = 2_000;
const MAX_FORM_DEPTH = 12;

export interface ContentResult {
	spans: PdfTextSpan[];
	images: PdfPageImage[];
	rules: PdfPageRule[];
	/** Painted paths that are not straight rules or rectangles (curves, diagonals, polygons). */
	paths: PdfPagePath[];
	notes: Set<string>;
	/** The marked-content id each span or picture was drawn under. */
	marks: Map<PdfTextSpan | PdfPageImage, number>;
	/** Drawings not returned as paths, by reason (shadings, patterns, clips, size, count). */
	drawings: Map<'shading' | 'pattern' | 'clip' | 'size' | 'count', number>;
}

export class ContentReader {
	readonly #fonts = new Map<PdfValue, PdfFont>();
	readonly #marked: MarkedEntry[] = [];
	readonly result: ContentResult = {
		spans: [],
		images: [],
		rules: [],
		paths: [],
		notes: new Set(),
		marks: new Map(),
		drawings: new Map(),
	};

	#leftOut(reason: 'shading' | 'pattern' | 'clip' | 'size' | 'count'): void {
		this.result.drawings.set(reason, (this.result.drawings.get(reason) ?? 0) + 1);
	}

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
		/** The graphics state a form XObject inherits from the content that draws it. */
		parent?: State,
	): Promise<void> {
		const lexer = new PdfLexer(data);
		const stack: State[] = [];
		let state: State = parent
			? { ...parent, ctm }
			: {
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
					fillAlpha: 1,
					strokeAlpha: 1,
					dash: [],
					fillPattern: false,
					strokePattern: false,
					clip: undefined,
				};
		let tm: Matrix = [1, 0, 0, 1, 0, 0];
		let tlm: Matrix = [1, 0, 0, 1, 0, 0];
		const path = new PathBuilder();
		let clipping = false;
		const markBase = this.#marked.length;
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
		const paint = (stroked: boolean, filled: boolean, evenOdd = false) => {
			// A clip set by `W` applies after the path that sets it is painted.
			const clip = clipping ? path.bounds(this.page, state.ctm) : undefined;
			const painted =
				stroked || filled ? path.paint(this.page, state, stroked, filled, evenOdd) : undefined;
			path.clear();
			if (clipping) {
				clipping = false;
				state.clip = intersect(state.clip, clip);
			}
			if (!painted) return;
			if ('rules' in painted) {
				const room = MAX_RULES - this.result.rules.length;
				this.result.rules.push(...painted.rules.slice(0, Math.max(0, room)));
			} else if ('unsupported' in painted) this.#leftOut(painted.unsupported);
			else if (this.result.paths.length < MAX_PATHS) this.result.paths.push(painted.path);
			else this.#leftOut('count');
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
				case 'd': {
					const array = operands[0];
					state.dash = Array.isArray(array)
						? array.filter(isNumber).filter((value) => value >= 0)
						: [];
					if (state.dash.every((value) => value === 0)) state.dash = [];
					break;
				}
				case 'gs': {
					const name = operands[0];
					if (name instanceof PdfName) await this.#extGState(resources, name.name, state);
					break;
				}
				case 'g':
				case 'rg':
				case 'k':
					state.fill = colorOf(operands.filter(isNumber));
					state.fillPattern = false;
					break;
				case 'G':
				case 'RG':
				case 'K':
					state.stroke = colorOf(operands.filter(isNumber));
					state.strokePattern = false;
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
					const pattern = isName(operands[operands.length - 1]);
					if (op.startsWith('s')) state.fillPattern = pattern;
					else state.strokePattern = pattern;
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
					path.moveTo(num(0), num(1));
					break;
				case 'l':
					path.lineTo(num(0), num(1));
					break;
				case 're':
					path.rect(num(0), num(1), num(2), num(3));
					break;
				case 'c':
					path.curveTo(num(0), num(1), num(2), num(3), num(4), num(5));
					break;
				case 'v': {
					const [x, y] = path.current ?? [num(0), num(1)];
					path.curveTo(x, y, num(0), num(1), num(2), num(3));
					break;
				}
				case 'y':
					path.curveTo(num(0), num(1), num(2), num(3), num(2), num(3));
					break;
				case 'h':
					path.close();
					break;
				case 'S':
					paint(true, false);
					break;
				case 's':
					path.close();
					paint(true, false);
					break;
				case 'f':
				case 'F':
				case 'f*':
					paint(false, true, op === 'f*');
					break;
				case 'B':
				case 'B*':
					paint(true, true, op === 'B*');
					break;
				case 'b':
				case 'b*':
					path.close();
					paint(true, true, op === 'b*');
					break;
				case 'n':
					paint(false, false);
					break;
				case 'W':
				case 'W*':
					clipping = true;
					break;
				case 'sh':
					this.#leftOut('shading');
					break;
				case 'BMC':
				case 'BDC':
					if (this.#marked.length < 256)
						this.#marked.push(await markedEntry(this.document, resources, operands));
					break;
				case 'EMC':
					if (this.#marked.length > markBase) this.#marked.pop();
					break;
				case 'Do': {
					const name = operands[0];
					if (name instanceof PdfName)
						await this.#xobject(resources, name.name, state, depth, forms);
					break;
				}
				case 'BI': {
					const box = unitBox(state.ctm, this.page);
					const image = await inlineImage(this.document, lexer, resources, box, {
						decode: this.withImages && box.width > 4 && box.height > 4,
						fill: state.fill,
					});
					if (image) this.#addImage(image);
					break;
				}
			}
			operands.length = 0;
		}
		this.#marked.length = Math.min(this.#marked.length, markBase);
	}

	/** Opacity and dashes from a graphics state parameter dictionary (`gs`). */
	async #extGState(resources: PdfDict, name: string, state: State): Promise<void> {
		const states = await this.document.dict(resources.get('ExtGState'));
		const dict = await this.document.dict(states?.get(name));
		if (!dict) return;
		const fill = await this.document.resolve(dict.get('ca'));
		const stroke = await this.document.resolve(dict.get('CA'));
		const width = await this.document.resolve(dict.get('LW'));
		if (isNumber(fill)) state.fillAlpha = Math.min(1, Math.max(0, fill));
		if (isNumber(stroke)) state.strokeAlpha = Math.min(1, Math.max(0, stroke));
		if (isNumber(width)) state.lineWidth = width;
	}

	#addImage(image: PdfPageImage): void {
		const mark = currentMark(this.#marked);
		if (mark.alt) image.alt = mark.alt;
		if (mark.mcid !== undefined) this.result.marks.set(image, mark.mcid);
		this.result.images.push(image);
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
		const mark = currentMark(this.#marked);
		if (mark.artifact) span.artifact = true;
		if (mark.mcid !== undefined) this.result.marks.set(span, mark.mcid);
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
			await this.run(data, formResources, multiply(matrix, state.ctm), depth + 1, forms, state);
			forms.delete(stream);
			return;
		}
		if (!isName(subtype, 'Image') || !this.withImages) return;
		const box = unitBox(state.ctm, this.page);
		this.#addImage(await xobjectImage(this.document, stream, box, { fill: state.fill }));
	}
}
