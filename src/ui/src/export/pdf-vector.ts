/**
 * A PDF 1.4 writer for vector pages: each page is a content stream plus the fonts (standard 14,
 * not embedded), constant-alpha states and image XObjects it names. Pages may also hold only a
 * JPEG picture, so one document can mix vector and raster pages. DOM-free; callers own the Blob.
 * Raster-only documents can use `pdf-pages.ts`.
 */
import { pdfLiteral } from './pdf-fonts';
import type { PdfEmbeddedFont } from './pdf-embedded-font';

/** Pixel data for one image XObject. `data` is the (optionally compressed) sample stream. */
export interface PdfImageResource {
	width: number;
	height: number;
	/** 8-bit RGB samples, a Flate stream of them, or JPEG bytes (`DCTDecode`). */
	data: Uint8Array;
	filter?: 'FlateDecode' | 'DCTDecode';
	/** Optional 8-bit alpha plane of the same pixel size. */
	alpha?: { data: Uint8Array; filter?: 'FlateDecode' };
}

/** An axial (linear) or radial gradient painted with the `sh` operator in the current space. */
export interface PdfShading {
	kind: 'axial' | 'radial';
	/** Axial: x0 y0 x1 y1. Radial: x0 y0 r0 x1 y1 r1. */
	coords: readonly number[];
	/** Colour stops with offsets in 0..1, in order. */
	stops: readonly { offset: number; r: number; g: number; b: number }[];
}

const num = (value: number): string => {
	if (!Number.isFinite(value)) return '0';
	const text = value.toFixed(6);
	return text.includes('.') ? text.replace(/\.?0+$/, '') || '0' : text;
};

/** One page under construction: append operators, then hand it to `buildVectorPdf`. */
export class PdfVectorPage {
	readonly #ops: string[] = [];
	readonly fonts = new Map<string, string>();
	readonly alphas = new Map<string, { fill: number; stroke: number }>();
	readonly images = new Map<string, PdfImageResource>();
	readonly embeddedFonts = new Map<string, PdfEmbeddedFont>();
	readonly shadings = new Map<string, PdfShading>();
	constructor(
		/** Page size in points. */
		readonly width: number,
		readonly height: number,
	) {
		if (!(width > 0 && height > 0) || !Number.isFinite(width + height))
			throw new Error('PDF page sizes must be positive and finite.');
	}
	get content(): string {
		return this.#ops.join('\n');
	}
	get operatorCount(): number {
		return this.#ops.length;
	}
	op(operator: string): void {
		this.#ops.push(operator);
	}
	save = () => this.op('q');
	restore = () => this.op('Q');
	concat(m: readonly number[]): void {
		this.op(`${m.map(num).join(' ')} cm`);
	}
	fillColor(r: number, g: number, b: number): void {
		this.op(`${num(r)} ${num(g)} ${num(b)} rg`);
	}
	strokeColor(r: number, g: number, b: number): void {
		this.op(`${num(r)} ${num(g)} ${num(b)} RG`);
	}
	lineStyle(width: number, cap: number, join: number, dash: readonly number[] = []): void {
		this.op(`${num(width)} w ${cap} J ${join} j [${dash.map(num).join(' ')}] 0 d`);
	}
	/** Constant fill and stroke opacity (1 emits nothing). */
	alpha(fill: number, stroke = fill): void {
		let name = [...this.alphas].find(([, v]) => v.fill === fill && v.stroke === stroke)?.[0];
		if (!name) {
			name = `GS${this.alphas.size}`;
			this.alphas.set(name, { fill, stroke });
		}
		this.op(`/${name} gs`);
	}
	rect(x: number, y: number, w: number, h: number): void {
		this.op(`${[x, y, w, h].map(num).join(' ')} re`);
	}
	font(baseFont: string): string {
		let name = [...this.fonts].find(([, base]) => base === baseFont)?.[0];
		if (!name) {
			name = `F${this.fonts.size}`;
			this.fonts.set(name, baseFont);
		}
		return name;
	}
	/**
	 * One run of text with a y-down text matrix, so glyphs read upright in a y-down space.
	 * `text` must already be a WinAnsi string.
	 */
	text(
		baseFont: string,
		size: number,
		x: number,
		y: number,
		text: string,
		spacing: { word?: number; character?: number } = {},
	): void {
		this.op(
			`BT /${this.font(baseFont)} ${num(size)} Tf 1 0 0 -1 ${num(x)} ${num(y)} Tm ${num(
				spacing.character ?? 0,
			)} Tc ${num(spacing.word ?? 0)} Tw ${pdfLiteral(text)} Tj ET`,
		);
	}
	/**
	 * Text in an embedded font (2-byte CIDs), y-down like `text`. Word spacing is emulated with
	 * kerning after each space, since `Tw` does not apply to multi-byte codes.
	 */
	embeddedText(
		font: PdfEmbeddedFont,
		size: number,
		x: number,
		y: number,
		text: string,
		spacing: { word?: number; character?: number } = {},
	): void {
		let name = [...this.embeddedFonts].find(([, existing]) => existing === font)?.[0];
		if (!name) {
			name = `E${this.embeddedFonts.size}`;
			this.embeddedFonts.set(name, font);
		}
		const word = spacing.word ? -(spacing.word * 1000) / size : 0;
		const parts: string[] = [];
		for (const piece of text.split(/( )/)) {
			if (!piece) continue;
			parts.push(`<${font.encode(piece)}>`);
			if (piece === ' ' && word) parts.push(num(word));
		}
		this.op(
			`BT /${name} ${num(size)} Tf 1 0 0 -1 ${num(x)} ${num(y)} Tm ${num(
				spacing.character ?? 0,
			)} Tc [${parts.join(' ')}] TJ ET`,
		);
	}
	/** Paint a gradient over the current clip, in the current space. */
	paintShading(shading: PdfShading): void {
		const name = `Sh${this.shadings.size}`;
		this.shadings.set(name, shading);
		this.op(`/${name} sh`);
	}
	image(resource: PdfImageResource): string {
		let name = [...this.images].find(([, existing]) => existing === resource)?.[0];
		if (!name) {
			name = `Im${this.images.size}`;
			this.images.set(name, resource);
		}
		return name;
	}
	/** Paint an image into the unit square of the current space (y up). */
	drawImage(resource: PdfImageResource): void {
		this.op(`/${this.image(resource)} Do`);
	}
}

export interface PdfVectorOptions {
	title?: string;
}

const latin1 = (text: string): Uint8Array => {
	const bytes = new Uint8Array(text.length);
	for (let index = 0; index < text.length; index++) bytes[index] = text.charCodeAt(index) & 0xff;
	return bytes;
};

function shadingDict(shading: PdfShading): string {
	const stops = shading.stops.length ? shading.stops : [{ offset: 0, r: 0, g: 0, b: 0 }];
	const colour = (s: { r: number; g: number; b: number }) =>
		`[${num(s.r)} ${num(s.g)} ${num(s.b)}]`;
	const pairs = stops.slice(1).map((stop, i) => {
		const from = stops[i]!;
		return `<< /FunctionType 2 /Domain [0 1] /C0 ${colour(from)} /C1 ${colour(stop)} /N 1 >>`;
	});
	const fn =
		pairs.length === 0
			? `<< /FunctionType 2 /Domain [0 1] /C0 ${colour(stops[0]!)} /C1 ${colour(stops[0]!)} /N 1 >>`
			: pairs.length === 1
				? pairs[0]!
				: `<< /FunctionType 3 /Domain [0 1] /Functions [${pairs.join(' ')}] /Bounds [${stops
						.slice(1, -1)
						.map((s) => num(Math.min(1, Math.max(0, s.offset))))
						.join(' ')}] /Encode [${pairs.map(() => '0 1').join(' ')}] >>`;
	return `<< /ShadingType ${shading.kind === 'axial' ? 2 : 3} /ColorSpace /DeviceRGB /Coords [${shading.coords.map(num).join(' ')}] /Function ${fn} /Extend [true true] >>`;
}

type ObjectWriter = (id: number, body: string) => void;
type StreamWriter = (id: number, dictionary: string, data: Uint8Array) => void;

/** The five objects of an embedded subset font starting at `id`; returns the next free id. */
function writeEmbeddedFont(
	font: PdfEmbeddedFont,
	id: number,
	object: ObjectWriter,
	stream: StreamWriter,
): number {
	const parts = font.build();
	const base = `${[...'ABCDEF'].map((_, i) => String.fromCharCode(65 + ((id * 7 + i * 5) % 26))).join('')}+${font.family.replace(/[^A-Za-z0-9]/g, '') || 'Font'}`;
	const [type0, cid, descriptor, file, unicode] = [id, id + 1, id + 2, id + 3, id + 4];
	const d = parts.descriptor;
	object(
		type0,
		`<< /Type /Font /Subtype /Type0 /BaseFont /${base} /Encoding /Identity-H /DescendantFonts [${cid} 0 R] /ToUnicode ${unicode} 0 R >>`,
	);
	object(
		cid,
		`<< /Type /Font /Subtype /CIDFontType2 /BaseFont /${base} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${descriptor} 0 R /CIDToGIDMap /Identity /DW 1000 /W [0 [${parts.widths.join(' ')}]] >>`,
	);
	object(
		descriptor,
		`<< /Type /FontDescriptor /FontName /${base} /Flags ${d.flags} /FontBBox [${d.bbox.join(' ')}] /ItalicAngle ${num(d.italicAngle)} /Ascent ${d.ascent} /Descent ${d.descent} /CapHeight ${d.capHeight} /StemV ${d.stemV} /FontFile2 ${file} 0 R >>`,
	);
	stream(file, `/Length1 ${parts.fontFile.length}`, parts.fontFile);
	stream(unicode, '', latin1(parts.toUnicode));
	return id + 5;
}

/** Assemble the PDF bytes for one or more pages. */
export function buildVectorPdf(
	pages: readonly PdfVectorPage[],
	options: PdfVectorOptions = {},
): Uint8Array<ArrayBuffer> {
	if (!pages.length) throw new Error('A PDF needs at least one page.');
	const chunks: Uint8Array[] = [];
	const offsets: number[] = [];
	let position = 0;
	const emit = (chunk: string | Uint8Array) => {
		const bytes = typeof chunk === 'string' ? latin1(chunk) : chunk;
		chunks.push(bytes);
		position += bytes.length;
	};
	const object = (id: number, body: string) => {
		offsets[id] = position;
		emit(`${id} 0 obj\n${body}\nendobj\n`);
	};
	const stream = (id: number, dictionary: string, data: Uint8Array) => {
		offsets[id] = position;
		emit(`${id} 0 obj\n<< ${dictionary} /Length ${data.length} >>\nstream\n`);
		emit(data);
		emit('\nendstream\nendobj\n');
	};
	emit('%PDF-1.4\n%\xe2\xe3\xcf\xd3\n');
	// 1 catalog, 2 page tree, 3 info; the rest from 4.
	let next = 4;
	const fontIds = new Map<string, number>();
	const embeddedIds = new Map<PdfEmbeddedFont, number>();
	const kids: number[] = [];
	for (const page of pages) {
		const xobjects: string[] = [];
		for (const [name, image] of page.images) {
			if (!(image.width > 0 && image.height > 0))
				throw new Error('PDF images need a positive pixel size.');
			let mask = '';
			if (image.alpha) {
				const maskId = next++;
				stream(
					maskId,
					`/Type /XObject /Subtype /Image /Width ${image.width} /Height ${image.height} /ColorSpace /DeviceGray /BitsPerComponent 8${image.alpha.filter ? ` /Filter /${image.alpha.filter}` : ''}`,
					image.alpha.data,
				);
				mask = ` /SMask ${maskId} 0 R`;
			}
			const id = next++;
			stream(
				id,
				`/Type /XObject /Subtype /Image /Width ${image.width} /Height ${image.height} /ColorSpace /DeviceRGB /BitsPerComponent 8${image.filter ? ` /Filter /${image.filter}` : ''}${mask}`,
				image.data,
			);
			xobjects.push(`/${name} ${id} 0 R`);
		}
		const fontEntries: string[] = [];
		for (const [name, base] of page.fonts) {
			let id = fontIds.get(base);
			if (!id) {
				id = next++;
				fontIds.set(base, id);
				object(
					id,
					`<< /Type /Font /Subtype /Type1 /BaseFont /${base} /Encoding /WinAnsiEncoding >>`,
				);
			}
			fontEntries.push(`/${name} ${id} 0 R`);
		}
		for (const [name, font] of page.embeddedFonts) {
			let id = embeddedIds.get(font);
			if (!id) {
				id = next;
				embeddedIds.set(font, id);
				next = writeEmbeddedFont(font, id, object, stream);
			}
			fontEntries.push(`/${name} ${id} 0 R`);
		}
		const shadings = [...page.shadings].map(
			([name, shading]) => `/${name} ${shadingDict(shading)}`,
		);
		const states = [...page.alphas].map(
			([name, v]) => `/${name} << /ca ${num(v.fill)} /CA ${num(v.stroke)} >>`,
		);
		const content = latin1(page.content);
		const contentId = next++;
		stream(contentId, '', content);
		const pageId = next++;
		kids.push(pageId);
		object(
			pageId,
			`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${num(page.width)} ${num(page.height)}] /Contents ${contentId} 0 R /Resources << /ProcSet [/PDF /Text /ImageC] /Font << ${fontEntries.join(' ')} >> /XObject << ${xobjects.join(' ')} >> /ExtGState << ${states.join(' ')} >> /Shading << ${shadings.join(' ')} >> >> >>`,
		);
	}
	object(1, '<< /Type /Catalog /Pages 2 0 R >>');
	object(
		2,
		`<< /Type /Pages /Kids [${kids.map((id) => `${id} 0 R`).join(' ')}] /Count ${kids.length} >>`,
	);
	object(
		3,
		`<< /Producer (ooxml-ui)${options.title ? ` /Title ${pdfLiteral(options.title.replace(/[^\x20-\x7e]/g, '?'))}` : ''} >>`,
	);
	const xref = position;
	let table = `xref\n0 ${next}\n0000000000 65535 f \n`;
	for (let id = 1; id < next; id++)
		table += `${String(offsets[id] ?? 0).padStart(10, '0')} 00000 n \n`;
	emit(`${table}trailer\n<< /Size ${next} /Root 1 0 R /Info 3 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
	const result = new Uint8Array(new ArrayBuffer(position));
	let offset = 0;
	for (const chunk of chunks) {
		result.set(chunk, offset);
		offset += chunk.length;
	}
	return result;
}
