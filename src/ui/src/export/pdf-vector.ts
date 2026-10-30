/**
 * A PDF 1.4 writer for vector pages: each page is a content stream plus the fonts (standard 14,
 * not embedded), constant-alpha states and image XObjects it names. Pages may also hold only a
 * JPEG picture, so one document can mix vector and raster pages. DOM-free; callers own the Blob.
 * Raster-only documents can use `pdf-pages.ts`.
 */
import { pdfLiteral } from './pdf-fonts';

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
			`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${num(page.width)} ${num(page.height)}] /Contents ${contentId} 0 R /Resources << /ProcSet [/PDF /Text /ImageC] /Font << ${fontEntries.join(' ')} >> /XObject << ${xobjects.join(' ')} >> /ExtGState << ${states.join(' ')} >> >> >>`,
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
