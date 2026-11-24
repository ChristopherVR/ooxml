/**
 * The JBIG2Decode filter (ISO 32000-1 section 7.4.7; ITU-T T.88 / ISO/IEC 14492): decodes the
 * embedded stream of a PDF image (with its /JBIG2Globals segments) or a JBIG2 file into pages.
 * Covers page information, end of stripe (pages of unknown height), symbol dictionaries, text
 * regions, generic regions (arithmetic with TPGDON, or MMR), generic refinement regions, pattern
 * dictionaries, halftone regions and user Huffman tables, intermediate and immediate. Extended
 * generic templates and colour extensions (T.88 amendments) are refused with a reason. Every
 * bitmap and every composition is charged to a budget, so damaged or hostile data stops early.
 */
import { Bitmap, compose, crop, Jbig2Budget, Jbig2Error, packInverted } from './jbig2-bitmap';
import { decodeHalftone, decodePatternDictionary, readHalftoneHeader } from './jbig2-halftone';
import { readTableSegment, type HuffmanTable } from './jbig2-huffman';
import { genericRegion, refinementRegion, textRegion } from './jbig2-regions';
import { readRegionInfo, readSegments, u32, type RegionInfo, type Segment } from './jbig2-segments';
import { decodeSymbolDictionary, type SymbolDictionary } from './jbig2-symbols';

export { Jbig2Error } from './jbig2-bitmap';

export interface Jbig2Options {
	/** The largest bitmap (pixels); default 40 million. */
	maxPixels?: number;
	/** Pixel operations allowed for the whole stream; default 400 million. */
	maxWork?: number;
}

type Result =
	| { kind: 'symbols'; dictionary: SymbolDictionary }
	| { kind: 'patterns'; patterns: Bitmap[] }
	| { kind: 'table'; table: HuffmanTable }
	| { kind: 'region'; info: RegionInfo; bitmap: Bitmap };

interface Page {
	number: number;
	bitmap: Bitmap;
	defaultPixel: number;
	/** The height is unknown until the end (striped pages, 7.4.8.2). */
	growing: boolean;
	/** Rows known to belong to a growing page. */
	extent: number;
}

class Decoder {
	readonly results = new Map<number, Result>();
	readonly pages: Page[] = [];
	constructor(readonly budget: Jbig2Budget) {}

	run(data: Uint8Array, segments: Segment[]): void {
		for (const segment of segments) this.segment(data, segment);
	}

	private referred<K extends Result['kind']>(
		segment: Segment,
		kind: K,
	): Extract<Result, { kind: K }>[] {
		return segment.referred.flatMap((number) => {
			const result = this.results.get(number);
			return result?.kind === kind ? [result as Extract<Result, { kind: K }>] : [];
		});
	}

	private page(segment: Segment): Page {
		const page =
			[...this.pages].reverse().find((item) => item.number === segment.page) ?? this.pages.at(-1);
		if (!page) throw new Jbig2Error('A JBIG2 region comes before the page information.');
		return page;
	}

	private segment(data: Uint8Array, segment: Segment): void {
		const { type, start, end } = segment;
		const budget = this.budget;
		if (type === 0) {
			const inputs = this.referred(segment, 'symbols');
			const dictionary = decodeSymbolDictionary(
				data,
				start,
				end,
				inputs.flatMap((item) => item.dictionary.exported),
				this.referred(segment, 'table').map((item) => item.table),
				inputs.at(-1)?.dictionary,
				budget,
			);
			this.results.set(segment.number, { kind: 'symbols', dictionary });
		} else if (type === 4 || type === 6 || type === 7) {
			const symbols = this.referred(segment, 'symbols').flatMap((item) => item.dictionary.exported);
			const tables = this.referred(segment, 'table').map((item) => item.table);
			this.region(segment, textRegion(data, segment, symbols, tables, budget));
		} else if (type === 16) {
			this.results.set(segment.number, {
				kind: 'patterns',
				patterns: decodePatternDictionary(data, start, end, budget),
			});
		} else if (type === 20 || type === 22 || type === 23) {
			const info = readRegionInfo(data, start, end);
			const header = readHalftoneHeader(data, start + 17, end, info.width, info.height);
			const patterns = this.referred(segment, 'patterns')[0]?.patterns;
			if (!patterns) throw new Jbig2Error('A JBIG2 halftone region has no pattern dictionary.');
			const bitmap = decodeHalftone(header, patterns, data, header.start, end, budget);
			this.region(segment, { info, bitmap });
		} else if (type === 36 || type === 38 || type === 39) {
			this.region(segment, genericRegion(data, segment, budget));
		} else if (type === 40 || type === 42 || type === 43) {
			const intermediate = this.referred(segment, 'region')[0];
			const result = refinementRegion(
				data,
				segment,
				(info) => {
					if (intermediate) return intermediate.bitmap;
					const page = this.page(segment);
					this.grow(page, info.y + info.height);
					return crop(page.bitmap, info.x, info.y, info.width, info.height, budget);
				},
				budget,
			);
			this.region(segment, result);
		} else if (type === 48) this.pageInformation(data, segment);
		else if (type === 50) {
			const page = this.page(segment);
			if (page.growing) this.grow(page, u32(data, start) + 1);
		} else if (type === 53) {
			this.results.set(segment.number, {
				kind: 'table',
				table: readTableSegment(data, start, end),
			});
		}
		// End of page (49), end of file (51), profiles (52), extensions (62) and reserved types
		// carry nothing this decoder draws.
	}

	private pageInformation(data: Uint8Array, segment: Segment): void {
		const { start, end } = segment;
		if (end - start < 19) throw new Jbig2Error('A JBIG2 page information segment is too short.');
		const width = u32(data, start);
		const height = u32(data, start + 4);
		const flags = data[start + 16]!;
		const defaultPixel = (flags >> 2) & 1;
		const growing = height === 0xffffffff;
		this.pages.push({
			number: segment.page,
			bitmap: new Bitmap(width, growing ? 0 : height, this.budget, defaultPixel),
			defaultPixel,
			growing,
			extent: 0,
		});
	}

	/** Makes a growing page at least `rows` tall. */
	private grow(page: Page, rows: number): void {
		if (!page.growing || rows <= page.extent) return;
		page.extent = rows;
		if (rows <= page.bitmap.height) return;
		const height = Math.max(
			rows,
			Math.min(
				page.bitmap.height * 2,
				Math.floor(this.budget.maxPixels / Math.max(1, page.bitmap.width)),
			),
		);
		const bitmap = new Bitmap(page.bitmap.width, height, this.budget, page.defaultPixel);
		bitmap.data.set(page.bitmap.data);
		page.bitmap = bitmap;
	}

	private region(segment: Segment, result: { info: RegionInfo; bitmap: Bitmap }): void {
		// Intermediate regions (types 4, 20, 36, 40) are kept for a refinement to refer to.
		if (segment.type === 4 || segment.type === 20 || segment.type === 36 || segment.type === 40) {
			this.results.set(segment.number, { kind: 'region', ...result });
			return;
		}
		const page = this.page(segment);
		const { info, bitmap } = result;
		this.grow(page, info.y + bitmap.height);
		compose(page.bitmap, bitmap, info.x, info.y, info.operator, this.budget);
	}

	/** The finished pages (growing ones cut to their extent). */
	finished(): Bitmap[] {
		return this.pages.map((page) =>
			page.growing
				? crop(page.bitmap, 0, 0, page.bitmap.width, page.extent, this.budget)
				: page.bitmap,
		);
	}
}

const budgetFor = (options: Jbig2Options): Jbig2Budget => {
	const maxPixels = options.maxPixels ?? 40_000_000;
	return new Jbig2Budget(maxPixels, 2 * maxPixels + 16_000_000, options.maxWork ?? 400_000_000);
};

/** Decodes every page of a JBIG2 file or embedded stream (with optional global segments). */
export function decodeJbig2Pages(
	data: Uint8Array,
	globals?: Uint8Array,
	options: Jbig2Options = {},
): Bitmap[] {
	const decoder = new Decoder(budgetFor(options));
	if (globals?.length) decoder.run(globals, readSegments(globals));
	decoder.run(data, readSegments(data));
	return decoder.finished();
}

/**
 * Decodes a PDF image's JBIG2 stream to `width` x `height` packed rows, 1 for white, as the
 * JBIG2Decode filter delivers them (the first page; missing rows and columns are white).
 */
export function decodeJbig2(
	data: Uint8Array,
	globals: Uint8Array | undefined,
	width: number,
	height: number,
	options: Jbig2Options = {},
): Uint8Array {
	const page = decodeJbig2Pages(data, globals, options)[0];
	if (!page) throw new Jbig2Error('The JBIG2 data has no page.');
	return packInverted(page, width, height);
}
