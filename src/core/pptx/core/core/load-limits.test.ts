import JSZip from 'jszip';
import { describe, expect, it, vi } from 'vitest';

import { declaredEntrySize } from '../../../opc/zip-limits';
import { serializePptxToJson } from '../../converter/json';
import { createShapeElement, createTextElement } from '../builders/sdk';
import { PptxHandler } from '../PptxHandler';
import type { PptxElement } from '../types';
import { flattenElementsDeep } from '../utils/flatten-elements';
import {
	DEFAULT_PPTX_LOAD_LIMITS,
	getPptxLoadLimitViolation,
	PptxLoadLimitError,
	resolvePptxLoadLimits,
} from './load-limits';
import type { PptxLoadLimits } from './load-limits';
import { DEFAULT_MAX_UNCOMPRESSED_BYTES, MAX_ZIP_ENTRY_COUNT, ZipBombError } from './types';

const SLIDES = 3;

/** A deck of {@link SLIDES} slides: two text boxes each, and a group of two on the last one. */
async function buildDeck(): Promise<ArrayBuffer> {
	const { handler, data, createSlide } = await PptxHandler.createBlank({
		initialSlideCount: 0,
	});
	for (let index = 0; index < SLIDES; index++) {
		const slide = createSlide('Blank')
			.addText(`Title ${index + 1}`, { x: 40, y: 40, width: 400, height: 40 })
			.addText(`Body ${index + 1}`, { x: 40, y: 120, width: 400, height: 200 });
		if (index === SLIDES - 1) {
			slide.addGroup(
				[
					createTextElement('Grouped', { x: 0, y: 0, width: 100, height: 30 }),
					createShapeElement('rect', { x: 0, y: 40, width: 100, height: 60 }),
				],
				{ x: 500, y: 100, width: 200, height: 150 },
			);
		}
		data.slides.push(slide.build());
	}
	const bytes = await handler.save(data.slides);
	handler.dispose();
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

const deck = buildDeck();

/** Elements a slide authored itself (inherited layout and master elements carry a prefix). */
function authored(elements: readonly PptxElement[]): number {
	return flattenElementsDeep(elements).filter((element) => !/^(?:layout|master)-/u.test(element.id))
		.length;
}

interface DeckFacts {
	bytes: ArrayBuffer;
	inputBytes: number;
	entries: number;
	uncompressedBytes: number;
	shapesPerSlide: number[];
}

async function deckFacts(): Promise<DeckFacts> {
	const bytes = await deck;
	const zip = await JSZip.loadAsync(bytes);
	const files = Object.values(zip.files);
	const handler = new PptxHandler();
	const data = await handler.load(bytes.slice(0));
	handler.dispose();
	return {
		bytes,
		inputBytes: bytes.byteLength,
		entries: files.length,
		uncompressedBytes: files
			.filter((file) => !file.dir)
			.reduce((sum, file) => sum + declaredEntrySize(file), 0),
		shapesPerSlide: data.slides.map((slide) => authored(slide.elements)),
	};
}

const facts = deckFacts();

async function loadWith(limits: PptxLoadLimits, handler = new PptxHandler()) {
	const { bytes } = await facts;
	return handler.load(bytes.slice(0), { limits });
}

async function rejection(promise: Promise<unknown>): Promise<unknown> {
	try {
		await promise;
	} catch (error) {
		return error;
	}
	throw new Error('expected the load to fail');
}

/** The runtime behind a handler, for checking what a failed load touched. */
function runtimeOf(handler: PptxHandler): Record<string, unknown> {
	return (handler as unknown as { runtime: Record<string, unknown> }).runtime;
}

describe('resolvePptxLoadLimits', () => {
	it('keeps the current defaults when nothing is set', () => {
		expect(resolvePptxLoadLimits()).toStrictEqual(DEFAULT_PPTX_LOAD_LIMITS);
		expect(resolvePptxLoadLimits({ limits: {} })).toStrictEqual(DEFAULT_PPTX_LOAD_LIMITS);
		expect(DEFAULT_PPTX_LOAD_LIMITS).toStrictEqual({
			maxInputBytes: Number.POSITIVE_INFINITY,
			maxUncompressedBytes: DEFAULT_MAX_UNCOMPRESSED_BYTES,
			maxEntries: MAX_ZIP_ENTRY_COUNT,
			maxSlides: Number.POSITIVE_INFINITY,
			maxShapes: Number.POSITIVE_INFINITY,
		});
		expect(DEFAULT_MAX_UNCOMPRESSED_BYTES).toBe(500 * 1024 * 1024);
		expect(MAX_ZIP_ENTRY_COUNT).toBe(65_536);
	});

	it('keeps the top-level maxUncompressedBytes as an alias that limits overrides', () => {
		expect(resolvePptxLoadLimits({ maxUncompressedBytes: 10 }).maxUncompressedBytes).toBe(10);
		// As before, a non-positive legacy value falls back to the default.
		expect(resolvePptxLoadLimits({ maxUncompressedBytes: 0 }).maxUncompressedBytes).toBe(
			DEFAULT_MAX_UNCOMPRESSED_BYTES,
		);
		expect(
			resolvePptxLoadLimits({ maxUncompressedBytes: 10, limits: { maxUncompressedBytes: 20 } })
				.maxUncompressedBytes,
		).toBe(20);
	});

	it.each([
		['maxSlides', 0],
		['maxShapes', -1],
		['maxInputBytes', Number.NaN],
		['maxUncompressedBytes', Number.POSITIVE_INFINITY],
		['maxEntries', MAX_ZIP_ENTRY_COUNT + 1],
		['maxSlides', '5'],
	])('rejects limits.%s = %s with a RangeError', (name, value) => {
		expect(() =>
			resolvePptxLoadLimits({ limits: { [name]: value } as unknown as PptxLoadLimits }),
		).toThrow(RangeError);
	});

	it('rejects an invalid limit before the handler does any work', async () => {
		const { bytes } = await facts;
		const handler = new PptxHandler();
		const zip = runtimeOf(handler).zip;
		const loadAsync = vi.spyOn(JSZip, 'loadAsync');
		try {
			await expect(handler.load(bytes.slice(0), { limits: { maxSlides: 0 } })).rejects.toThrow(
				'limits.maxSlides must be a positive finite number',
			);
			expect(loadAsync).not.toHaveBeenCalled();
			expect(runtimeOf(handler).zip).toBe(zip);
		} finally {
			loadAsync.mockRestore();
		}
	});
});

describe('getPptxLoadLimitViolation', () => {
	it('reads both limit error types and ignores everything else', () => {
		expect(getPptxLoadLimitViolation(new PptxLoadLimitError('x', 'maxSlides', 2, 3))).toStrictEqual(
			{ limit: 'maxSlides', maximum: 2, actual: 3 },
		);
		expect(
			getPptxLoadLimitViolation(
				new ZipBombError('x', { limit: 9, loadLimit: 'maxEntries', entryCount: 10, actual: 10 }),
			),
		).toStrictEqual({ limit: 'maxEntries', maximum: 9, actual: 10 });
		// A forged size or unsafe part name is not a limit a host can raise.
		expect(getPptxLoadLimitViolation(new ZipBombError('x', { limit: 9 }))).toBeUndefined();
		expect(getPptxLoadLimitViolation(new Error('x'))).toBeUndefined();
		expect(getPptxLoadLimitViolation('x')).toBeUndefined();
		expect(getPptxLoadLimitViolation(null)).toBeUndefined();
		// Matched by code, so a copy from another bundle is read too.
		expect(
			getPptxLoadLimitViolation({ code: 'PPTX_LOAD_LIMIT', limit: 'maxShapes', maximum: 1 }),
		).toStrictEqual({ limit: 'maxShapes', maximum: 1, actual: undefined });
	});
});

describe('PptxHandler.load limits', () => {
	it('loads the fixture deck unchanged without limits', async () => {
		const { shapesPerSlide } = await facts;
		expect(shapesPerSlide).toHaveLength(SLIDES);
		// The group and its two children count on the last slide.
		expect(shapesPerSlide).toStrictEqual([2, 2, 5]);
	});

	it('maxInputBytes: the exact size loads, one byte less fails before the archive is opened', async () => {
		const { inputBytes } = await facts;
		await expect(loadWith({ maxInputBytes: inputBytes })).resolves.toHaveProperty('slides');
		const loadAsync = vi.spyOn(JSZip, 'loadAsync');
		try {
			const error = await rejection(loadWith({ maxInputBytes: inputBytes - 1 }));
			expect(error).toBeInstanceOf(PptxLoadLimitError);
			expect(error).toMatchObject({
				code: 'PPTX_LOAD_LIMIT',
				limit: 'maxInputBytes',
				maximum: inputBytes - 1,
				actual: inputBytes,
			});
			expect((error as Error).message).toContain('maxInputBytes');
			expect(loadAsync).not.toHaveBeenCalled();
		} finally {
			loadAsync.mockRestore();
		}
	});

	it('maxEntries: the exact count loads, one less fails as a ZipBombError naming the limit', async () => {
		const { entries } = await facts;
		await expect(loadWith({ maxEntries: entries })).resolves.toHaveProperty('slides');
		const error = await rejection(loadWith({ maxEntries: entries - 1 }));
		expect(error).toBeInstanceOf(ZipBombError);
		expect(error).toMatchObject({
			code: 'ZIP_BOMB',
			loadLimit: 'maxEntries',
			limit: entries - 1,
			entryCount: entries,
			actual: entries,
		});
		expect(getPptxLoadLimitViolation(error)).toStrictEqual({
			limit: 'maxEntries',
			maximum: entries - 1,
			actual: entries,
		});
	});

	it('maxUncompressedBytes: the exact size loads, one byte less fails before inflating', async () => {
		const { uncompressedBytes } = await facts;
		await expect(loadWith({ maxUncompressedBytes: uncompressedBytes })).resolves.toHaveProperty(
			'slides',
		);
		const error = await rejection(loadWith({ maxUncompressedBytes: uncompressedBytes - 1 }));
		expect(error).toBeInstanceOf(ZipBombError);
		expect(error).toMatchObject({
			loadLimit: 'maxUncompressedBytes',
			limit: uncompressedBytes - 1,
			uncompressedBytes,
			actual: uncompressedBytes,
		});
	});

	it('keeps the top-level maxUncompressedBytes working, with limits taking precedence', async () => {
		const { bytes, uncompressedBytes } = await facts;
		await expect(
			new PptxHandler().load(bytes.slice(0), { maxUncompressedBytes: uncompressedBytes - 1 }),
		).rejects.toMatchObject({ loadLimit: 'maxUncompressedBytes' });
		await expect(
			new PptxHandler().load(bytes.slice(0), {
				maxUncompressedBytes: uncompressedBytes - 1,
				limits: { maxUncompressedBytes: uncompressedBytes },
			}),
		).resolves.toHaveProperty('slides');
	});

	it('maxSlides: the exact count loads, one less fails before themes or slides are parsed', async () => {
		await expect(loadWith({ maxSlides: SLIDES })).resolves.toHaveProperty('slides.length', SLIDES);
		const handler = new PptxHandler();
		const runtime = runtimeOf(handler);
		const loadThemeData = vi.spyOn(runtime as never, 'loadThemeData' as never);
		const parseSlide = vi.spyOn(runtime as never, 'parseSlide' as never);
		const error = await rejection(loadWith({ maxSlides: SLIDES - 1 }, handler));
		expect(error).toBeInstanceOf(PptxLoadLimitError);
		expect(error).toMatchObject({ limit: 'maxSlides', maximum: SLIDES - 1, actual: SLIDES });
		expect((error as Error).message).toBe(
			'PPTX presentation has 3 slides, more than the maxSlides limit of 2',
		);
		expect(loadThemeData).not.toHaveBeenCalled();
		expect(parseSlide).not.toHaveBeenCalled();
	});

	it('maxShapes: the exact total loads, one less fails on the slide that passes it', async () => {
		const { shapesPerSlide } = await facts;
		const total = shapesPerSlide.reduce((sum, count) => sum + count, 0);
		await expect(loadWith({ maxShapes: total })).resolves.toHaveProperty('slides.length', SLIDES);
		const error = await rejection(loadWith({ maxShapes: total - 1 }));
		expect(error).toBeInstanceOf(PptxLoadLimitError);
		expect(error).toMatchObject({ limit: 'maxShapes', maximum: total - 1, actual: total });
	});

	it('maxShapes: stops parsing the slides after the one that passes the limit', async () => {
		const { shapesPerSlide } = await facts;
		const handler = new PptxHandler();
		const parseSlide = vi.spyOn(runtimeOf(handler) as never, 'parseSlide' as never);
		// The first slide alone fits; the second passes the limit.
		const error = await rejection(loadWith({ maxShapes: shapesPerSlide[0] }, handler));
		expect(error).toMatchObject({
			limit: 'maxShapes',
			actual: shapesPerSlide[0]! + shapesPerSlide[1]!,
		});
		expect(parseSlide).toHaveBeenCalledTimes(2);
	});

	it('releases a load stopped by a slide limit and stays usable and disposable', async () => {
		const handler = new PptxHandler();
		const runtime = runtimeOf(handler);
		await rejection(loadWith({ maxShapes: 1 }, handler));
		// The partial session was released, not left behind.
		expect(runtime.zip).toBeNull();
		expect((runtime.slideMap as Map<string, unknown>).size).toBe(0);
		expect(runtime.presentationData).toBeNull();
		expect(() => handler.dispose()).not.toThrow();
		// The same handler still loads the next deck.
		await expect(loadWith({}, handler)).resolves.toHaveProperty('slides.length', SLIDES);
		handler.dispose();
	});

	it('leaves a loaded deck in place when the next input is rejected up front', async () => {
		const handler = new PptxHandler();
		await loadWith({}, handler);
		const runtime = runtimeOf(handler);
		const zip = runtime.zip;
		await rejection(loadWith({ maxInputBytes: 1 }, handler));
		await rejection(loadWith({ maxEntries: 1 }, handler));
		expect(runtime.zip).toBe(zip);
		handler.dispose();
	});

	it('applies the slide and shape limits to a pptx-viewer-json input', async () => {
		const { bytes, shapesPerSlide } = await facts;
		const handler = new PptxHandler();
		const data = await handler.load(bytes.slice(0));
		const json = new TextEncoder().encode(serializePptxToJson(data)).buffer as ArrayBuffer;
		handler.dispose();
		await expect(
			new PptxHandler().load(json.slice(0), { limits: { maxSlides: SLIDES } }),
		).resolves.toHaveProperty('slides.length', SLIDES);
		await expect(
			new PptxHandler().load(json.slice(0), { limits: { maxSlides: SLIDES - 1 } }),
		).rejects.toMatchObject({ limit: 'maxSlides', actual: SLIDES });
		await expect(
			new PptxHandler().load(json.slice(0), { limits: { maxShapes: shapesPerSlide[0] } }),
		).rejects.toMatchObject({ limit: 'maxShapes' });
		await expect(
			new PptxHandler().load(json.slice(0), { limits: { maxInputBytes: json.byteLength - 1 } }),
		).rejects.toMatchObject({ limit: 'maxInputBytes', actual: json.byteLength });
	});
});
