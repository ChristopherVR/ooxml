import type { PptxElement, PptxSlide } from '../../index';
import { afterEach, describe, expect, it, vi } from 'vitest';
import * as Y from 'yjs';

import {
	configureCollaborationMedia,
	MAX_SHARED_MEDIA_BYTES,
	YDOC_MEDIA_KEY,
} from './collaboration-media';
import { reconcileSlidesInYDoc } from './collaboration-reconcile';
import { REMOTE_MEDIA_WITHHELD, REMOTE_VALUE_REJECTED } from './collaboration-remote-values';
import type { YDocLike, YjsFactories } from './collaboration-sync';
import { observeYDocSlides, readSlidesFromYDoc } from './collaboration-sync';

const factories: YjsFactories = {
	createMap: () => new Y.Map() as unknown as ReturnType<YjsFactories['createMap']>,
	createArray: () => new Y.Array() as unknown as ReturnType<YjsFactories['createArray']>,
	createText: () => new Y.Text() as unknown as ReturnType<YjsFactories['createText']>,
};
const asDoc = (doc: Y.Doc): YDocLike => doc as unknown as YDocLike;

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]);
const MP4 = new Uint8Array([0, 0, 0, 0x20, 0x66, 0x74, 0x79, 0x70, 9, 9]);

const urls: string[] = [];
function objectUrl(bytes: Uint8Array, type: string): string {
	const url = URL.createObjectURL(new Blob([bytes.slice()], { type }));
	urls.push(url);
	return url;
}
afterEach(() => {
	for (const url of urls.splice(0)) {
		URL.revokeObjectURL(url);
	}
});

/** Two participants of one room: every update of either reaches the other at once. */
function room(): { sharer: Y.Doc; joiner: Y.Doc } {
	const sharer = new Y.Doc();
	const joiner = new Y.Doc();
	sharer.on('update', (update: Uint8Array, origin: unknown) => {
		if (origin !== 'remote') Y.applyUpdate(joiner, update, 'remote');
	});
	joiner.on('update', (update: Uint8Array, origin: unknown) => {
		if (origin !== 'remote') Y.applyUpdate(sharer, update, 'remote');
	});
	return { sharer, joiner };
}

const slideOf = (...elements: unknown[]): PptxSlide =>
	({ id: 's1', slideNumber: 1, elements }) as unknown as PptxSlide;
const picture = (imageData: string, extra: Record<string, unknown> = {}): PptxElement =>
	({
		type: 'picture',
		id: 'pic',
		x: 0,
		y: 0,
		width: 10,
		height: 10,
		imagePath: 'ppt/media/image1.png',
		imageData,
		...extra,
	}) as unknown as PptxElement;
const mediaEntries = (doc: Y.Doc): Y.Map<unknown> => doc.getMap(YDOC_MEDIA_KEY);

async function bytesOf(url: string | undefined): Promise<Uint8Array> {
	expect(url).toBeTypeOf('string');
	return new Uint8Array(await (await fetch(url!)).arrayBuffer());
}

function firstElement(doc: Y.Doc): Record<string, unknown> {
	return readSlidesFromYDoc(asDoc(doc))[0]!.elements[0] as unknown as Record<string, unknown>;
}

describe('collaboration media: the room carries the bytes, never page-local URLs', () => {
	it('gives a joiner its own URL for a loaded picture', async () => {
		const { sharer, joiner } = room();
		const source = objectUrl(PNG, 'image/png');
		reconcileSlidesInYDoc([slideOf(picture(source))], asDoc(sharer), factories);
		await vi.waitFor(() => expect(mediaEntries(joiner).size).toBe(1));

		expect(JSON.stringify(sharer.toJSON())).not.toContain('blob:');
		const seen = firstElement(joiner).imageData as string;
		expect(seen).not.toBe(source);
		expect(await bytesOf(seen)).toEqual(PNG);
		expect(readSlidesFromYDoc(asDoc(joiner))[0]!.warnings).toBeUndefined();
		// The sharer keeps reading its own URL back.
		expect(firstElement(sharer).imageData).toBe(source);
	});

	it('shares an embedded video that names only its part, through the URL this page plays', async () => {
		const { sharer, joiner } = room();
		const source = objectUrl(MP4, 'video/mp4');
		configureCollaborationMedia(asDoc(sharer), {
			localMediaUrl: (path) => (path === 'ppt/media/media1.mp4' ? source : undefined),
		});
		const video = {
			type: 'media',
			id: 'vid',
			x: 0,
			y: 0,
			width: 10,
			height: 10,
			mediaType: 'video',
			mediaPath: 'ppt/media/media1.mp4',
		};
		reconcileSlidesInYDoc([slideOf(video)], asDoc(sharer), factories);
		await vi.waitFor(() => expect(mediaEntries(joiner).size).toBe(1));

		const seen = firstElement(joiner);
		expect(seen.mediaPath).toBe('ppt/media/media1.mp4');
		expect(await bytesOf(seen.mediaData as string)).toEqual(MP4);
	});

	it('covers pictures inside fills and text runs', async () => {
		const { sharer, joiner } = room();
		const fill = objectUrl(PNG, 'image/png');
		const shape = {
			type: 'shape',
			id: 'shp',
			x: 0,
			y: 0,
			width: 10,
			height: 10,
			shapeStyle: { fillMode: 'image', fillImageUrl: fill },
			textSegments: [{ text: 'Run', style: { textFillBlipUrl: fill } }],
		};
		reconcileSlidesInYDoc([slideOf(shape)], asDoc(sharer), factories);
		await vi.waitFor(() => expect(mediaEntries(joiner).size).toBe(1));

		expect(JSON.stringify(sharer.toJSON())).not.toContain('blob:');
		const seen = firstElement(joiner) as {
			shapeStyle: { fillImageUrl: string };
			textSegments: { style: { textFillBlipUrl: string } }[];
		};
		expect(await bytesOf(seen.shapeStyle.fillImageUrl)).toEqual(PNG);
		expect(seen.textSegments[0]!.style.textFillBlipUrl).toBe(seen.shapeStyle.fillImageUrl);
	});

	it('writes nothing new when a joiner publishes the slides it read', async () => {
		const { sharer, joiner } = room();
		reconcileSlidesInYDoc(
			[slideOf(picture(objectUrl(PNG, 'image/png')))],
			asDoc(sharer),
			factories,
		);
		await vi.waitFor(() => expect(mediaEntries(joiner).size).toBe(1));
		const before = Y.encodeStateVector(sharer);

		reconcileSlidesInYDoc(readSlidesFromYDoc(asDoc(joiner)), asDoc(joiner), factories);
		await new Promise((resolve) => setTimeout(resolve, 10));
		expect(Y.encodeStateVector(sharer)).toEqual(before);
	});

	it('re-reads the slides when the bytes arrive after them', async () => {
		const { sharer, joiner } = room();
		const changes: unknown[] = [];
		const stop = observeYDocSlides(asDoc(joiner), (_events, transaction) =>
			changes.push(transaction?.origin),
		);
		reconcileSlidesInYDoc(
			[slideOf(picture(objectUrl(PNG, 'image/png')))],
			asDoc(sharer),
			factories,
		);
		// The slide arrives first, with only the token of its picture.
		expect(firstElement(joiner).imageData).toMatch(/^ooxml-media:/u);
		await vi.waitFor(() => expect(changes.length).toBeGreaterThanOrEqual(2));
		expect(await bytesOf(firstElement(joiner).imageData as string)).toEqual(PNG);
		stop();
	});

	it('keeps the bytes of a picture over the room budget out and tells the joiner why', async () => {
		const { sharer, joiner } = room();
		const large = new Uint8Array(MAX_SHARED_MEDIA_BYTES + 1);
		reconcileSlidesInYDoc(
			[slideOf(picture(objectUrl(large, 'image/png')))],
			asDoc(sharer),
			factories,
		);
		await vi.waitFor(() => expect(mediaEntries(joiner).size).toBe(1));

		const entry = [...mediaEntries(sharer).values()][0] as Record<string, unknown>;
		expect(entry.b).toBeUndefined();
		expect(entry.n).toBe(large.length);
		const [slide] = readSlidesFromYDoc(asDoc(joiner));
		// The token stays (it draws nothing), so the joiner's next write does not delete it.
		expect((slide!.elements[0] as { imageData?: string }).imageData).toMatch(/^ooxml-media:/u);
		expect(slide!.warnings?.map((warning) => warning.code)).toEqual([REMOTE_MEDIA_WITHHELD]);
	});

	it('keeps to the smaller budget a peer-to-peer transport sets', async () => {
		const { sharer, joiner } = room();
		configureCollaborationMedia(asDoc(sharer), { maxBytes: PNG.length - 1 });
		// A later call that does not name the budget keeps it.
		configureCollaborationMedia(asDoc(sharer), { localMediaUrl: () => undefined });
		reconcileSlidesInYDoc(
			[slideOf(picture(objectUrl(PNG, 'image/png')))],
			asDoc(sharer),
			factories,
		);
		await vi.waitFor(() => expect(mediaEntries(joiner).size).toBe(1));
		const entry = [...mediaEntries(joiner).values()][0] as Record<string, unknown>;
		expect(entry).toEqual({ t: 'image/png', n: PNG.length });
	});

	it('removes the bytes no slide uses before adding new ones', async () => {
		const { sharer, joiner } = room();
		reconcileSlidesInYDoc(
			[slideOf(picture(objectUrl(PNG, 'image/png')))],
			asDoc(sharer),
			factories,
		);
		await vi.waitFor(() => expect(mediaEntries(joiner).size).toBe(1));
		const [first] = [...mediaEntries(sharer).keys()];

		const replacement = picture(objectUrl(MP4, 'image/png'), { id: 'other' });
		reconcileSlidesInYDoc([slideOf(replacement)], asDoc(sharer), factories);
		await vi.waitFor(() => expect(mediaEntries(joiner).has(first!)).toBe(false));
		expect(mediaEntries(joiner).size).toBe(1);
	});

	it('gives an SVG picture a data URL, never a URL of the page origin', () => {
		const { joiner } = room();
		const id = 'a'.repeat(32);
		joiner.transact(() => {
			mediaEntries(joiner).set(id, { t: 'image/svg+xml', b: new TextEncoder().encode('<svg/>') });
		});
		reconcileSlidesInYDoc([slideOf(picture(`ooxml-media:${id}`))], asDoc(joiner), factories);
		expect(firstElement(joiner).imageData).toMatch(/^data:image\/svg\+xml;base64,/u);
	});

	it('leaves out a page-local URL or a malformed entry a peer wrote, with a warning', () => {
		const { joiner } = room();
		const id = 'b'.repeat(32);
		joiner.transact(() => mediaEntries(joiner).set(id, { t: 'not a type', b: PNG }));
		const slides = joiner.getArray('pptx:slides');
		joiner.transact(() => {
			const slide = new Y.Map<unknown>();
			const elements = new Y.Array<Y.Map<unknown>>();
			for (const [elementId, imageData] of [
				['raw', 'blob:http://elsewhere.test/1'],
				['bad', `ooxml-media:${id}`],
			]) {
				const element = new Y.Map<unknown>();
				element.set('id', elementId);
				element.set('type', 'picture');
				element.set('imageData', imageData);
				elements.push([element]);
			}
			slide.set('id', 's1');
			slide.set('elements', elements);
			slides.push([slide]);
		});
		const [slide] = readSlidesFromYDoc(asDoc(joiner));
		expect(slide!.elements.map((element) => (element as { imageData?: string }).imageData)).toEqual(
			[undefined, undefined],
		);
		expect(slide!.warnings?.map((warning) => [warning.code, warning.elementId])).toEqual([
			[REMOTE_VALUE_REJECTED, 'raw'],
			[REMOTE_VALUE_REJECTED, 'bad'],
		]);
	});
});

describe('collaboration media: what a joiner keeps while bytes are missing', () => {
	it('keeps a linked picture as its link, for the viewer to gate where it draws', () => {
		const { sharer, joiner } = room();
		const linked = picture('https://linked.test/picture.png', {
			shapeStyle: { fillImageUrl: 'https://linked.test/fill.png' },
		});
		reconcileSlidesInYDoc([slideOf(linked)], asDoc(sharer), factories);
		const seen = firstElement(joiner) as Record<string, Record<string, unknown> | string>;
		expect(seen.imageData).toBe('https://linked.test/picture.png');
		expect((seen.shapeStyle as Record<string, unknown>).fillImageUrl).toBe(
			'https://linked.test/fill.png',
		);
		expect(mediaEntries(sharer).size).toBe(0);
	});

	it('keeps a token whose bytes have not arrived, so writing the slides back keeps it', () => {
		const { joiner } = room();
		const token = `ooxml-media:${'c'.repeat(32)}`;
		reconcileSlidesInYDoc([slideOf(picture(token))], asDoc(joiner), factories);
		const [slide] = readSlidesFromYDoc(asDoc(joiner));
		expect((slide!.elements[0] as { imageData?: string }).imageData).toBe(token);
		expect(slide!.warnings).toBeUndefined();
		reconcileSlidesInYDoc(readSlidesFromYDoc(asDoc(joiner)), asDoc(joiner), factories);
		expect(firstElement(joiner).imageData).toBe(token);
	});
});
