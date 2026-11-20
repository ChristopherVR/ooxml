import { readdirSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type { PptxElement, PptxSlide } from '../../index';
import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';

import { PptxHandler } from '../../core/PptxHandler';
import { configureCollaborationMedia } from './collaboration-media';
import { reconcileSlidesInYDoc } from './collaboration-reconcile';
import { REMOTE_VALUE_REJECTED } from './collaboration-remote-values';
import type { YDocLike, YjsFactories } from './collaboration-sync';
import { readSlidesFromYDoc, writeSlidesToYDoc } from './collaboration-sync';

const factories: YjsFactories = {
	createMap: () => new Y.Map() as unknown as ReturnType<YjsFactories['createMap']>,
	createArray: () => new Y.Array() as unknown as ReturnType<YjsFactories['createArray']>,
	createText: () => new Y.Text() as unknown as ReturnType<YjsFactories['createText']>,
};
const asDoc = (doc: Y.Doc): YDocLike => doc as unknown as YDocLike;

/**
 * What a second participant reads after `slides` was written by a first one. It allows
 * linked content, as the viewer does by default (collaboration-media.test.ts covers both).
 */
function receive(slides: PptxSlide[]): { doc: Y.Doc; slides: PptxSlide[] } {
	const sender = new Y.Doc();
	writeSlidesToYDoc(slides, asDoc(sender), factories);
	const doc = new Y.Doc();
	configureCollaborationMedia(asDoc(doc), { allowExternalContent: true });
	Y.applyUpdate(doc, Y.encodeStateAsUpdate(sender));
	return { doc, slides: readSlidesFromYDoc(asDoc(doc)) };
}

const slideWith = (...elements: Record<string, unknown>[]): PptxSlide =>
	({ id: 'ppt/slides/slide1.xml', rId: 'rId2', slideNumber: 1, elements }) as unknown as PptxSlide;

const ole = (fields: Record<string, unknown>): Record<string, unknown> => ({
	id: 'ppt/slides/slide1.xml-ole-0',
	type: 'ole',
	x: 0,
	y: 0,
	width: 100,
	height: 100,
	...fields,
});

const rejections = (slide: PptxSlide | undefined) =>
	(slide?.warnings ?? []).filter((warning) => warning.code === REMOTE_VALUE_REJECTED);

describe('remote values at the collaboration boundary', () => {
	it('drops script and markup payloads a peer puts in an embedded object and warns', () => {
		for (const payload of [
			'data:text/html,<script>localStorage.x=1</script>',
			'javascript:localStorage.x=1',
		]) {
			const { slides } = receive([
				slideWith(ole({ oleEmbeddedData: payload, oleEmbeddedMimeType: 'text/plain' })),
			]);
			const element = slides[0]?.elements[0] as Record<string, unknown> | undefined;
			if (payload.startsWith('javascript:')) {
				expect(element?.oleEmbeddedData).toBeUndefined();
			} else {
				// A markup file stays downloadable, but its claimed type cannot disagree with it.
				expect(element?.oleEmbeddedData).toBe(payload);
				expect(element?.oleEmbeddedMimeType).toBeUndefined();
			}
			expect(rejections(slides[0]).length).toBeGreaterThan(0);
			expect(rejections(slides[0])[0]?.elementId).toBe('ppt/slides/slide1.xml-ole-0');
		}
	});

	it('puts a hostile element id in safe token form', () => {
		const hostile = 'x</style><img src=x onerror=alert(1)>';
		const { slides } = receive([slideWith(ole({ id: hostile }))]);
		const id = slides[0]?.elements[0]?.id ?? '';
		expect(id).toMatch(/^[A-Za-z0-9\-._~/%]+$/u);
		expect(id).not.toBe(hostile);
	});

	it('checks group children and notes shapes too', () => {
		const group = {
			id: 'g',
			type: 'group',
			x: 0,
			y: 0,
			width: 1,
			height: 1,
			children: [
				ole({ id: 'child "1"', oleEmbeddedData: 'javascript:alert(1)' }),
				{ type: 'shape', x: 0, y: 0, width: 1, height: 1 },
			],
		};
		const slide = {
			...slideWith(group),
			notesShapes: [{ id: 'n1', type: 'image', imageData: 'vbscript:x' }],
		} as unknown as PptxSlide;
		const { slides } = receive([slide]);
		const [first] = slides;
		if (!first) throw new Error('no slide');
		const children = (first.elements[0] as { children: PptxElement[] }).children;
		expect(children).toHaveLength(1);
		expect(children[0]?.id).toBe('child%20%221%22');
		expect((children[0] as Record<string, unknown>).oleEmbeddedData).toBeUndefined();
		expect(
			(first.notesShapes?.[0] as Record<string, unknown> | undefined)?.imageData,
		).toBeUndefined();
		expect(rejections(slides[0]).map((warning) => warning.elementId)).toEqual([
			'child%20%221%22',
			undefined,
			'n1',
		]);
	});

	it('drops resources of the wrong kind and malformed content types', () => {
		const { slides } = receive([
			{
				...slideWith(
					{ id: 'p', type: 'image', imageData: 'data:text/html,<b>x</b>' },
					{
						id: 'm',
						type: 'media',
						mediaData: 'file:///etc/passwd',
						mediaMimeType: 'video/mp4"><b>',
					},
				),
				backgroundImage: 'javascript:alert(1)',
			} as unknown as PptxSlide,
		]);
		const [picture, media] = (slides[0]?.elements ?? []) as Record<string, unknown>[];
		expect(picture?.imageData).toBeUndefined();
		expect(media?.mediaData).toBeUndefined();
		expect(media?.mediaMimeType).toBeUndefined();
		expect(slides[0]?.backgroundImage).toBeUndefined();
		expect(rejections(slides[0])).toHaveLength(4);
	});

	it('keeps the values a real deck holds', () => {
		const png = 'data:image/png;base64,iVBORw0KGgo=';
		const { slides } = receive([
			slideWith(
				ole({
					oleEmbeddedData: 'data:application/pdf;base64,JVBERi0=',
					oleEmbeddedMimeType: 'application/pdf',
					previewImageData: png,
				}),
				{ id: 'p', type: 'image', imageData: 'https://example.com/a.png', svgData: png },
				{
					id: 'm',
					type: 'media',
					mediaData: 'data:video/mp4;base64,AAAA',
					mediaMimeType: 'video/mp4',
				},
			),
		]);
		expect(rejections(slides[0])).toHaveLength(0);
		const [embedded, picture, media] = (slides[0]?.elements ?? []) as Record<string, unknown>[];
		expect(embedded?.oleEmbeddedMimeType).toBe('application/pdf');
		expect(picture?.imageData).toBe('https://example.com/a.png');
		expect(media?.mediaData).toBe('data:video/mp4;base64,AAAA');
	});

	it('never writes its own rejection warnings back into the shared document', () => {
		const { doc, slides } = receive([slideWith(ole({ oleEmbeddedData: 'javascript:x' }))]);
		expect(rejections(slides[0]).length).toBeGreaterThan(0);
		reconcileSlidesInYDoc(slides, asDoc(doc), factories);
		const peer = new Y.Doc();
		Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
		const shared = (peer.getArray('pptx:slides').get(0) as Y.Map<unknown>).get('_wa');
		expect(shared).toBeUndefined();
		// The rejected value is gone from the room once this client writes its view back.
		expect(rejections(readSlidesFromYDoc(asDoc(peer))[0])).toHaveLength(0);
	});
});

describe('remote values from real decks', () => {
	const fixtures = resolve(__dirname, '../../__tests__/fixtures');
	const decks = [
		...readdirSync(fixtures).map((name) => resolve(fixtures, name)),
		...readdirSync(resolve(fixtures, 'e2e')).map((name) => resolve(fixtures, 'e2e', name)),
	].filter((path) => path.endsWith('.pptx'));

	it('passes every value of every fixture deck unchanged', { timeout: 240_000 }, async () => {
		const failures: string[] = [];
		let elements = 0;
		for (const path of decks) {
			let loaded: PptxSlide[];
			try {
				const bytes = readFileSync(path);
				const data = await new PptxHandler().load(
					bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
				);
				loaded = data.slides;
			} catch {
				continue; // encrypted or deliberately broken fixtures
			}
			const { slides } = receive(loaded);
			for (const slide of slides) {
				elements += slide.elements.length;
				for (const warning of rejections(slide)) {
					failures.push(`${path}: ${warning.elementId ?? slide.id} ${warning.message}`);
				}
			}
			const sent = loaded.flatMap((slide) => slide.elements.map((element) => element.id));
			const read = slides.flatMap((slide) => slide.elements.map((element) => element.id));
			expect(read, path).toEqual(sent);
		}
		expect(failures).toEqual([]);
		expect(elements).toBeGreaterThan(500);
	});
});
