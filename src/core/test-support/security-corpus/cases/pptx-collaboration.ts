import * as Y from 'yjs';

import type { PptxSlide } from '../../../pptx/core/types';
import {
	readSlidesFromYDoc,
	writeSlidesToYDoc,
	type YDocLike,
	type YjsFactories,
} from '../../../pptx/editor/render/collaboration-sync';
import type { SecurityCase } from '../case';

const SAFE_ID = /^[A-Za-z0-9\-._~/%]+$/u;
const factories = {
	createMap: () => new Y.Map(),
	createArray: () => new Y.Array(),
	createText: () => new Y.Text(),
} as unknown as YjsFactories;

/** What a second participant reads after a first one shares `elements` on one slide. */
function received(elements: Record<string, unknown>[]): Record<string, unknown>[] {
	const sender = new Y.Doc();
	const slide = { id: 'ppt/slides/slide1.xml', rId: 'rId2', slideNumber: 1, elements };
	writeSlidesToYDoc([slide as unknown as PptxSlide], sender as unknown as YDocLike, factories);
	const receiver = new Y.Doc();
	Y.applyUpdate(receiver, Y.encodeStateAsUpdate(sender));
	const [first] = readSlidesFromYDoc(receiver as unknown as YDocLike);
	return (first?.elements ?? []) as unknown as Record<string, unknown>[];
}

const box = { x: 0, y: 0, width: 10, height: 10 };

export const cases: readonly SecurityCase[] = [
	{
		name: 'script URLs, mismatched types and markup ids a peer shares arrive inert',
		limits: { ms: 2_000, heapMb: 160 },
		run() {
			const [scriptUrl, markupFile, picture] = received([
				{ ...box, id: 'ole-1', type: 'ole', oleEmbeddedData: 'javascript:void 0' },
				{
					...box,
					id: 'ole-2"></style><i>',
					type: 'ole',
					oleEmbeddedData: 'data:text/html,<p>x</p>',
					oleEmbeddedMimeType: 'text/plain',
				},
				{ ...box, id: 'pic-1', type: 'image', imageData: 'data:text/html,<p>x</p>' },
			]);
			return {
				scriptUrl: scriptUrl?.oleEmbeddedData ?? null,
				claimedType: markupFile?.oleEmbeddedMimeType ?? null,
				safeId: SAFE_ID.test(String(markupFile?.id)),
				markupPicture: picture?.imageData ?? null,
			};
		},
		expected: { scriptUrl: null, claimedType: null, safeId: true, markupPicture: null },
	},
];
