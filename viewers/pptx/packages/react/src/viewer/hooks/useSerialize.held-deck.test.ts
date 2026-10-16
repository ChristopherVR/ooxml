// @vitest-environment node
/**
 * A Design gallery hover preview swaps previewed slides into the viewer state. Save and the
 * autosave recovery snapshot must keep serializing the real deck while it shows.
 */
import type { PptxHandler, PptxSlide } from 'pptx-viewer-core';
import type React from 'react';
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';

import { useSerialize } from './useSerialize';
import type { UseSerializeInput } from './useSerialize';

const slide = (color: string): PptxSlide =>
	({ id: 's1', elements: [], backgroundColor: color }) as unknown as PptxSlide;

function serializerFor(
	handler: PptxHandler,
	overrides: Partial<UseSerializeInput>,
): () => Promise<Uint8Array | null> {
	const input: UseSerializeInput = {
		slides: [slide('#000000')],
		templateElementsBySlideId: {},
		activeSlideIndex: 0,
		canvasSize: { width: 960, height: 540 },
		guides: [],
		headerFooter: {},
		presentationProperties: {},
		customShows: [],
		sections: [],
		coreProperties: undefined,
		appProperties: undefined,
		customProperties: [],
		tagCollections: [],
		slideMasters: [],
		notesMaster: undefined,
		handoutMaster: undefined,
		handlerRef: { current: handler } as React.RefObject<PptxHandler | null>,
		inlineEditingElementIdRef: { current: null },
		inlineEditingTextRef: { current: '' },
		...overrides,
	};
	let serialize: (() => Promise<Uint8Array | null>) | null = null;
	function Harness(): null {
		serialize = useSerialize(input);
		return null;
	}
	renderToStaticMarkup(createElement(Harness));
	return serialize!;
}

describe('serializing while a gallery preview shows', () => {
	it('writes the real slides, not the previewed ones', async () => {
		const saved: PptxSlide[][] = [];
		const handler = {
			save: (slides: PptxSlide[]) => {
				saved.push(slides);
				return Promise.resolve(new Uint8Array([1]));
			},
		} as unknown as PptxHandler;
		const heldDeckRef = {
			current: { slides: [slide('#FFFFFF')], templateElementsBySlideId: {} },
		};
		await serializerFor(handler, { heldDeckRef, purpose: 'recovery-snapshot' })();
		expect(saved[0]![0]!.backgroundColor).toBe('#FFFFFF');
		heldDeckRef.current = undefined as never;
		await serializerFor(handler, { heldDeckRef })();
		expect(saved[1]![0]!.backgroundColor).toBe('#000000');
	});
});
