// @vitest-environment happy-dom
/**
 * A Design gallery hover preview swaps previewed slides into the viewer state. Save and the
 * autosave recovery snapshot must keep serializing the real deck while it shows.
 */
import type { PptxHandler, PptxSlide } from 'pptx-viewer-core';
import type React from 'react';
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it, vi } from 'vitest';

import { useSerialize } from './useSerialize';
import type { UseSerializeInput } from './useSerialize';

const slide = (color: string): PptxSlide =>
	({ id: 's1', elements: [], backgroundColor: color }) as unknown as PptxSlide;

function inputFor(handler: PptxHandler, overrides: Partial<UseSerializeInput>): UseSerializeInput {
	return {
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
}

function serializerFor(
	handler: PptxHandler,
	overrides: Partial<UseSerializeInput>,
): () => Promise<Uint8Array | null> {
	const input = inputFor(handler, overrides);
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
		const serialize = serializerFor(handler, { heldDeckRef, purpose: 'recovery-snapshot' });
		await serialize();
		expect(saved[0]![0]!.backgroundColor).toBe('#FFFFFF');
		heldDeckRef.current = undefined as never;
		await serialize();
		expect(saved[1]![0]!.backgroundColor).toBe('#000000');
	});

	it('uses a replacement held-deck ref after a render with otherwise stable inputs', async () => {
		const saved: PptxSlide[][] = [];
		const handler = {
			save: (slides: PptxSlide[]) => {
				saved.push(slides);
				return Promise.resolve(new Uint8Array([1]));
			},
		} as unknown as PptxHandler;
		const input = inputFor(handler, {
			tableStylesToDelete: [],
			heldDeckRef: { current: { slides: [slide('#FFFFFF')], templateElementsBySlideId: {} } },
		});
		let serialize: ReturnType<typeof useSerialize> | undefined;
		function Harness({ value }: { value: UseSerializeInput }): null {
			serialize = useSerialize(value);
			return null;
		}
		const container = document.body.appendChild(document.createElement('div'));
		const root = createRoot(container);
		vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
		try {
			act(() => root.render(createElement(Harness, { value: input })));
			await serialize!();
			const heldDeckRef = {
				current: { slides: [slide('#FF0000')], templateElementsBySlideId: {} },
			};
			act(() => root.render(createElement(Harness, { value: { ...input, heldDeckRef } })));
			await serialize!();
			expect(saved.map((slides) => slides[0]!.backgroundColor)).toEqual(['#FFFFFF', '#FF0000']);
		} finally {
			act(() => root.unmount());
			container.remove();
			vi.unstubAllGlobals();
		}
	});
});
