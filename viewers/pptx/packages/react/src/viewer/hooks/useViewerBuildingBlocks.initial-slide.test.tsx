// @vitest-environment happy-dom
/**
 * `useViewerBuildingBlocks({ initialSlide })` opens a loaded deck on that slide, as
 * `PowerPointViewer`'s `initialSlide` does. Hosts used to call `goTo()` in an effect after load
 * and hide the canvas until it landed, because every render with the deck showed slide 1 first.
 */
import { PptxHandler } from 'pptx-viewer-core';
import React, { act, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import type { PowerPointViewerHandle } from '../types';
import { useViewerBuildingBlocks } from './useViewerBuildingBlocks';
import type { ViewerBuildingBlocksResult } from './useViewerBuildingBlocks';

(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT =
	true;

let deck: Uint8Array;
let root: Root;
let container: HTMLDivElement;
let latest: ViewerBuildingBlocksResult;
/** The slide index and id every render with a loaded deck put on the canvas, in order. */
let rendered: [number | undefined, string][];
const onActiveSlideChange = vi.fn();

beforeAll(async () => {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 3 });
	deck = await handler.save(data.slides);
	handler.dispose();
});

function Harness(props: {
	content: Uint8Array;
	initialSlide?: number;
	handle?: React.RefObject<PowerPointViewerHandle | null>;
}) {
	latest = useViewerBuildingBlocks({
		content: props.content,
		initialSlide: props.initialSlide,
		handle: props.handle,
		onActiveSlideChange,
	});
	const slide = latest.canvasProps.activeSlide;
	if (!latest.loading && slide) {
		rendered.push([latest.canvasProps.activeSlideIndex, slide.id]);
	}
	return <div />;
}

async function mount(
	initialSlide?: number,
	handle?: React.RefObject<PowerPointViewerHandle | null>,
) {
	rendered = [];
	container = document.createElement('div');
	document.body.append(container);
	root = createRoot(container);
	await act(async () =>
		root.render(<Harness content={deck} initialSlide={initialSlide} handle={handle} />),
	);
	const deadline = Date.now() + 10000;
	while ((latest.loading || !latest.canvasProps.activeSlide) && Date.now() < deadline) {
		await act(async () => {
			await new Promise<void>((resolve) => {
				setTimeout(resolve, 0);
			});
		});
	}
	expect(latest.error).toBeNull();
}

describe('useViewerBuildingBlocks initialSlide', () => {
	afterEach(async () => {
		await act(async () => root?.unmount());
		container?.remove();
		onActiveSlideChange.mockClear();
	});

	it('renders the loaded deck on the initial slide first, never slide 1', async () => {
		await mount(2);
		expect(latest.canvasProps.activeSlideIndex).toBe(2);
		expect(rendered.length).toBeGreaterThan(0);
		// Every render that had the deck showed the third slide: no frame of slide 1.
		expect(new Set(rendered.map(([index]) => index))).toStrictEqual(new Set([2]));
		expect(new Set(rendered.map(([, id]) => id)).size).toBe(1);
		// The index is reported once on mount, before any deck exists (as PowerPointViewer does),
		// then as the loaded deck's slide.
		expect(onActiveSlideChange).toHaveBeenLastCalledWith(2);
	});

	it('clamps an out-of-range initial slide into the deck', async () => {
		await mount(99);
		expect(latest.canvasProps.activeSlideIndex).toBe(2);
		await act(async () => root.unmount());
		await mount(-4);
		expect(latest.canvasProps.activeSlideIndex).toBe(0);
	});

	it('keeps slide 1 as the default', async () => {
		await mount();
		expect(latest.canvasProps.activeSlideIndex).toBe(0);
	});

	it('ignores a later change for the open deck; the handle still navigates', async () => {
		const handle = createRef<PowerPointViewerHandle>();
		await mount(1, handle);
		await act(async () => root.render(<Harness content={deck} initialSlide={2} handle={handle} />));
		expect(latest.canvasProps.activeSlideIndex).toBe(1);
		await act(async () => handle.current?.goTo(2));
		expect(latest.canvasProps.activeSlideIndex).toBe(2);
	});
});
