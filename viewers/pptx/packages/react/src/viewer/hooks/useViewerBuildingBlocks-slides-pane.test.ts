// @vitest-environment happy-dom
/**
 * `useViewerBuildingBlocks`'s `slidesPaneProps` and the exported `SlidesPane`:
 * the slide thumbnail rail a host composes beside its own `SlideCanvas`.
 * Loads a real deck through the hook and mounts the rail, with the same
 * manual `createRoot` + `act` harness as `useViewerBuildingBlocks.test.ts`.
 */
import { readFileSync } from 'node:fs';
import { resolve as resolvePath } from 'node:path';

import { PptxHandler } from 'pptx-viewer-core';
import React, { act, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';

import { SlidesPane } from '../../index';
import type { SlidesPaneLayoutGalleryRequest } from '../../index';
import type { PowerPointViewerHandle } from '../types';
import { useViewerBuildingBlocks } from './useViewerBuildingBlocks';
import type { ViewerBuildingBlocksResult } from './useViewerBuildingBlocks';

// Autosave's timer and storage are not under test here.
// oxlint-disable-next-line prefer-ending-with-an-expect
vi.mock(import('./useAutosave'), () => ({
	useAutosave: () => ({
		autosaveStatus: { state: 'idle' as const },
		triggerAutosave: async () => {},
	}),
}));

let threeSlides: Uint8Array;

beforeAll(async () => {
	const created = await PptxHandler.create({ title: 'Rail Fixture', initialSlideCount: 3 });
	threeSlides = await created.handler.save(created.data.slides);
	created.handler.dispose();
});

let container: HTMLDivElement;
let root: Root;
let latest: ViewerBuildingBlocksResult | null = null;

function Harness({
	content,
	canEdit = true,
	handle,
	onDirtyChange,
	onOpenLayoutGallery,
}: {
	content?: Uint8Array;
	canEdit?: boolean;
	handle?: React.RefObject<PowerPointViewerHandle | null>;
	onDirtyChange?: (dirty: boolean) => void;
	onOpenLayoutGallery?: (request: SlidesPaneLayoutGalleryRequest) => void;
}): React.ReactElement {
	const result = useViewerBuildingBlocks({
		content: content ?? threeSlides,
		canEdit,
		handle,
		onDirtyChange,
		onOpenLayoutGallery,
	});
	latest = result;
	return React.createElement(SlidesPane, result.slidesPaneProps);
}

async function flush(): Promise<void> {
	await act(async () => {
		await new Promise<void>((resolve) => {
			setTimeout(resolve, 0);
		});
	});
}

async function flushUntil(isDone: () => boolean, timeoutMs = 10_000): Promise<void> {
	const deadline = Date.now() + timeoutMs;
	while (!isDone() && Date.now() < deadline) {
		await flush();
	}
}

async function mount(
	props: Parameters<typeof Harness>[0] = {},
): Promise<ViewerBuildingBlocksResult> {
	await act(async () => {
		root.render(React.createElement(Harness, props));
	});
	await flushUntil(
		() => latest?.loading === false && (latest?.slidesPaneProps.slides.length ?? 0) > 0,
	);
	return latest as ViewerBuildingBlocksResult;
}

function rows(): HTMLElement[] {
	return [...container.querySelectorAll<HTMLElement>('[data-pptx-chrome="slide-row"]')];
}

beforeEach(() => {
	latest = null;
	vi.stubGlobal('innerWidth', 1280);
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	vi.unstubAllGlobals();
});

describe('useViewerBuildingBlocks slidesPaneProps', () => {
	it('lists every slide of the loaded deck, with the active slide and canvas size', async () => {
		const { slidesPaneProps, canvasProps } = await mount();

		expect(slidesPaneProps.slides).toHaveLength(3);
		expect(slidesPaneProps.activeSlideIndex).toBe(0);
		expect(slidesPaneProps.canvasSize).toStrictEqual(canvasProps.canvasSize);
		expect(slidesPaneProps.canEdit).toBeTruthy();
		expect(slidesPaneProps.isOpen).toBeTruthy();
		// The rail's thumbnail of the active slide draws the same layout and
		// master elements the canvas does.
		const activeId = slidesPaneProps.slides[0].id;
		expect(slidesPaneProps.templateElementsBySlideId[activeId] ?? []).toStrictEqual(
			canvasProps.templateElements,
		);
	}, 15_000);

	it('numbers thumbnails from the deck start number and shares the canvas contexts', async () => {
		// `p:presentation/@firstSlideNum="45"`: slide i is numbered 44 + i.
		// Without IntersectionObserver the lazy thumbnails render at once.
		vi.stubGlobal('IntersectionObserver', undefined);
		const content = new Uint8Array(
			readFileSync(
				resolvePath(process.cwd(), '../../../../e2e/pptx/fixtures/first-slide-number.pptx'),
			),
		);
		const { slidesPaneProps, canvasProps } = await mount({ content });

		expect(slidesPaneProps.fieldContext?.firstSlideNumber).toBe(45);
		// Deck-wide fields match the canvas; the slide number and title are
		// per-slide, and each thumbnail re-points them at its own slide.
		const {
			slideNumber: _number,
			slideTitle: _title,
			...canvasDeckFields
		} = canvasProps.fieldContext ?? {};
		// The rail leaves the per-slide keys present but undefined.
		const defined = (context: object | undefined) =>
			Object.fromEntries(Object.entries(context ?? {}).filter(([, value]) => value !== undefined));
		expect(defined(slidesPaneProps.fieldContext)).toStrictEqual(defined(canvasDeckFields));
		expect(slidesPaneProps.fieldContext?.slideNumber).toBeUndefined();
		expect(slidesPaneProps.tableStyleContext).toStrictEqual(canvasProps.tableStyleContext);

		const frames = rows().map((row) => row.querySelector('[data-pptx-chrome="slide-frame"]'));
		expect(frames).toHaveLength(3);
		for (const [at, frame] of frames.entries()) {
			expect(frame?.textContent).toContain(`Slide ${45 + at}`);
		}
		expect(frames[0]?.textContent).not.toContain('Slide 1');
	}, 15_000);

	it('renders a thumbnail rail row per slide in the mounted SlidesPane', async () => {
		await mount();

		expect(rows()).toHaveLength(3);
		for (const row of rows()) {
			expect(row.querySelector('[data-pptx-chrome="slide-frame"]')).not.toBeNull();
		}
		expect(rows()[0].getAttribute('aria-current')).toBe('true');
	}, 15_000);

	it('moves the canvas to a clicked slide without marking the deck dirty', async () => {
		const handle = createRef<PowerPointViewerHandle>();
		const dirtyChanges: boolean[] = [];
		await mount({ handle, onDirtyChange: (dirty) => dirtyChanges.push(dirty) });

		await act(async () => {
			rows()[2].click();
		});
		await flushUntil(() => latest?.slidesPaneProps.activeSlideIndex === 2);

		expect(handle.current?.getActiveSlideIndex()).toBe(2);
		expect(latest?.canvasProps.activeSlide?.id).toBe(latest?.slidesPaneProps.slides[2].id);
		expect(handle.current?.isDirty()).toBeFalsy();
		expect(dirtyChanges).not.toContain(true);
	}, 15_000);

	it('shares its open state with the toolbar sidebar toggle', async () => {
		const result = await mount();
		expect(result.toolbarProps.isSidebarCollapsed).toBeFalsy();

		await act(async () => {
			latest?.slidesPaneProps.onCollapse();
		});
		expect(latest?.slidesPaneProps.isOpen).toBeFalsy();
		expect(latest?.toolbarProps.isSidebarCollapsed).toBeTruthy();
		expect(rows()).toHaveLength(0);

		await act(async () => {
			latest?.toolbarProps.onToggleSidebar?.();
		});
		expect(latest?.slidesPaneProps.isOpen).toBeTruthy();
		expect(rows()).toHaveLength(3);
	}, 15_000);

	it('makes the slide active and hands the Layout gallery request to the host', async () => {
		const requests: SlidesPaneLayoutGalleryRequest[] = [];
		await mount({ onOpenLayoutGallery: (request) => requests.push(request) });

		await act(async () => {
			latest?.slidesPaneProps.onOpenLayoutForSlide(1, 40, 120);
		});

		expect(latest?.slidesPaneProps.activeSlideIndex).toBe(1);
		expect(requests).toStrictEqual([{ slideIndex: 1, x: 40, y: 120 }]);
	}, 15_000);

	it('still makes the slide active when the host renders no Layout gallery', async () => {
		await mount();

		await act(async () => {
			latest?.slidesPaneProps.onOpenLayoutForSlide(2, 0, 0);
		});

		expect(latest?.slidesPaneProps.activeSlideIndex).toBe(2);
	}, 15_000);

	it('adds a slide from the rail footer, as the full viewer does', async () => {
		const handle = createRef<PowerPointViewerHandle>();
		await mount({ handle });

		await act(async () => {
			container
				.querySelector<HTMLButtonElement>('[data-pptx-chrome="slide-footer"] button')
				?.click();
			await Promise.resolve();
		});
		await flushUntil(() => latest?.slidesPaneProps.slides.length === 4);

		expect(handle.current?.getSlideCount()).toBe(4);
		expect(rows()).toHaveLength(4);
		expect(handle.current?.isDirty()).toBeTruthy();
	}, 15_000);

	it('is read-only without host edit permission', async () => {
		const { slidesPaneProps } = await mount({ canEdit: false });

		expect(slidesPaneProps.canEdit).toBeFalsy();
		expect(rows()).toHaveLength(3);
		expect(container.querySelector('[data-pptx-chrome="slide-footer"]')).toBeNull();
	}, 15_000);
});
