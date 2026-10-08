// @vitest-environment happy-dom
/**
 * The root display props `showToolbar`, `showThumbnails` and `initialSlide`,
 * end to end through the real `PowerPointViewer`. The defaults and clamping
 * come from the shared `resolveViewerRootOptions` / `resolveInitialSlideIndex`
 * (ooxml-ui/pptx), so these props behave as they do in every other binding.
 */
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import React, { act, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

// oxlint-disable-next-line prefer-ending-with-an-expect
vi.mock<typeof import('react-i18next')>(import('react-i18next'), () => ({
	useTranslation: () => ({
		t: (key: string) => translationsEn[key] ?? key,
		i18n: {
			language: 'en',
			languages: ['en'],
			options: { resources: { en: {} } },
			changeLanguage: () => Promise.resolve(),
		},
	}),
}));

const { PptxHandler } = await import('pptx-viewer-core');
const { PowerPointViewer } = await import('../index');
type ViewerHandle = import('../index').PowerPointViewerHandle;
type ViewerProps = import('../index').PowerPointViewerProps;

async function sampleDeck(slideCount: number): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: slideCount });
	try {
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	// A desktop viewport, so the ribbon and the slides pane (not the mobile bar) render.
	window.happyDOM?.setViewport({ width: 1600, height: 950 });
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

async function mount(
	props: Partial<Omit<ViewerProps, 'content'>>,
	slideCount = 1,
): Promise<ViewerHandle> {
	const content = await sampleDeck(slideCount);
	const ref = createRef<ViewerHandle>();
	await act(async () => {
		root.render(<PowerPointViewer ref={ref} content={content} canEdit {...props} />);
	});
	for (let attempt = 0; attempt < 80; attempt += 1) {
		if ((ref.current?.getSlideCount() ?? 0) > 0) {
			// One more tick so the loaded deck's chrome commits.
			await act(async () => {
				await new Promise((resolve) => {
					setTimeout(resolve, 20);
				});
			});
			return ref.current as ViewerHandle;
		}
		await act(async () => {
			await new Promise((resolve) => {
				setTimeout(resolve, 20);
			});
		});
	}
	throw new Error('the viewer never finished loading its deck');
}

const ribbon = (): Element | null => container.querySelector('[data-pptx-chrome="ribbon"]');
const slidesPane = (): Element | null => container.querySelector('[data-pptx-chrome="slides"]');
const statusBar = (): Element | null => container.querySelector('pptx-ui-status-bar');

describe('powerPointViewer root display props', () => {
	it('shows the toolbar, status bar and thumbnails and opens slide 0 by default', async () => {
		const handle = await mount({}, 3);
		expect(ribbon()).not.toBeNull();
		expect(statusBar()).not.toBeNull();
		expect(slidesPane()).not.toBeNull();
		expect(handle.getActiveSlideIndex()).toBe(0);
	});

	it('hides the ribbon and status bar when showToolbar is false', async () => {
		await mount({ showToolbar: false });
		expect(ribbon()).toBeNull();
		expect(statusBar()).toBeNull();
		expect(slidesPane()).not.toBeNull();
	});

	it('hides the slide thumbnail pane when showThumbnails is false', async () => {
		await mount({ showThumbnails: false });
		expect(slidesPane()).toBeNull();
		expect(ribbon()).not.toBeNull();
	});

	it('opens the deck on initialSlide', async () => {
		const handle = await mount({ initialSlide: 2 }, 3);
		expect(handle.getActiveSlideIndex()).toBe(2);
	});

	it('clamps an out-of-range initialSlide into the deck', async () => {
		const handle = await mount({ initialSlide: 99 }, 3);
		expect(handle.getActiveSlideIndex()).toBe(2);
	});
});
