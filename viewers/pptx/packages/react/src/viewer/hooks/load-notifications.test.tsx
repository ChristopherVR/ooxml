// @vitest-environment happy-dom
/**
 * Host `onLoad` / `onError` are wired to the real load pipeline: a valid deck reports its
 * slide count and canvas size once, and unreadable bytes report a message without `onLoad`.
 */
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import React, { act } from 'react';
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
const { PowerPointViewer } = await import('../../index');

async function sampleDeck(initialSlideCount: number): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount });
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

async function settle(done: () => boolean): Promise<void> {
	for (let attempt = 0; attempt < 100 && !done(); attempt += 1) {
		await act(async () => {
			await new Promise((resolve) => {
				setTimeout(resolve, 20);
			});
		});
	}
}

describe('onLoad / onError', () => {
	it('reports the slide count and canvas size once a deck has loaded', async () => {
		const content = await sampleDeck(3);
		const onLoad = vi.fn();
		const onError = vi.fn();
		await act(async () => {
			root.render(<PowerPointViewer content={content} onLoad={onLoad} onError={onError} />);
		});
		await settle(() => onLoad.mock.calls.length > 0);

		expect(onLoad).toHaveBeenCalledTimes(1);
		const detail = onLoad.mock.calls[0]?.[0] as {
			slideCount: number;
			canvasSize: { width: number };
		};
		expect(detail.slideCount).toBe(3);
		expect(detail.canvasSize.width).toBeGreaterThan(0);
		expect(onError).not.toHaveBeenCalled();
	});

	it('reports a message, and no load, when the bytes are not a presentation', async () => {
		const onLoad = vi.fn();
		const onError = vi.fn();
		const errorSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
		try {
			await act(async () => {
				root.render(
					<PowerPointViewer
						content={new Uint8Array([1, 2, 3, 4])}
						onLoad={onLoad}
						onError={onError}
					/>,
				);
			});
			await settle(() => onError.mock.calls.length > 0);
		} finally {
			errorSpy.mockRestore();
		}

		expect(onError).toHaveBeenCalledTimes(1);
		expect(typeof onError.mock.calls[0]?.[0]).toBe('string');
		expect(onLoad).not.toHaveBeenCalled();
	});
});
