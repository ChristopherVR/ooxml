// @vitest-environment happy-dom
import { PptxHandler } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { PropertyHandlersResult, UsePropertyHandlersInput } from './usePropertyHandlers';
import { usePropertyHandlers } from './usePropertyHandlers';

(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT =
	true;

/** Answer the next file picker with a deck and capture the options Compare loads it with. */
function pickDeckAndCaptureLoad() {
	vi.spyOn(HTMLInputElement.prototype, 'click').mockImplementation(function (
		this: HTMLInputElement,
	) {
		Object.defineProperty(this, 'files', { value: [new File(['pptx'], 'other.pptx')] });
		void this.onchange?.(new Event('change'));
	});
	return vi.spyOn(PptxHandler.prototype, 'load').mockResolvedValue({ slides: [] } as never);
}

let root: Root;
let latest: PropertyHandlersResult;

function Harness(props: { allowExternalImages?: boolean }) {
	latest = usePropertyHandlers({
		slides: [],
		activeSlideIndex: 0,
		canvasSize: { width: 960, height: 540 },
		setContent: vi.fn(),
		setPresentationProperties: vi.fn(),
		setCoreProperties: vi.fn(),
		setAppProperties: vi.fn(),
		setCustomProperties: vi.fn(),
		setSlides: vi.fn(),
		setIsDirty: vi.fn(),
		ops: {},
		history: {},
		allowExternalImages: props.allowExternalImages,
	} as unknown as UsePropertyHandlersInput);
	return <div />;
}

async function compareWith(allowExternalImages: boolean | undefined) {
	const load = pickDeckAndCaptureLoad();
	root = createRoot(document.createElement('div'));
	await act(async () => root.render(<Harness allowExternalImages={allowExternalImages} />));
	await act(async () => latest.handleCompare());
	await vi.waitFor(() => expect(load).toHaveBeenCalled());
	return load.mock.calls[0][1];
}

describe('usePropertyHandlers Compare linked content', () => {
	afterEach(async () => {
		await act(async () => root?.unmount());
		vi.restoreAllMocks();
	});

	it('loads the incoming deck with the resolved external-content switch', async () => {
		await expect(compareWith(true)).resolves.toStrictEqual({ allowExternalImages: true });
	});

	it('blocks linked content by default', async () => {
		await expect(compareWith(undefined)).resolves.toStrictEqual({ allowExternalImages: false });
	});
});
