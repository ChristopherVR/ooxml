import type { PptxLoadLimits } from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import {
	LOAD_LIMIT_DECK_SLIDES,
	loadLimitDeck,
	OVER_SLIDE_LIMIT_MESSAGE,
} from '../../../shared/test-utils/load-limits.mjs';
import { createPptxViewer } from './PptxViewer';
import type { PptxViewerInstance } from './types';

let active: PptxViewerInstance[] = [];

afterEach(async () => {
	for (const viewer of active) {
		viewer.destroy();
	}
	active = [];
	document.body.replaceChildren();
	vi.restoreAllMocks();
	await new Promise<void>((resolve) => {
		setTimeout(resolve, 0);
	});
});

async function loadWith(limits: PptxLoadLimits | undefined) {
	const deck = await loadLimitDeck(PptxHandler);
	const load = vi.spyOn(PptxHandler.prototype, 'load');
	const onError = vi.fn();
	const container = document.createElement('div');
	document.body.appendChild(container);
	const viewer = createPptxViewer(container, { limits, onError, googleFonts: false });
	active.push(viewer);
	await viewer.loadFile(deck);
	return { container, viewer, load, onError };
}

describe('createPptxViewer limits option', () => {
	it('passes the limits to PptxHandler.load and shows the localized limit error', async () => {
		const limits = { maxSlides: LOAD_LIMIT_DECK_SLIDES - 1 };
		const { container, load, onError } = await loadWith(limits);
		expect(load).toHaveBeenCalledWith(expect.any(ArrayBuffer), expect.objectContaining({ limits }));
		expect(container.querySelector<HTMLElement>('.pptxv-error')?.hidden).toBeFalsy();
		expect(container.querySelector('.pptxv-error-message')?.textContent).toBe(
			OVER_SLIDE_LIMIT_MESSAGE,
		);
		expect(onError).toHaveBeenCalledWith(
			OVER_SLIDE_LIMIT_MESSAGE,
			expect.objectContaining({ code: 'PPTX_LOAD_LIMIT', limit: 'maxSlides' }),
		);
	});

	it('loads within the limits, and passes undefined when the option is unset', async () => {
		const within = await loadWith({ maxSlides: LOAD_LIMIT_DECK_SLIDES });
		expect(within.onError).not.toHaveBeenCalled();
		expect(within.viewer.getHandler()).not.toBeNull();

		vi.restoreAllMocks();
		const unset = await loadWith(undefined);
		expect(unset.load.mock.calls[0]?.[1]).toHaveProperty('limits', undefined);
		expect(unset.onError).not.toHaveBeenCalled();
	});
});
