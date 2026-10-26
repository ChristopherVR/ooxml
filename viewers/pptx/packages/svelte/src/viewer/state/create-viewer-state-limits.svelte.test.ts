import type { PptxLoadLimits } from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import {
	LOAD_LIMIT_DECK_SLIDES,
	loadLimitDeck,
	OVER_SLIDE_LIMIT_MESSAGE,
} from '../../../../shared/test-utils/load-limits.mjs';
import { translate } from '../../i18n/translator';
import type { ViewerStateBag } from './create-viewer-state-types';
import CreateViewerStateHarness from './CreateViewerStateHarness.svelte';
import { toViewerStateOptions } from './viewer-state-options';

let cleanup: (() => void) | undefined;
beforeEach(() => localStorage.clear());
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
	vi.restoreAllMocks();
});

function renderHarness(limits: PptxLoadLimits | undefined): ViewerStateBag {
	let captured: ViewerStateBag | undefined;
	const instance = mount(CreateViewerStateHarness, {
		target: document.createElement('div'),
		props: {
			limits,
			t: (key: string, params?: Record<string, string | number>) => translate('en', key, params),
			onready: (state: ViewerStateBag) => {
				captured = state;
			},
		},
	});
	cleanup = () => unmount(instance);
	flushSync();
	return captured!;
}

describe('createViewerState host limits', () => {
	it('maps the limits prop live from the props', () => {
		let props: { limits?: PptxLoadLimits } = { limits: { maxSlides: 1 } };
		const options = toViewerStateOptions(() => props as never, {
			t: (key) => key,
			getRootEl: () => undefined,
			getStageHolderEl: () => undefined,
			getViewportWidth: () => 960,
			getViewportHeight: () => 540,
			getMasterScale: () => 1,
		});
		expect(options.getLimits?.()).toStrictEqual({ maxSlides: 1 });
		props = {};
		expect(options.getLimits?.()).toBeUndefined();
	});

	it('passes the limits to PptxHandler.load and shows the localized limit error', async () => {
		const limits = { maxSlides: LOAD_LIMIT_DECK_SLIDES - 1 };
		const state = renderHarness(limits);
		const deck = await loadLimitDeck(PptxHandler);
		const load = vi.spyOn(PptxHandler.prototype, 'load');
		await state.loader.load(deck);
		expect(load).toHaveBeenCalledWith(expect.any(ArrayBuffer), expect.objectContaining({ limits }));
		expect(state.loader.error).toBe(OVER_SLIDE_LIMIT_MESSAGE);
		expect(state.loader.slides).toHaveLength(0);
	});

	it('loads within the limits, and passes undefined when the prop is unset', async () => {
		const within = renderHarness({ maxSlides: LOAD_LIMIT_DECK_SLIDES });
		await within.loader.load(await loadLimitDeck(PptxHandler));
		expect(within.loader.error).toBeNull();
		expect(within.loader.slides).toHaveLength(LOAD_LIMIT_DECK_SLIDES);
		cleanup?.();

		const unset = renderHarness(undefined);
		const deck = await loadLimitDeck(PptxHandler);
		const load = vi.spyOn(PptxHandler.prototype, 'load');
		await unset.loader.load(deck);
		expect(load.mock.calls[0]?.[1]).toHaveProperty('limits', undefined);
		expect(unset.loader.slides).toHaveLength(LOAD_LIMIT_DECK_SLIDES);
	});
});
