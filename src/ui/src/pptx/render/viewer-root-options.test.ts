import { describe, expect, it } from 'vitest';

import {
	resolveInitialSlideIndex,
	resolveViewerRootOptions,
	VIEWER_ROOT_OPTION_DEFAULTS,
} from './viewer-root-options';

describe('resolveViewerRootOptions', () => {
	it('defaults to the first slide with the toolbar and thumbnails shown', () => {
		expect(resolveViewerRootOptions(undefined)).toEqual({
			initialSlide: 0,
			showToolbar: true,
			showThumbnails: true,
		});
		expect(resolveViewerRootOptions({})).toEqual(VIEWER_ROOT_OPTION_DEFAULTS);
	});

	it('keeps explicit values, including false', () => {
		expect(
			resolveViewerRootOptions({ initialSlide: 3, showToolbar: false, showThumbnails: false }),
		).toEqual({ initialSlide: 3, showToolbar: false, showThumbnails: false });
	});

	it('treats an explicit undefined as unset', () => {
		expect(
			resolveViewerRootOptions({
				initialSlide: undefined,
				showToolbar: undefined,
				showThumbnails: undefined,
			}),
		).toEqual(VIEWER_ROOT_OPTION_DEFAULTS);
	});

	it('falls back to slide 0 for a non-finite initial slide', () => {
		expect(resolveViewerRootOptions({ initialSlide: Number.NaN }).initialSlide).toBe(0);
		expect(resolveViewerRootOptions({ initialSlide: Infinity }).initialSlide).toBe(0);
	});
});

describe('resolveInitialSlideIndex', () => {
	it('returns the requested slide when it is in range', () => {
		expect(resolveInitialSlideIndex(2, 5)).toBe(2);
	});

	it('clamps into the deck', () => {
		expect(resolveInitialSlideIndex(99, 5)).toBe(4);
		expect(resolveInitialSlideIndex(-3, 5)).toBe(0);
	});

	it('truncates fractional indexes', () => {
		expect(resolveInitialSlideIndex(2.9, 5)).toBe(2);
	});

	it('gives 0 for a missing or non-finite value and for an empty deck', () => {
		expect(resolveInitialSlideIndex(undefined, 5)).toBe(0);
		expect(resolveInitialSlideIndex(Number.NaN, 5)).toBe(0);
		expect(resolveInitialSlideIndex(3, 0)).toBe(0);
	});
});
