import { describe, expect, it, vi } from 'vitest';

import { buildFieldSubstitutionContext } from './field-context';
import {
	buildSlidesPaneFieldContext,
	buildSlidesPaneTableStyleContext,
	createSlidesPaneBulkSlideOps,
	openSlidesPaneLayoutGallery,
	slidesPaneRehearsalTimings,
} from './slides-pane-model';

describe('buildSlidesPaneFieldContext', () => {
	const headerFooter = {
		hasDateTime: true,
		dateTimeText: '1 May',
		hasFooter: true,
		footerText: 'Footer',
	} as const;
	const customProperties = [{ name: 'Client', value: 'Contoso' }] as never;

	it('carries the deck start number, so thumbnails number like the canvas', () => {
		const context = buildSlidesPaneFieldContext({
			headerFooter,
			customProperties,
			firstSlideNumber: 45,
		});

		expect(context.firstSlideNumber).toBe(45);
		expect(context.footerText).toBe('Footer');
		expect(context.dateTimeText).toBe('1 May');
		expect(context.customProperties).toStrictEqual([{ name: 'Client', value: 'Contoso' }]);
	});

	it('is the canvas context without the per-slide number and title', () => {
		const sources = { headerFooter, customProperties, firstSlideNumber: 3, locale: 'fr-FR' };

		expect(buildSlidesPaneFieldContext(sources)).toStrictEqual(
			buildFieldSubstitutionContext({ ...sources, slide: undefined }),
		);
		expect(buildSlidesPaneFieldContext(sources).slideNumber).toBeUndefined();
	});

	it('leaves numbering at its default when the deck declares no start number', () => {
		expect(buildSlidesPaneFieldContext({ headerFooter })).not.toHaveProperty('firstSlideNumber');
	});
});

describe('buildSlidesPaneTableStyleContext', () => {
	it('is undefined when the deck has neither a theme nor table styles', () => {
		expect(buildSlidesPaneTableStyleContext(undefined, undefined)).toBeUndefined();
	});

	it('carries whichever of the theme and style map the deck has', () => {
		const theme = { name: 'Office' };
		const map = { styles: [] };

		expect(buildSlidesPaneTableStyleContext(theme, undefined)).toStrictEqual({
			theme,
			tableStyleMap: undefined,
		});
		expect(buildSlidesPaneTableStyleContext(undefined, map)).toStrictEqual({
			theme: undefined,
			tableStyleMap: map,
		});
	});
});

describe('slidesPaneRehearsalTimings', () => {
	it('is undefined for an unrehearsed deck', () => {
		expect(slidesPaneRehearsalTimings(undefined)).toBeUndefined();
		expect(slidesPaneRehearsalTimings({})).toBeUndefined();
	});

	it('copies recorded timings so the rail cannot mutate the source', () => {
		const recorded = { 0: 12, 2: 30 };
		const timings = slidesPaneRehearsalTimings(recorded);

		expect(timings).toStrictEqual(recorded);
		expect(timings).not.toBe(recorded);
	});
});

describe('openSlidesPaneLayoutGallery', () => {
	it('makes the slide active, then hands the request to the host', () => {
		const calls: string[] = [];
		const setActive = vi.fn((index: number) => calls.push(`active ${index}`));
		const onOpen = vi.fn(() => calls.push('open'));

		openSlidesPaneLayoutGallery({ slideIndex: 2, x: 40, y: 120 }, setActive, onOpen);

		expect(calls).toStrictEqual(['active 2', 'open']);
		expect(onOpen).toHaveBeenCalledWith({ slideIndex: 2, x: 40, y: 120 });
	});

	it('still makes the slide active without a host gallery', () => {
		const setActive = vi.fn();

		openSlidesPaneLayoutGallery({ slideIndex: 1, x: 0, y: 0 }, setActive, undefined);

		expect(setActive).toHaveBeenCalledWith(1);
	});
});

describe('createSlidesPaneBulkSlideOps', () => {
	/** A deck of slide ids with the single-slide operations a binding has. */
	function deck(ids: string[]) {
		const slides = [...ids];
		const hidden = new Set<string>();
		const ops = createSlidesPaneBulkSlideOps({
			duplicateSlide: (index) => slides.splice(index + 1, 0, `${slides[index]}'`),
			deleteSlide: (index) => slides.splice(index, 1),
			toggleSlideHidden: (index) => {
				const id = slides[index];
				if (hidden.has(id)) {
					hidden.delete(id);
				} else {
					hidden.add(id);
				}
			},
		});
		return { slides, hidden, ops };
	}

	it('duplicates each selected slide right after itself', () => {
		const { slides, ops } = deck(['a', 'b', 'c', 'd']);

		ops.duplicateSlides([0, 2]);

		expect(slides).toStrictEqual(['a', "a'", 'b', 'c', "c'", 'd']);
	});

	it('deletes exactly the selected slides whatever the selection order', () => {
		const { slides, ops } = deck(['a', 'b', 'c', 'd']);

		ops.deleteSlides([1, 3, 1]);

		expect(slides).toStrictEqual(['a', 'c']);
	});

	it('toggles each selected slide once', () => {
		const { hidden, ops } = deck(['a', 'b', 'c']);

		ops.toggleHideSlides([0, 2, 2]);

		expect([...hidden].toSorted()).toStrictEqual(['a', 'c']);
	});
});
