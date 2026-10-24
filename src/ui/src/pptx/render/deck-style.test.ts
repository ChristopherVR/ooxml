import type { PptxElement, PptxSlide, PptxThemeColorScheme } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import {
	commitGalleryResult,
	commitsThroughDeckStyle,
	createGalleryPreviewController,
	previewGalleryResult,
} from './deck-style';
import type { DeckStyleHandler, DeckStyleState } from './deck-style';
import { settleThemeWrites } from './theme-history';

const OLD: PptxThemeColorScheme = {
	dk1: '#000000',
	lt1: '#FFFFFF',
	dk2: '#111111',
	lt2: '#EEEEEE',
	accent1: '#FF0000',
	accent2: '#00FF00',
	accent3: '#0000FF',
	accent4: '#FFFF00',
	accent5: '#FF00FF',
	accent6: '#00FFFF',
	hlink: '#0000EE',
	folHlink: '#551A8B',
};
const NEW: PptxThemeColorScheme = { ...OLD, accent1: '#AA0000' };

function slide(id: string, fill: string): PptxSlide {
	return {
		id,
		rId: id,
		slideNumber: 1,
		elements: [
			{
				id: `${id}-shape`,
				type: 'shape',
				x: 0,
				y: 0,
				width: 10,
				height: 10,
				shapeStyle: { fillColor: fill },
			} as unknown as PptxElement,
		],
	};
}

function state(): DeckStyleState {
	return {
		slides: [slide('a', '#FF0000')],
		templateElementsBySlideId: {},
		theme: { name: 'T', colorScheme: OLD },
		themeColorMap: undefined,
	};
}

function fillOf(deck: DeckStyleState): string | undefined {
	const element = deck.slides[0]!.elements[0]!;
	return 'shapeStyle' in element ? element.shapeStyle?.fillColor : undefined;
}

function fakeHandler(): DeckStyleHandler {
	return {
		applyTheme: vi.fn(async () => {}),
		updateThemeColorScheme: vi.fn(async () => {}),
		updateThemeFontScheme: vi.fn(async () => {}),
		applyThemeEffectScheme: vi.fn(async (id: string) => ({
			name: id,
			fillStyles: [],
			lineStyles: [],
			effectStyles: [],
			backgroundFillStyles: [],
		})),
		applyThemeFormatSchemeXml: vi.fn(async () => undefined),
		reResolveThemeStyleReferences: vi.fn((slides, templates) => ({
			slides: slides.map((s) => ({ ...s, name: 'reresolved' })),
			templateElementsBySlideId: { ...templates },
		})),
		applyBackgroundStyleToSlides: vi.fn((slides, templates, index) => ({
			slides: slides.map((s) => ({ ...s, name: `style-${index}` })),
			templateElementsBySlideId: { ...templates },
		})),
		commitMasterBackgroundStyle: vi.fn(async () => true),
	};
}

const target = { selectedSlideIds: ['a'] };

describe('previewGalleryResult', () => {
	it('re-resolves colours for a colour scheme without touching the handler', () => {
		const handler = fakeHandler();
		const out = previewGalleryResult(
			{ kind: 'themeColorScheme', colorScheme: NEW, name: 'T' },
			state(),
			handler,
			target,
		)!;
		expect(fillOf(out)).toBe('#AA0000');
		expect(out.theme?.colorScheme).toEqual(NEW);
		expect(out.themeColorMap?.['accent1']).toBe('AA0000');
		expect(handler.updateThemeColorScheme).not.toHaveBeenCalled();
	});

	it('previews effects against the named scheme and backgrounds with the scope', () => {
		const handler = fakeHandler();
		const effects = previewGalleryResult(
			{ kind: 'themeEffects', schemeId: 'glossy', name: 'Glossy' },
			state(),
			handler,
			target,
		)!;
		expect(effects.slides[0]!.name).toBe('reresolved');
		expect(handler.reResolveThemeStyleReferences).toHaveBeenCalledWith(
			expect.anything(),
			expect.anything(),
			'glossy',
		);
		previewGalleryResult(
			{ kind: 'backgroundStyle', styleIndex: 4, scope: 'selected' },
			state(),
			handler,
			target,
		);
		expect(handler.applyBackgroundStyleToSlides).toHaveBeenCalledWith(
			expect.anything(),
			expect.anything(),
			4,
			{ kind: 'slides', slideIds: ['a'] },
		);
	});

	it('has nothing to show for commands, and only for schemes it can name', () => {
		expect(
			previewGalleryResult({ kind: 'command', command: 'formatBackground' }, state(), null, target),
		).toBeUndefined();
		expect(
			previewGalleryResult(
				{ kind: 'themeEffects', fmtSchemeXml: '<a:fmtScheme/>', name: 'x' },
				state(),
				fakeHandler(),
				target,
			),
		).toBeUndefined();
	});

	it('folds a sequence in order', () => {
		const out = previewGalleryResult(
			{
				kind: 'sequence',
				steps: [
					{
						kind: 'theme',
						colorScheme: NEW,
						fontScheme: { majorFont: { latin: 'Arial' } },
						name: 'Renamed',
					},
					{ kind: 'backgroundStyle', styleIndex: 2, scope: 'all' },
				],
			},
			state(),
			fakeHandler(),
			target,
		)!;
		expect(fillOf(out)).toBe('#AA0000');
		expect(out.theme?.name).toBe('Renamed');
		expect(out.slides[0]!.name).toBe('style-2');
	});
});

describe('commitGalleryResult', () => {
	it('writes the effect scheme to the handler and records it on the theme', async () => {
		const handler = fakeHandler();
		const out = (await commitGalleryResult(
			{ kind: 'themeEffects', schemeId: 'glossy', name: 'Glossy' },
			state(),
			handler,
			target,
		))!;
		expect(handler.applyThemeEffectScheme).toHaveBeenCalledWith('glossy');
		expect(out.theme?.formatScheme?.name).toBe('glossy');
		expect(out.slides[0]!.name).toBe('reresolved');
	});

	it('saves a deck-wide background on the master, not a selected-slides one', async () => {
		const handler = fakeHandler();
		await commitGalleryResult(
			{ kind: 'backgroundStyle', styleIndex: 3, scope: 'selected' },
			state(),
			handler,
			target,
		);
		expect(handler.commitMasterBackgroundStyle).not.toHaveBeenCalled();
		await commitGalleryResult(
			{ kind: 'backgroundStyle', styleIndex: 3, scope: 'all' },
			state(),
			handler,
			target,
		);
		expect(handler.commitMasterBackgroundStyle).toHaveBeenCalledWith(3);
	});

	it('writes only the font scheme for a fonts pick, keeping the colours and the name', async () => {
		// `applyTheme` would also rewrite the colour scheme (a `sysClr` slot becomes `srgbClr`).
		const handler = fakeHandler();
		const fontScheme = { majorFont: { latin: 'Arial' }, minorFont: { latin: 'Arial' } };
		const out = (await commitGalleryResult(
			{ kind: 'themeFontScheme', fontScheme, name: 'Arial' },
			state(),
			handler,
			target,
		))!;
		expect(handler.updateThemeFontScheme).toHaveBeenCalledExactlyOnceWith(fontScheme);
		expect(handler.updateThemeColorScheme).not.toHaveBeenCalled();
		expect(handler.applyTheme).not.toHaveBeenCalled();
		expect(out.theme).toStrictEqual({ name: 'T', colorScheme: OLD, fontScheme });
	});

	it('queues its theme part writes, so a save asked for mid-pick waits for them', async () => {
		const handler = fakeHandler();
		const order: string[] = [];
		handler.applyTheme = vi.fn(async () => {
			await new Promise<void>((done) => {
				setTimeout(done, 30);
			});
			order.push('theme written');
		});
		const commit = commitGalleryResult(
			{ kind: 'theme', colorScheme: NEW, fontScheme: {}, name: 'New' },
			state(),
			handler,
			target,
		);
		await settleThemeWrites(handler).then(() => order.push('save'));
		await commit;
		expect(order).toStrictEqual(['theme written', 'save']);
	});

	it('does nothing without a handler', async () => {
		expect(
			await commitGalleryResult(
				{ kind: 'backgroundStyle', styleIndex: 3, scope: 'all' },
				state(),
				null,
				target,
			),
		).toBeUndefined();
	});
});

describe('commitsThroughDeckStyle', () => {
	it('takes every Design gallery pick that changes the theme or the backgrounds', () => {
		const picks = [
			{ kind: 'theme', colorScheme: NEW, fontScheme: {}, name: 'New' },
			{ kind: 'themeColorScheme', colorScheme: NEW, name: 'New' },
			{ kind: 'themeFontScheme', fontScheme: {}, name: 'New' },
			{ kind: 'themeEffects', schemeId: 'glossy', name: 'Glossy' },
			{ kind: 'backgroundStyle', styleIndex: 3, scope: 'all' },
			{ kind: 'sequence', steps: [] },
		] as const;
		for (const pick of picks) {
			expect(commitsThroughDeckStyle(pick), pick.kind).toBe(true);
		}
		expect(commitsThroughDeckStyle({ kind: 'element', elementId: 'a', patch: {} })).toBe(false);
	});
});

describe('createGalleryPreviewController', () => {
	function setup() {
		let live = state();
		const shown: DeckStyleState[] = [];
		const hold = vi.fn();
		const controller = createGalleryPreviewController({
			read: () => live,
			show: (next) => {
				live = next;
				shown.push(next);
			},
			hold,
		});
		return { controller, hold, shown, live: () => live };
	}

	it('shows previews over the live state and puts the original back', () => {
		const { controller, hold, live } = setup();
		const original = live();
		controller.preview({ ...original, theme: { name: 'Preview' } });
		controller.preview({ ...original, theme: { name: 'Next' } });
		expect(live().theme?.name).toBe('Next');
		expect(hold.mock.calls).toEqual([[true]]);
		expect(controller.active).toBe(true);
		expect(controller.base()).toBe(original);
		controller.end();
		expect(live()).toBe(original);
		expect(hold.mock.calls).toEqual([[true], [false]]);
		expect(controller.active).toBe(false);
	});

	it('settles on the original for a commit, and ends a preview given nothing', () => {
		const { controller, live } = setup();
		const original = live();
		controller.preview({ ...original, theme: { name: 'Preview' } });
		expect(controller.settle()).toBe(original);
		expect(live()).toBe(original);
		expect(controller.settle()).toBeUndefined();
		controller.preview({ ...original, theme: { name: 'Again' } });
		controller.preview(undefined);
		expect(live()).toBe(original);
	});
});
