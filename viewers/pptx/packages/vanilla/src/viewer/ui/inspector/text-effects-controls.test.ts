import { TEXT_WARP_PRESETS } from 'ooxml-ui/pptx';
import type { TextStyle } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../i18n';
import { createTextEffectsControls } from './text-effects-controls';
import { createTextSection } from './text-section';
import type { InspectorHandlers, InspectorState } from './types';

const t = createTranslator();

function state(overrides: Partial<InspectorState> = {}): InspectorState {
	return { hasSelection: true, canText: true, textStyle: {}, ...overrides } as InspectorState;
}

function mount() {
	const setTextStyle = vi.fn();
	const controls = createTextEffectsControls(document, t, {
		setTextStyle,
		pushRecentColor: vi.fn(),
	} as unknown as InspectorHandlers);
	const gallery = controls.el.querySelector('pptx-ui-ribbon-gallery')!;
	const tile = (preset: string) =>
		gallery.querySelector<HTMLButtonElement>(`button.tile[data-gallery-item="${preset}"]`)!;
	const pressed = () =>
		[...gallery.querySelectorAll<HTMLElement>('.tile[aria-pressed="true"]')].map(
			(node) => node.dataset.galleryItem,
		);
	return { controls, gallery, setTextStyle, tile, pressed };
}

describe('text effects Transform (Text Warp) gallery', () => {
	it('draws every warp preset as a tile of the shared compact gallery panel', () => {
		const { controls, gallery } = mount();
		controls.update(state());

		expect(gallery.getAttribute('mode')).toBe('panel');
		expect(gallery.hasAttribute('compact')).toBeTruthy();
		expect(gallery.descriptor?.id).toBe('textWarp');
		expect(controls.el.querySelector('legend')?.textContent).toBe(t('pptx.textEffects.transform'));
		const panel = gallery.querySelector('.gallery-panel[role="group"]')!;
		expect(panel.getAttribute('aria-label')).toBe(t('pptx.textWarp.title'));
		const tiles = [...panel.querySelectorAll<HTMLButtonElement>('button.tile')];
		expect(tiles).toHaveLength(39);
		expect(tiles.map((node) => node.dataset.galleryItem)).toStrictEqual(
			TEXT_WARP_PRESETS.map((preset) => preset.value),
		);
		const wave = tiles.find((node) => node.dataset.galleryItem === 'textWave1')!;
		expect(wave.getAttribute('aria-label')).toBe(t('pptx.textWarp.preset.wave1'));
		expect(wave.title).toBe(t('pptx.textWarp.preset.wave1'));
		expect(gallery.querySelector('.tile-label')).toBeNull();
	});

	it('marks the current preset, and No Transform when the text has none', () => {
		const { controls, pressed } = mount();
		controls.update(state({ textStyle: { textWarpPreset: 'textWave1' } as TextStyle }));
		expect(pressed()).toStrictEqual(['textWave1']);

		controls.update(state({ textStyle: {} as TextStyle }));
		expect(pressed()).toStrictEqual(['textNoShape']);
	});

	it('warps the whole text body with a picked preset, and No Transform clears it', () => {
		const { controls, setTextStyle, tile } = mount();
		controls.update(
			state({
				textStyle: { textWarpPreset: 'textArchUp' } as TextStyle,
				selectedTextRange: { startSegIdx: 0, startOffset: 0, endSegIdx: 0, endOffset: 2 },
			}),
		);

		tile('textWave1').click();
		expect(setTextStyle.mock.calls).toStrictEqual([[{ textWarpPreset: 'textWave1' }]]);

		tile('textNoShape').click();
		expect(setTextStyle.mock.calls[1]).toStrictEqual([{ textWarpPreset: undefined }]);
	});

	it('is disabled while the selection cannot take text', () => {
		const { controls, gallery, setTextStyle, tile } = mount();
		controls.update(state({ canText: false }));

		expect(gallery.disabled).toBeTruthy();
		expect(gallery.descriptor?.disabled).toBeTruthy();
		const tiles = [...gallery.querySelectorAll<HTMLButtonElement>('button.tile')];
		expect(tiles).toHaveLength(39);
		expect(tiles.every((node) => node.disabled)).toBeTruthy();
		tile('textWave1').click();
		expect(setTextStyle).not.toHaveBeenCalled();

		controls.update(state());
		expect(gallery.disabled).toBeFalsy();
		expect(tile('textWave1').disabled).toBeFalsy();
	});

	it('is in the text section with its descriptor set', () => {
		const section = createTextSection(document, t, () => document.createElement('section'), {
			setTextStyle: vi.fn(),
			pushRecentColor: vi.fn(),
		} as unknown as InspectorHandlers);
		section.update(
			state({ textStyle: { textWarpPreset: 'textWave1' } as TextStyle, recentColors: [] }),
		);

		const gallery = section.el.querySelector('pptx-ui-ribbon-gallery')!;
		expect(gallery.descriptor?.id).toBe('textWarp');
		expect(
			gallery.querySelector('[data-gallery-item="textWave1"]')?.getAttribute('aria-pressed'),
		).toBe('true');
	});
});
