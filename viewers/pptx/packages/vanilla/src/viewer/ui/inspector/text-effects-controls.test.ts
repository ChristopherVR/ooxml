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

describe('text shadow, glow and reflection', () => {
	/** The checkbox a caption names (each toggle carries its caption as its accessible name). */
	function toggleNamed(root: HTMLElement, caption: string) {
		const box = root.querySelector<HTMLElement & { checked: boolean; disabled: boolean }>(
			`pptx-ui-checkbox[aria-label="${caption}"]`,
		);
		if (!box) {
			throw new Error(`the text effects have no "${caption}" toggle`);
		}
		return box;
	}

	function tick(box: HTMLElement & { checked: boolean }, checked: boolean): void {
		box.checked = checked;
		box.dispatchEvent(new Event('change', { bubbles: true }));
	}

	it('switches each effect on with the defaults the other bindings write', () => {
		const { controls, setTextStyle } = mount();
		controls.update(state());
		tick(toggleNamed(controls.el, t('pptx.textEffects.shadow')), true);
		tick(toggleNamed(controls.el, t('pptx.textEffects.glow')), true);
		tick(toggleNamed(controls.el, t('pptx.textEffects.reflection')), true);
		expect(setTextStyle.mock.calls.map(([patch]) => patch)).toStrictEqual([
			{
				textShadowColor: '#000000',
				textShadowBlur: 4,
				textShadowOffsetX: 2,
				textShadowOffsetY: 2,
				textShadowOpacity: 0.5,
			},
			{ textGlowColor: '#ffff00', textGlowRadius: 6, textGlowOpacity: 0.6 },
			{
				textReflection: true,
				textReflectionBlur: 1,
				textReflectionStartOpacity: 0.5,
				textReflectionEndOpacity: 0,
				textReflectionOffset: 3,
			},
		]);
	});

	it('shows a glowing run as glowing, edits its size and clears it, on the selected range', () => {
		const { controls, setTextStyle } = mount();
		const range = { startSegIdx: 0, startOffset: 0, endSegIdx: 0, endOffset: 2 };
		controls.update(
			state({
				textStyle: { textGlowColor: '#ff0000', textGlowRadius: 8 } as TextStyle,
				selectedTextRange: range,
			}),
		);
		const glow = toggleNamed(controls.el, t('pptx.textEffects.glow'));
		expect(glow.checked).toBeTruthy();
		const size = controls.el.querySelector<HTMLInputElement>('[data-testid="text-fx-glow-size"]')!;
		expect(size.value).toBe('8');
		size.value = '400';
		size.dispatchEvent(new Event('change', { bubbles: true }));
		tick(glow, false);
		expect(setTextStyle.mock.calls).toStrictEqual([
			[{ textGlowRadius: 50 }, range],
			[{ textGlowColor: undefined, textGlowRadius: undefined, textGlowOpacity: undefined }, range],
		]);
	});

	it('keeps the effect toggles disabled while the selection cannot take text', () => {
		const { controls } = mount();
		controls.update(state({ canText: false }));
		expect(toggleNamed(controls.el, t('pptx.textEffects.glow')).disabled).toBeTruthy();
	});
});
