import { mount } from '@vue/test-utils';
import type { PptxElement, TextStyle } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';
import { ref } from 'vue';

import { RecentColorsKey } from '../../composables/recent-colors-context';
import { ThemeColorMapKey } from '../../composables/theme-color-map-context';
import { setControlValue } from './test-control-value';
import { themeSwatch } from './test-theme-swatch';
import TextPanel from './TextPanel.vue';

const OFFICE_THEME = {
	dk1: '#000000',
	lt1: '#FFFFFF',
	dk2: '#44546A',
	lt2: '#E7E6E6',
	accent1: '#4472C4',
	accent2: '#ED7D31',
	accent3: '#A5A5A5',
	accent4: '#FFC000',
	accent5: '#5B9BD5',
	accent6: '#70AD47',
	bg1: '#FFFFFF',
	tx1: '#000000',
	bg2: '#E7E6E6',
	tx2: '#44546A',
};

function textEl(textStyle?: TextStyle): PptxElement {
	return {
		type: 'text',
		id: 't1',
		x: 0,
		y: 0,
		width: 100,
		height: 40,
		textStyle,
	} as PptxElement;
}

describe('textPanel', () => {
	it('shows a muted note for elements without text properties', () => {
		const wrapper = mount(TextPanel, {
			props: { element: { type: 'image', id: 'i1' } as PptxElement },
		});
		expect(wrapper.find('.pptx-vue-text-muted').exists()).toBeTruthy();
		expect(wrapper.find('pptx-ui-select').exists()).toBeFalsy();
	});

	it('renders current font size and color', () => {
		const wrapper = mount(TextPanel, {
			props: {
				element: textEl({
					fontSize: 48.1 * (96 / 72),
					color: '#112233',
					fontFamily: 'Georgia',
				}),
			},
		});
		const size = wrapper.find('input[type="number"]').element as HTMLInputElement;
		expect(size.value).toBe('48.1');
		expect(size.step).toBe('any');
		expect((wrapper.find('input[type="color"]').element as HTMLInputElement).value).toBe('#112233');
		expect((wrapper.find('pptx-ui-select').element as HTMLSelectElement).value).toBe('Georgia');
	});

	it('emits a full merged textStyle patch when size changes', async () => {
		const element = textEl({ bold: true, color: '#000000' });
		element.textSegments = [{ text: 'Hello', style: { fontSize: 16, italic: true } }];
		const wrapper = mount(TextPanel, { props: { element } });
		const num = wrapper.find('input[type="number"]');
		(num.element as HTMLInputElement).value = '48.1';
		await num.trigger('input');

		const patch = wrapper.emitted('update')?.[0]?.[0] as Partial<PptxElement>;
		expect(patch.textStyle).toMatchObject({ bold: true, color: '#000000' });
		expect(patch.textStyle?.fontSize).toBeCloseTo(48.1 * (96 / 72));
		expect(patch.textSegments?.[0]?.style).toMatchObject({ italic: true });
		expect(patch.textSegments?.[0]?.style.fontSize).toBeCloseTo(48.1 * (96 / 72));
	});

	it('toggles bold against current textStyle', async () => {
		const wrapper = mount(TextPanel, { props: { element: textEl({ bold: true }) } });
		const boldBtn = wrapper.findAll('.pptx-vue-text-toggle')[0];
		await boldBtn.trigger('click');

		const patch = wrapper.emitted('update')?.[0]?.[0] as Partial<PptxElement>;
		expect(patch).toStrictEqual({ textStyle: { bold: false } });
	});

	it('emits align and vAlign field names', async () => {
		const wrapper = mount(TextPanel, { props: { element: textEl({}) } });
		const buttons = wrapper.findAll('.pptx-vue-text-toggle');
		// toggles: B I U S | Left Center Right Justify; vertical alignment is the
		// shared Text Box options' drop-down.
		await buttons[5].trigger('click'); // Center
		await setControlValue(wrapper.find('[data-testid="text-box-valign"]'), 'middle');

		const events = wrapper.emitted('update');
		expect(events?.[0]?.[0]).toStrictEqual({ textStyle: { align: 'center' } });
		expect(events?.[1]?.[0]).toStrictEqual({ textStyle: { vAlign: 'middle' } });
	});

	it('pushes a committed text colour onto the injected recent-colours list and offers it back', async () => {
		const recent = ref<string[]>(['#112233']);
		const push = (hex: string): void => {
			recent.value = [hex, ...recent.value.filter((c) => c !== hex)];
		};
		const wrapper = mount(TextPanel, {
			props: { element: textEl({ color: '#000000' }) },
			global: { provide: { [RecentColorsKey as symbol]: { recent, push } } },
		});

		expect(wrapper.find('[data-testid="pptx-color-recent"]').exists()).toBeTruthy();

		const color = wrapper.find('input[type="color"]');
		await setControlValue(color, '#00ff00');
		expect(recent.value[0]).toBe('#00ff00');
	});
});

describe('textPanel theme colour picker', () => {
	it('commits both the resolved hex and the ref on a theme swatch click', async () => {
		const wrapper = mount(TextPanel, {
			props: { element: textEl({ color: '#000000' }) },
			global: { provide: { [ThemeColorMapKey as symbol]: ref(OFFICE_THEME) } },
		});
		const accent1 = themeSwatch(wrapper, 'Accent 1');
		await accent1.trigger('click');
		const patch = wrapper.emitted('update')?.[0]?.[0] as Partial<PptxElement>;
		expect(patch.textStyle).toMatchObject({ color: '#4472c4', colorRef: { scheme: 'accent1' } });
	});

	it('clears a previously-stored ref when the native colour input changes', async () => {
		const wrapper = mount(TextPanel, {
			props: {
				element: textEl({ color: '#4472c4', colorRef: { scheme: 'accent1' } }),
			},
			global: { provide: { [ThemeColorMapKey as symbol]: ref(OFFICE_THEME) } },
		});
		const color = wrapper.find('input[type="color"]');
		await setControlValue(color, '#ff0000');
		const patch = wrapper.emitted('update')?.[0]?.[0] as Partial<PptxElement>;
		expect(patch.textStyle?.colorRef).toBeFalsy();
	});
});

describe('textPanel text warp', () => {
	function warpTiles(wrapper: ReturnType<typeof mount>) {
		return wrapper.findAll<HTMLButtonElement>('pptx-ui-ribbon-gallery [data-gallery-item]');
	}

	it('renders the shared Text Warp gallery with the current preset marked', () => {
		const wrapper = mount(TextPanel, {
			props: { element: textEl({ textWarpPreset: 'textArchUp' }) },
		});
		const tiles = warpTiles(wrapper);
		expect(tiles).toHaveLength(39);
		expect(
			tiles
				.filter((tile) => tile.attributes('aria-pressed') === 'true')
				.map((tile) => tile.attributes('data-gallery-item')),
		).toStrictEqual(['textArchUp']);
	});

	it('merges a picked warp into the full textStyle patch, and No Transform clears it', async () => {
		const wrapper = mount(TextPanel, {
			props: { element: textEl({ bold: true, textWarpPreset: 'textArchUp' }) },
		});
		await wrapper.get('.pptx-vue-textwarp-toggle').trigger('click');
		wrapper.get<HTMLButtonElement>('[data-gallery-item="textInflate"]').element.click();
		wrapper.get<HTMLButtonElement>('[data-gallery-item="textNoShape"]').element.click();
		expect(wrapper.emitted('update')).toStrictEqual([
			[{ textStyle: { bold: true, textWarpPreset: 'textInflate' } }],
			[{ textStyle: { bold: true, textWarpPreset: undefined } }],
		]);
	});

	it('keeps the warp tiles unpickable when the deck is read-only', () => {
		const wrapper = mount(TextPanel, {
			props: { element: textEl({ textWarpPreset: 'textArchUp' }), canEdit: false },
		});
		expect(warpTiles(wrapper).every((tile) => tile.element.disabled)).toBeTruthy();
		wrapper.get<HTMLButtonElement>('[data-gallery-item="textInflate"]').element.click();
		expect(wrapper.emitted('update')).toBeUndefined();
	});

	it('has no warp gallery for an element without text properties', () => {
		const wrapper = mount(TextPanel, {
			props: { element: { type: 'image', id: 'i1' } as PptxElement },
		});
		expect(wrapper.find('pptx-ui-ribbon-gallery').exists()).toBeFalsy();
	});
});

describe('textPanel font family', () => {
	it('applies a picked font to every run, not only the body', async () => {
		const sameFace = { fontFamily: 'Old Sans', eastAsiaFont: 'Old Sans' };
		const element = {
			...textEl({ ...sameFace }),
			text: 'Sample',
			textSegments: [{ text: 'Sample', style: { ...sameFace } }],
		} as PptxElement;
		const wrapper = mount(TextPanel, { props: { element } });

		const font = wrapper
			.findAll('pptx-ui-select')
			.find((select) => select.find('option[value="Georgia"]').exists());
		await setControlValue(font!, 'Georgia');

		const patch = wrapper.emitted('update')?.at(-1)?.[0] as {
			textSegments?: Array<{ style: TextStyle }>;
		};
		wrapper.unmount();
		expect(patch.textSegments?.[0]?.style).toMatchObject({
			fontFamily: 'Georgia',
			eastAsiaFont: 'Georgia',
		});
	});
});

describe('textPanel text effects and 3-D text', () => {
	/** The checkbox a caption names: each toggle sits in a label with its caption. */
	function toggleNamed(wrapper: ReturnType<typeof mount>, caption: string) {
		const label = wrapper.findAll('label').find((node) => node.text().trim() === caption);
		if (!label) {
			throw new Error(`the text panel has no "${caption}" toggle`);
		}
		return label.get('pptx-ui-checkbox');
	}

	it('shows Text Effects with its shadow, glow and reflection toggles, then 3-D text', () => {
		const wrapper = mount(TextPanel, { props: { element: textEl({}) } });
		const text = wrapper.text();
		expect(text).toContain('Text Effects');
		expect(text).toContain('3D Text');
		// Order as in the other bindings: Text Warp, then the effects, then 3-D text.
		expect(text.indexOf('Text Warp')).toBeLessThan(text.indexOf('Text Effects'));
		expect(text.indexOf('Text Effects')).toBeLessThan(text.indexOf('3D Text'));
		for (const caption of ['Text Shadow', 'Text Glow', 'Text Reflection', '3D Extrusion']) {
			expect(toggleNamed(wrapper, caption).attributes('aria-checked')).not.toBe('true');
		}
	});

	it('turns glow on into the full textStyle patch, and off again', async () => {
		const on = mount(TextPanel, { props: { element: textEl({ bold: true }) } });
		await setControlValue(toggleNamed(on, 'Text Glow'), true);
		expect(on.emitted('update')?.[0]?.[0]).toStrictEqual({
			textStyle: { bold: true, textGlowColor: '#ffff00', textGlowRadius: 6, textGlowOpacity: 0.6 },
		});

		const glowing = { textGlowColor: '#ff0000', textGlowRadius: 8, textGlowOpacity: 0.4 };
		const off = mount(TextPanel, { props: { element: textEl({ ...glowing }) } });
		await setControlValue(toggleNamed(off, 'Text Glow'), false);
		expect(off.emitted('update')?.[0]?.[0]).toStrictEqual({
			textStyle: {
				textGlowColor: undefined,
				textGlowRadius: undefined,
				textGlowOpacity: undefined,
			},
		});
	});

	it('glows every run, not only the body, so the slide shows it', async () => {
		const element = {
			...textEl({ fontSize: 20 }),
			text: 'Hello world',
			textSegments: [
				{ text: 'Hello ', style: { fontSize: 20 } },
				{ text: 'world', style: { fontSize: 20, bold: true } },
			],
		} as PptxElement;
		const wrapper = mount(TextPanel, { props: { element } });
		await setControlValue(toggleNamed(wrapper, 'Text Glow'), true);
		const patch = wrapper.emitted('update')?.[0]?.[0] as Partial<PptxElement> & {
			textSegments?: Array<{ text: string; style?: TextStyle }>;
		};
		const glow = { textGlowColor: '#ffff00', textGlowRadius: 6, textGlowOpacity: 0.6 };
		expect(patch.textStyle).toMatchObject({ fontSize: 20, ...glow });
		expect(patch.textSegments?.map((segment) => segment.style)).toStrictEqual([
			{ fontSize: 20, ...glow },
			{ fontSize: 20, bold: true, ...glow },
		]);
	});

	it('keeps 3-D text on the body: it is an a:bodyPr property, not a run one', async () => {
		const element = {
			...textEl({}),
			text: 'Hi',
			textSegments: [{ text: 'Hi', style: { fontSize: 20 } }],
		} as PptxElement;
		const wrapper = mount(TextPanel, { props: { element } });
		await setControlValue(toggleNamed(wrapper, '3D Extrusion'), true);
		const patch = wrapper.emitted('update')?.[0]?.[0] as Partial<PptxElement>;
		expect(patch).toStrictEqual({
			textStyle: { text3d: { extrusionHeight: 76200 }, flatText: undefined },
		});
	});

	it('edits the size of a glow, clamped to the offered range', async () => {
		const wrapper = mount(TextPanel, {
			props: { element: textEl({ textGlowColor: '#ff0000', textGlowRadius: 8 }) },
		});
		const size = wrapper
			.findAll('input[type="number"]')
			.find((input) => input.attributes('aria-label') === 'Size');
		expect(size?.element).toBeDefined();
		(size!.element as HTMLInputElement).value = '400';
		await size!.trigger('input');
		expect(wrapper.emitted('update')?.at(-1)?.[0]).toStrictEqual({
			textStyle: { textGlowColor: '#ff0000', textGlowRadius: 50 },
		});
	});

	it('turns shadow and reflection on with the same defaults as the other bindings', async () => {
		const wrapper = mount(TextPanel, { props: { element: textEl({}) } });
		await setControlValue(toggleNamed(wrapper, 'Text Shadow'), true);
		await setControlValue(toggleNamed(wrapper, 'Text Reflection'), true);
		expect(wrapper.emitted('update')?.map(([patch]) => patch)).toStrictEqual([
			{
				textStyle: {
					textShadowColor: '#000000',
					textShadowBlur: 4,
					textShadowOffsetX: 2,
					textShadowOffsetY: 2,
					textShadowOpacity: 0.5,
				},
			},
			{
				textStyle: {
					textReflection: true,
					textReflectionBlur: 1,
					textReflectionStartOpacity: 0.5,
					textReflectionEndOpacity: 0,
					textReflectionOffset: 3,
				},
			},
		]);
	});

	it('seeds a visible depth when 3-D extrusion is switched on', async () => {
		const wrapper = mount(TextPanel, { props: { element: textEl({ bold: true }) } });
		await setControlValue(toggleNamed(wrapper, '3D Extrusion'), true);
		expect(wrapper.emitted('update')?.[0]?.[0]).toStrictEqual({
			textStyle: { bold: true, text3d: { extrusionHeight: 6 * 12700 }, flatText: undefined },
		});
	});

	it('switches placeholder 3-D text off with a:flatTx, so the layout extrusion cannot return', async () => {
		const element = {
			...textEl({ text3d: { extrusionHeight: 76200 } }),
			placeholderType: 'title',
		} as PptxElement;
		const wrapper = mount(TextPanel, { props: { element } });
		await setControlValue(toggleNamed(wrapper, '3D Extrusion'), false);
		expect(wrapper.emitted('update')?.[0]?.[0]).toStrictEqual({
			textStyle: { text3d: undefined, flatText: true },
		});
	});

	it('keeps the effect toggles inert on a read-only deck', () => {
		const wrapper = mount(TextPanel, { props: { element: textEl({}), canEdit: false } });
		expect(toggleNamed(wrapper, 'Text Glow').attributes('aria-disabled')).toBe('true');
	});
});
