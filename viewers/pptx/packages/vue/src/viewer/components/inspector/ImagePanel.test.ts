import { mount } from '@vue/test-utils';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import ImagePanel from './ImagePanel.vue';
import { setControlValue } from './test-control-value';

function imageEl(overrides: Partial<PptxElement> = {}): PptxElement {
	return {
		type: 'image',
		id: 'img 1',
		x: 0,
		y: 0,
		width: 100,
		height: 100,
		...overrides,
	} as PptxElement;
}

function shapeEl(): PptxElement {
	return {
		type: 'shape',
		id: 'sh 1',
		x: 0,
		y: 0,
		width: 100,
		height: 100,
	} as PptxElement;
}

describe('imagePanel', () => {
	it('draws the artistic effect gallery over the picture and emits the picked effect', async () => {
		const wrapper = mount(ImagePanel, {
			props: {
				element: imageEl({
					imageData: 'data:image/png;base64,AA',
					imageEffects: { artisticEffect: 'blur' },
				} as Partial<PptxElement>),
			},
		});
		const blur = wrapper.get('[data-testid="picture-artistic"] [data-value="blur"]');
		expect(blur.attributes('aria-pressed')).toBe('true');
		expect((blur.get('img').element as HTMLImageElement).style.filter).toBe('blur(4px)');
		await wrapper.get('[data-testid="picture-artistic"] [data-value="sepia"]').trigger('click');
		expect(wrapper.emitted('update')?.at(-1)?.[0]).toStrictEqual({
			imageEffects: { artisticEffect: 'sepia' },
		});
	});

	it('shows a muted note for non-image elements', () => {
		const wrapper = mount(ImagePanel, { props: { element: shapeEl() } });
		expect(wrapper.find('.pptx-vue-image-panel__note').exists()).toBeTruthy();
		expect(wrapper.find('input').exists()).toBeFalsy();
	});

	it('renders alt text and emits a shallow patch on input', async () => {
		const wrapper = mount(ImagePanel, {
			props: { element: imageEl({ altText: 'hello' } as Partial<PptxElement>) },
		});
		const input = wrapper.get('[data-testid="element-alt-text"]');
		expect((input.element as HTMLTextAreaElement).value).toBe('hello');

		await setControlValue(input, 'logo');
		const events = wrapper.emitted('update');
		expect(events).toBeTruthy();
		expect(events?.at(-1)?.[0]).toStrictEqual({ altText: 'logo' });
	});

	it('relays the full merged imageEffects sub-object on brightness change', async () => {
		const wrapper = mount(ImagePanel, {
			props: {
				element: imageEl({ imageEffects: { contrast: 25 } } as Partial<PptxElement>),
			},
		});
		// The adjustments are the shared picture descriptors; ImagePanel relays their patch.
		const brightness = wrapper.get('[data-testid="picture-brightness"]');
		await setControlValue(brightness, '40');

		const events = wrapper.emitted('update');
		expect(events?.at(-1)?.[0]).toStrictEqual({ imageEffects: { contrast: 25, brightness: 40 } });
	});

	it('toggles grayscale into the merged imageEffects sub-object', async () => {
		const wrapper = mount(ImagePanel, {
			props: {
				element: imageEl({ imageEffects: { contrast: 10 } } as Partial<PptxElement>),
			},
		});
		const grayscale = wrapper.get('[data-testid="picture-grayscale"]');
		await setControlValue(grayscale, true);

		const events = wrapper.emitted('update');
		expect(events?.at(-1)?.[0]).toStrictEqual({ imageEffects: { contrast: 10, grayscale: true } });
	});

	it('draws the shared colour effects and applies a duotone preset', async () => {
		const wrapper = mount(ImagePanel, {
			props: { element: imageEl({ imageEffects: { contrast: 10 } } as Partial<PptxElement>) },
		});
		await setControlValue(wrapper.get('[data-testid="picture-colorChange-toggle"]'), true);
		expect(wrapper.emitted('update')?.at(-1)?.[0]).toStrictEqual({
			imageEffects: {
				contrast: 10,
				clrChange: { clrFrom: '#ffffff', clrTo: '#000000', clrToTransparent: false },
			},
		});
		await wrapper.get('[data-testid="picture-duotone-preset"]').trigger('click');
		const last = wrapper.emitted('update')?.at(-1)?.[0] as {
			imageEffects?: { duotone?: unknown };
		};
		expect(last.imageEffects?.duotone).toBeTruthy();
	});

	it('disables Reset Picture until an effect override exists and clears effects when clicked', async () => {
		const clean = mount(ImagePanel, { props: { element: imageEl() } });
		expect(
			(clean.get('.pptx-vue-image-panel__reset-picture').element as HTMLButtonElement).disabled,
		).toBeTruthy();

		const dirty = mount(ImagePanel, {
			props: { element: imageEl({ imageEffects: { brightness: 5 } } as Partial<PptxElement>) },
		});
		const reset = dirty.get('.pptx-vue-image-panel__reset-picture');
		await reset.trigger('click');
		const events = dirty.emitted('update');
		expect(events?.at(-1)?.[0]).toStrictEqual({ imageEffects: undefined, cropShape: 'none' });
	});
});
