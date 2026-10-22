import { mount } from '@vue/test-utils';
import type { PptxUiRibbonGalleryElement } from 'ooxml-ui/pptx';
import type { TextStyle } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import TextWarpGallery from './TextWarpGallery.vue';

/**
 * TextWarpGallery hosts the shared `pptx-ui-ribbon-gallery` panel over
 * `buildTextWarpGallery`, so these read the tiles the element draws in its
 * light DOM rather than any markup of this binding.
 */
function mountGallery(ts: TextStyle | undefined, disabled = false) {
	return mount(TextWarpGallery, { props: { ts, disabled } });
}

type Wrapper = ReturnType<typeof mountGallery>;

function tiles(wrapper: Wrapper) {
	return wrapper.findAll<HTMLButtonElement>('pptx-ui-ribbon-gallery [data-gallery-item]');
}

function pressed(wrapper: Wrapper): string[] {
	return tiles(wrapper)
		.filter((entry) => entry.attributes('aria-pressed') === 'true')
		.map((entry) => entry.attributes('data-gallery-item')!);
}

function tile(wrapper: Wrapper, preset: string): HTMLButtonElement {
	return wrapper.get<HTMLButtonElement>(`[data-gallery-item="${preset}"]`).element;
}

describe('textWarpGallery', () => {
	it('draws every preset as a tile of the shared compact panel', () => {
		const wrapper = mountGallery({ textWarpPreset: 'textArchUp' });
		const gallery = wrapper.get('pptx-ui-ribbon-gallery');
		expect(gallery.attributes('mode')).toBe('panel');
		expect(gallery.attributes('compact')).toBeDefined();
		expect((gallery.element as PptxUiRibbonGalleryElement).descriptor?.id).toBe('textWarp');
		expect(tiles(wrapper)).toHaveLength(39);
		expect(wrapper.get('.gallery-panel').attributes()).toMatchObject({
			role: 'group',
			'aria-label': 'Text Warp',
		});
		const archUp = wrapper.get('[data-gallery-item="textArchUp"]');
		expect(archUp.attributes('aria-label')).toBe('Arch Up');
		expect(archUp.attributes('title')).toBe('Arch Up');
	});

	it('marks the current preset, and No Transform when the text has none', () => {
		expect(pressed(mountGallery({ textWarpPreset: 'textArchUp' }))).toStrictEqual(['textArchUp']);
		expect(pressed(mountGallery({}))).toStrictEqual(['textNoShape']);
		expect(pressed(mountGallery(undefined))).toStrictEqual(['textNoShape']);
	});

	it('keeps the gallery in the DOM while collapsed and only shows it when expanded', async () => {
		const wrapper = mountGallery({ textWarpPreset: 'textInflate' });
		const toggle = wrapper.get('.pptx-vue-textwarp-toggle');
		const container = wrapper.get<HTMLElement>('.pptx-vue-textwarp-gallery').element;
		expect(toggle.attributes('aria-expanded')).toBe('false');
		expect(container.style.display).toBe('none');
		expect(tiles(wrapper)).toHaveLength(39);
		expect(wrapper.get('.pptx-vue-textwarp-current').text()).toBe('Inflate');

		await toggle.trigger('click');
		expect(toggle.attributes('aria-expanded')).toBe('true');
		expect(container.style.display).toBe('');
		expect(wrapper.find('.pptx-vue-textwarp-current').exists()).toBeFalsy();
		expect(
			wrapper.get('pptx-ui-ribbon-gallery').element.contains(tile(wrapper, 'textInflate')),
		).toBeTruthy();
	});

	it('emits the text style change of a picked tile, clearing it for No Transform', async () => {
		const wrapper = mountGallery({ textWarpPreset: 'textArchUp' });
		await wrapper.get('.pptx-vue-textwarp-toggle').trigger('click');
		tile(wrapper, 'textWave1').click();
		tile(wrapper, 'textNoShape').click();
		expect(wrapper.emitted('select')).toStrictEqual([
			[{ textWarpPreset: 'textWave1' }],
			[{ textWarpPreset: undefined }],
		]);
	});

	it('shows the current warp but takes no pick when read-only', async () => {
		const wrapper = mountGallery({ textWarpPreset: 'textArchUp' }, true);
		await wrapper.get('.pptx-vue-textwarp-toggle').trigger('click');
		expect(tiles(wrapper).every((entry) => entry.element.disabled)).toBeTruthy();
		expect(pressed(wrapper)).toStrictEqual(['textArchUp']);
		tile(wrapper, 'textWave1').click();
		expect(wrapper.emitted('select')).toBeUndefined();
	});
});
