import { mount } from '@vue/test-utils';
import { SHAPE_QUICK_STYLES } from 'ooxml-ui/pptx';
import type { PptxUiRibbonGalleryElement } from 'ooxml-ui/pptx';
import { describe, expect, it } from 'vitest';

import QuickStylesGallery from './QuickStylesGallery.vue';

/**
 * QuickStylesGallery hosts the shared `pptx-ui-ribbon-gallery` panel over
 * `buildQuickStylesGallery`, so these read the tiles the element draws in its
 * light DOM rather than any markup of this binding.
 */
function mountGallery(disabled = false) {
	return mount(QuickStylesGallery, { props: { disabled } });
}

type Wrapper = ReturnType<typeof mountGallery>;

function tiles(wrapper: Wrapper) {
	return wrapper.findAll<HTMLButtonElement>('pptx-ui-ribbon-gallery [data-gallery-item]');
}

describe('quickStylesGallery', () => {
	it('draws every preset as a tile of the shared compact panel, under its caption', () => {
		const wrapper = mountGallery();
		expect(wrapper.get('.pptx-vue-quickstyles-label').text()).toBe('Quick Styles');
		const gallery = wrapper.get('pptx-ui-ribbon-gallery');
		expect(gallery.attributes('mode')).toBe('panel');
		expect(gallery.attributes('compact')).toBeDefined();
		expect((gallery.element as PptxUiRibbonGalleryElement).descriptor?.id).toBe('quickStyles');
		expect(wrapper.get('.gallery-panel').attributes()).toMatchObject({
			role: 'group',
			'aria-label': 'Quick Styles',
		});
		const names = SHAPE_QUICK_STYLES.map((preset) => preset.name);
		expect(names).toHaveLength(18);
		expect(tiles(wrapper).map((tile) => tile.attributes('aria-label'))).toStrictEqual(names);
		expect(tiles(wrapper).map((tile) => tile.attributes('title'))).toStrictEqual(names);
		expect(tiles(wrapper)[0]!.find('svg').exists()).toBeTruthy();
	});

	it("emits the picked preset's shape style, exactly as the catalogue holds it", () => {
		const wrapper = mountGallery();
		for (const tile of tiles(wrapper)) {
			tile.element.click();
		}
		expect(wrapper.emitted('select')).toStrictEqual(
			SHAPE_QUICK_STYLES.map((preset) => [preset.style]),
		);
	});

	it('shows the tiles but takes no pick when read-only', () => {
		const wrapper = mountGallery(true);
		expect(tiles(wrapper)).toHaveLength(SHAPE_QUICK_STYLES.length);
		expect(tiles(wrapper).every((tile) => tile.element.disabled)).toBeTruthy();
		tiles(wrapper)[0]!.element.click();
		expect(wrapper.emitted('select')).toBeUndefined();
	});
});
