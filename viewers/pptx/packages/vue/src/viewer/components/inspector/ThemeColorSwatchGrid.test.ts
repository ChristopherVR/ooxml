/**
 * W3-G2: the theme colour picker grid renders the deck's REAL theme palette
 * (via `ooxml-ui/pptx`'s `buildThemeColorSwatchGrid`) instead of a
 * hard-coded Office palette, on the shared `pptx-ui-color-grid`, highlights
 * the swatch matching the element's current ref/hex, and commits both the
 * resolved hex and the ref on click.
 */
import { mount } from '@vue/test-utils';
import { describe, expect, it } from 'vitest';
import { ref } from 'vue';

import { ThemeColorMapKey } from '../../composables/theme-color-map-context';
import { themeColorGrid, themeSwatch } from './test-theme-swatch';
import ThemeColorSwatchGrid from './ThemeColorSwatchGrid.vue';

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

function mountGrid(props: Record<string, unknown> = {}, theme = ref<unknown>(OFFICE_THEME)) {
	return mount(ThemeColorSwatchGrid, {
		props,
		global: { provide: { [ThemeColorMapKey as symbol]: theme } },
	});
}

const checkedTitles = (wrapper: ReturnType<typeof mountGrid>) =>
	[...themeColorGrid(wrapper)!.querySelectorAll('button[aria-checked="true"]')].map((el) =>
		el.getAttribute('title'),
	);

describe('themeColorSwatchGrid', () => {
	it('renders nothing when no deck theme is loaded', () => {
		expect(mountGrid({}, ref(undefined)).find('pptx-ui-color-grid').exists()).toBeFalsy();
		expect(mount(ThemeColorSwatchGrid).find('pptx-ui-color-grid').exists()).toBeFalsy();
	});

	it('renders a base + variant swatch button per theme column', () => {
		const wrapper = mountGrid();
		const shadow = themeColorGrid(wrapper)!;
		// 10 columns x (1 base + 5 variants) = 60 swatch buttons, and no standard row.
		expect(shadow.querySelectorAll('button.swatch[data-source="theme"]')).toHaveLength(60);
		expect(shadow.querySelectorAll('button.swatch')).toHaveLength(60);
		// Only the Theme Colors heading: the inspector keeps its own standard and recent rows.
		const headings = [...shadow.querySelectorAll('.heading')].map((el) => el.textContent);
		expect(headings).toStrictEqual(['Theme Colors']);
		expect(wrapper.get('pptx-ui-color-grid').attributes('aria-label')).toBe('Theme Colors');
	});

	it('highlights the swatch matching the element ref', () => {
		const wrapper = mountGrid({ selectedRef: { scheme: 'accent1', lumMod: 0.2, lumOff: 0.8 } });
		expect(checkedTitles(wrapper)).toStrictEqual(['Accent 1, Lighter 80%']);
	});

	it('moves the highlight when the element ref changes', async () => {
		const wrapper = mountGrid({ selectedRef: { scheme: 'accent1' } });
		expect(checkedTitles(wrapper)).toStrictEqual(['Accent 1']);
		await wrapper.setProps({ selectedRef: { scheme: 'accent2' } });
		expect(checkedTitles(wrapper)).toStrictEqual(['Accent 2']);
	});

	it('highlights by colour when no ref is stored', () => {
		const wrapper = mountGrid({ selectedHex: '#ED7D31' });
		expect(checkedTitles(wrapper)).toStrictEqual(['Accent 2']);
	});

	it('disables every swatch when the field is disabled', () => {
		const wrapper = mountGrid({ disabled: true });
		const swatches = [...themeColorGrid(wrapper)!.querySelectorAll('button.swatch')];
		expect(swatches.every((button) => (button as HTMLButtonElement).disabled)).toBeTruthy();
		(swatches[0] as HTMLButtonElement).click();
		expect(wrapper.emitted('pick')).toBeUndefined();
	});

	it('commits both the hex and the ref on click', async () => {
		const wrapper = mountGrid();
		await themeSwatch(wrapper, 'Accent 1').trigger('click');
		expect(wrapper.emitted('pick')).toStrictEqual([
			[{ hex: '#4472c4', ref: { scheme: 'accent1' } }],
		]);
	});
});
