import { mount } from '@vue/test-utils';
import { describe, expect, it } from 'vitest';

import RecentColorsRow from './RecentColorsRow.vue';

/** The shared grid's recent swatches, inside its shadow root. */
function swatches(wrapper: ReturnType<typeof mount>): HTMLButtonElement[] {
	const grid = wrapper.find('[data-testid="pptx-color-recent"]').element;
	return [...(grid.shadowRoot?.querySelectorAll<HTMLButtonElement>('button') ?? [])];
}

describe('recentColorsRow', () => {
	it('renders nothing when there are no recent colours', () => {
		const wrapper = mount(RecentColorsRow, { props: { colors: [] } });
		expect(wrapper.find('[data-testid="pptx-color-recent"]').exists()).toBeFalsy();
	});

	it('is the shared grid with one swatch per colour, in order', () => {
		const wrapper = mount(RecentColorsRow, {
			props: { colors: ['#112233', '#445566'] },
			attachTo: document.body,
		});
		expect(wrapper.find('[data-testid="pptx-color-recent"]').element.localName).toBe(
			'pptx-ui-color-grid',
		);
		expect(swatches(wrapper).map((swatch) => swatch.dataset.color)).toStrictEqual([
			'#112233',
			'#445566',
		]);
		wrapper.unmount();
	});

	it('emits pick with the clicked colour', () => {
		const wrapper = mount(RecentColorsRow, {
			props: { colors: ['#112233', '#445566'] },
			attachTo: document.body,
		});
		swatches(wrapper)[1]!.click();
		expect(wrapper.emitted('pick')).toStrictEqual([['#445566']]);
		wrapper.unmount();
	});

	it('labels the row and each swatch per the row contract', () => {
		const wrapper = mount(RecentColorsRow, {
			props: { colors: ['#112233'] },
			attachTo: document.body,
		});
		const grid = wrapper.get('[data-testid="pptx-color-recent"]');
		expect(grid.attributes('aria-label')).toBe('Recent Colors');
		expect(swatches(wrapper)[0]!.getAttribute('aria-label')).toBe('Recent #112233');
		wrapper.unmount();
	});

	it('disables every swatch when the picker is disabled', () => {
		const wrapper = mount(RecentColorsRow, {
			props: { colors: ['#112233', '#445566'], disabled: true },
			attachTo: document.body,
		});
		expect(swatches(wrapper).every((swatch) => swatch.disabled)).toBeTruthy();
		wrapper.unmount();
	});
});
