import { mount } from '@vue/test-utils';
import type { ShapeFillStopView } from 'ooxml-ui/pptx';
import { describe, expect, it } from 'vitest';

import FillGradientControls from './FillGradientControls.vue';
import { setControlValue } from './test-control-value';

const stops: ShapeFillStopView[] = [
	{ index: 0, color: '#000000', position: 0, opacity: 100, removable: true },
	{ index: 1, color: '#888888', position: 50, opacity: 40, removable: true },
	{ index: 2, color: '#ffffff', position: 100, opacity: 100, removable: true },
];

const actions = (wrapper: ReturnType<typeof mount>) =>
	(wrapper.emitted('action') ?? []).map((args) => args[0]);

describe('fillGradientControls', () => {
	it('draws one row per shared stop view', () => {
		const wrapper = mount(FillGradientControls, { props: { stops } });
		expect(wrapper.findAll('[data-testid="fill-gradient-stop"]')).toHaveLength(3);
		const numbers = wrapper.findAll('input[type="number"]');
		expect((numbers[3]!.element as HTMLInputElement).value).toBe('40');
	});

	it('emits the shared stop actions', async () => {
		const wrapper = mount(FillGradientControls, { props: { stops } });
		await setControlValue(wrapper.findAll('input[type="number"]')[2]!, '30');
		await wrapper.findAll('.pptx-vue-gradient-remove')[1]!.trigger('click');
		await wrapper.find('.pptx-vue-gradient-add').trigger('click');
		expect(actions(wrapper)).toStrictEqual([
			{ type: 'stopPosition', index: 1, position: 30 },
			{ type: 'removeStop', index: 1 },
			{ type: 'addStop' },
		]);
	});

	it('disables remove on the last two stops', () => {
		const wrapper = mount(FillGradientControls, {
			props: { stops: stops.slice(0, 2).map((s) => ({ ...s, removable: false })) },
		});
		expect(wrapper.find('.pptx-vue-gradient-remove').attributes('disabled')).toBeDefined();
	});
});
