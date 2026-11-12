import { mount } from '@vue/test-utils';
import { createRibbonDisplayController } from 'ooxml-ui/pptx';
import { describe, expect, it, vi } from 'vitest';

import RibbonDisplayMenu from './RibbonDisplayMenu.vue';

/**
 * RibbonDisplayMenu: the Ribbon Display Options menu. The three modes come in PowerPoint's
 * order, the current one is checked, and a pick goes through the controller (which is the
 * only place that fires the change notification).
 */
describe('ribbonDisplayMenu', () => {
	function mountMenu(initial: 'always' | 'tabs' | 'fullScreen' = 'always') {
		const onChange = vi.fn();
		const controller = createRibbonDisplayController({ initial, onChange });
		controller.setMenuOpen(true);
		const wrapper = mount(RibbonDisplayMenu, {
			props: { state: controller.getSnapshot(), controller },
		});
		return { wrapper, controller, onChange };
	}

	it('lists Full-screen mode, Show tabs only, Always show Ribbon in that order', () => {
		const { wrapper } = mountMenu();
		const modes = wrapper
			.findAll('[role="menuitemradio"]')
			.map((item) => item.attributes('data-ribbon-display'));
		expect(modes).toEqual(['fullScreen', 'tabs', 'always']);
	});

	it('checks only the current mode', () => {
		const { wrapper } = mountMenu('tabs');
		const checked = wrapper
			.findAll('[role="menuitemradio"]')
			.map((item) => [item.attributes('data-ribbon-display'), item.attributes('aria-checked')]);
		expect(checked).toEqual([
			['fullScreen', 'false'],
			['tabs', 'true'],
			['always', 'false'],
		]);
	});

	it('choosing a mode fires the change notification with that mode', async () => {
		const { wrapper, onChange } = mountMenu('always');
		await wrapper.find('[data-ribbon-display="tabs"]').trigger('click');
		expect(onChange).toHaveBeenCalledTimes(1);
		expect(onChange).toHaveBeenCalledWith('tabs');
	});

	it('draws nothing while the menu is closed', () => {
		const onChange = vi.fn();
		const controller = createRibbonDisplayController({ initial: 'always', onChange });
		const wrapper = mount(RibbonDisplayMenu, {
			props: { state: controller.getSnapshot(), controller },
		});
		expect(wrapper.find('[role="menu"]').exists()).toBe(false);
	});

	it('a host change of the mode does not fire the change notification', () => {
		const onChange = vi.fn();
		const controller = createRibbonDisplayController({ initial: 'always', onChange });
		controller.setMode('fullScreen');
		expect(controller.getSnapshot().mode).toBe('fullScreen');
		expect(onChange).not.toHaveBeenCalled();
	});
});
