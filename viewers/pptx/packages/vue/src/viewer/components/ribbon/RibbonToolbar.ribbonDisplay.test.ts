import { mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import { createRibbonDisplayController } from 'ooxml-ui/pptx';
import type { RibbonDisplayController } from 'ooxml-ui/pptx';
import { describe, expect, it, vi } from 'vitest';

import { createRibbonPropsFixture } from './ribbon-props-fixture';
import type { ToolbarSection } from './ribbon-types';
import RibbonRevealStrip from './RibbonRevealStrip.vue';
import RibbonToolbar from './RibbonToolbar.vue';

/**
 * RibbonToolbar under Ribbon Display Options: the command content of "Show tabs only" lays over
 * the slide once a tab is chosen, Escape puts it away, and the menu switches the mode. The
 * controller is the source of truth, so its snapshot is handed back in after every action, as
 * the root component's store adapter does.
 */
describe('ribbonToolbar ribbon display', () => {
	function setup(
		initial: 'always' | 'tabs' | 'fullScreen' = 'always',
		toolbarSection: ToolbarSection = 'home',
	) {
		const onChange = vi.fn();
		const controller: RibbonDisplayController = createRibbonDisplayController({
			initial,
			onChange,
		});
		const wrapper = mount(RibbonToolbar, {
			props: {
				...createRibbonPropsFixture({ toolbarSection }),
				ribbonDisplay: controller.getSnapshot(),
				ribbonDisplayController: controller,
			},
		});
		const sync = async () => {
			await wrapper.setProps({ ribbonDisplay: controller.getSnapshot() });
		};
		return { wrapper, controller, onChange, sync };
	}

	/** The wrapper that shows or hides the command content (`display: contents` when docked). */
	function contentWrapper(wrapper: VueWrapper): HTMLElement {
		const content = wrapper.find('[data-pptx-chrome="ribbon-content"]').element as HTMLElement;
		return content.parentElement as HTMLElement;
	}

	it('"Show tabs only" hides the command content until a tab is chosen', async () => {
		const { wrapper, controller, sync } = setup('tabs');
		await sync();
		expect(wrapper.find('[data-pptx-chrome="ribbon"]').attributes('data-ribbon-display')).toBe(
			'tabs',
		);
		expect(contentWrapper(wrapper).style.display).toBe('none');

		const home = wrapper.findAll('[role="tab"]').find((tab) => tab.text() === 'Home');
		expect(home).toBeDefined();
		await home?.trigger('click');
		await sync();
		expect(controller.getSnapshot().peek).toBe(true);
		expect(contentWrapper(wrapper).style.display).not.toBe('none');
		expect(contentWrapper(wrapper).className).toContain('absolute');

		controller.handleKeydown({ key: 'Escape', ctrlKey: false, preventDefault() {} });
		await sync();
		expect(contentWrapper(wrapper).style.display).toBe('none');
	});

	it('choosing "Show tabs only" from the menu hides the commands and fires the change with tabs', async () => {
		const { wrapper, controller, onChange, sync } = setup('always');
		controller.setMenuOpen(true);
		await sync();
		expect(contentWrapper(wrapper).style.display).not.toBe('none');

		await wrapper.find('[data-ribbon-display="tabs"]').trigger('click');
		await sync();
		expect(onChange).toHaveBeenCalledWith('tabs');
		expect(contentWrapper(wrapper).style.display).toBe('none');
	});

	it('the File tab is docked, never hidden or overlaid, in "Show tabs only"', async () => {
		const { wrapper, sync } = setup('tabs', 'file');
		await sync();
		expect(contentWrapper(wrapper).style.display).not.toBe('none');
		expect(contentWrapper(wrapper).className).not.toContain('absolute');
	});

	it('draws the menu button only with a controller', () => {
		const { wrapper } = setup('always');
		expect(wrapper.find('[data-pptx-chrome="ribbon-display-options"]').exists()).toBe(true);
	});
});

describe('ribbonRevealStrip', () => {
	it('pressing the strip brings the ribbon back over the slide', async () => {
		const controller = createRibbonDisplayController({ initial: 'fullScreen' });
		const wrapper = mount(RibbonRevealStrip, { props: { controller } });
		expect(wrapper.attributes('data-pptx-chrome')).toBe('ribbon-reveal');
		await wrapper.trigger('click');
		expect(controller.getSnapshot().peek).toBe(true);
	});
});
