import { createRibbonDisplayController } from 'ooxml-ui/pptx';
import { mount, unmount, flushSync } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import RibbonTabBar from './RibbonTabBar.svelte';

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function render(props: Record<string, unknown>): HTMLElement {
	const target = document.createElement('div');
	document.body.append(target);
	const instance = mount(RibbonTabBar, { target, props });
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return target;
}

describe('ribbonTabBar Ribbon Display Options', () => {
	it('renders the menu button after the right-hand actions and toggles the menu', () => {
		const controller = createRibbonDisplayController({ initial: 'always' });
		const target = render({ active: 'home', onselect: vi.fn(), ribbonDisplay: controller });
		flushSync();

		const button = target.querySelector<HTMLButtonElement>(
			'[data-pptx-chrome="ribbon-display-options"]',
		);
		expect(button?.getAttribute('aria-haspopup')).toBe('menu');
		expect(button?.getAttribute('aria-expanded')).toBe('false');

		button?.click();
		flushSync();

		expect(controller.getSnapshot().menuOpen).toBe(true);
		expect(button?.getAttribute('aria-expanded')).toBe('true');
	});

	it('omits the menu button when no controller is given', () => {
		const target = render({ active: 'home', onselect: vi.fn() });
		flushSync();

		expect(target.querySelector('[data-pptx-chrome="ribbon-display-options"]')).toBeNull();
	});

	it('a double click on a tab other than File collapses the ribbon, and notifies the change', () => {
		const onChange = vi.fn();
		const controller = createRibbonDisplayController({ initial: 'always', onChange });
		const target = render({ active: 'home', onselect: vi.fn(), ribbonDisplay: controller });
		flushSync();

		const home = [...target.querySelectorAll<HTMLButtonElement>('[role="tab"]')].find(
			(tab) => tab.textContent?.trim() === 'Home',
		);
		home?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));

		expect(controller.getSnapshot().effectiveMode).toBe('tabs');
		expect(onChange).toHaveBeenCalledTimes(1);
		expect(onChange).toHaveBeenCalledWith('tabs');
	});
});
