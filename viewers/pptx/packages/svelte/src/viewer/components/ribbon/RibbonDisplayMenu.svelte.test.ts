import { createRibbonDisplayController } from 'ooxml-ui/pptx';
import { mount, unmount, flushSync } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import RibbonDisplayMenu from './RibbonDisplayMenu.svelte';
import RibbonDisplayReveal from './RibbonDisplayReveal.svelte';

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function render(component: typeof RibbonDisplayMenu | typeof RibbonDisplayReveal, props: object) {
	const target = document.createElement('div');
	document.body.append(target);
	const instance = mount(component, { target, props });
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return target;
}

describe('ribbon display options menu', () => {
	it('lists the three modes in PowerPoint order with the current one checked', () => {
		const controller = createRibbonDisplayController({ initial: 'always' });
		controller.toggleMenu();
		const target = render(RibbonDisplayMenu, { ribbonDisplay: controller });
		flushSync();

		const items = [...target.querySelectorAll<HTMLButtonElement>('[role="menuitemradio"]')];
		expect(items.map((item) => item.dataset.ribbonDisplay)).toEqual([
			'fullScreen',
			'tabs',
			'always',
		]);
		expect(items.map((item) => item.getAttribute('aria-checked'))).toEqual([
			'false',
			'false',
			'true',
		]);
	});

	it('choosing Show tabs only fires the change notification once and closes the menu', () => {
		const onChange = vi.fn();
		const controller = createRibbonDisplayController({ initial: 'always', onChange });
		controller.toggleMenu();
		const target = render(RibbonDisplayMenu, { ribbonDisplay: controller });
		flushSync();

		target.querySelector<HTMLButtonElement>('[data-ribbon-display="tabs"]')?.click();
		flushSync();

		expect(onChange).toHaveBeenCalledTimes(1);
		expect(onChange).toHaveBeenCalledWith('tabs');
		expect(target.querySelector('[role="menu"]')).toBeNull();
	});

	it('draws nothing while the menu is closed', () => {
		const controller = createRibbonDisplayController({ initial: 'always' });
		const target = render(RibbonDisplayMenu, { ribbonDisplay: controller });
		flushSync();

		expect(target.querySelector('[role="menu"]')).toBeNull();
	});
});

describe('ribbon reveal strip', () => {
	it('shows only in full-screen mode, and pressing it peeks the ribbon without notifying', () => {
		const onChange = vi.fn();
		const controller = createRibbonDisplayController({ initial: 'fullScreen', onChange });
		const target = render(RibbonDisplayReveal, { ribbonDisplay: controller });
		flushSync();

		const strip = target.querySelector<HTMLButtonElement>('[data-pptx-chrome="ribbon-reveal"]');
		expect(strip).not.toBeNull();
		expect(strip?.textContent?.trim()).toBe('···');

		strip?.click();
		flushSync();

		expect(controller.getSnapshot().peek).toBe(true);
		// The strip keeps its place under the peeked ribbon, so the slide does not jump.
		expect(target.querySelector('[data-pptx-chrome="ribbon-reveal"]')).not.toBeNull();
		expect(onChange).not.toHaveBeenCalled();

		controller.setMode('tabs');
		flushSync();
		expect(target.querySelector('[data-pptx-chrome="ribbon-reveal"]')).toBeNull();
	});

	it('is not drawn in the docked modes', () => {
		const controller = createRibbonDisplayController({ initial: 'always' });
		const target = render(RibbonDisplayReveal, { ribbonDisplay: controller });
		flushSync();

		expect(target.querySelector('[data-pptx-chrome="ribbon-reveal"]')).toBeNull();
	});
});
