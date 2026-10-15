// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { registerViewerControls } from './office-ui';
import { createRibbon } from './ribbon';
import { localizeRibbon } from './ribbon-localize';
import { wireRibbonTabCue } from './ribbon-tab-cue';

afterEach(() => {
	vi.restoreAllMocks();
	document.body.replaceChildren();
});

/** A tab strip of `scrollWidth` px in a `clientWidth` px viewport, scrolled to `scrollLeft`. */
function strip(ribbon: HTMLElement, sizes: { scrollWidth: number; clientWidth: number }) {
	const list = ribbon.shadowRoot!.querySelector<HTMLElement>('[role="tablist"]')!;
	Object.defineProperty(list, 'scrollWidth', { value: sizes.scrollWidth, configurable: true });
	Object.defineProperty(list, 'clientWidth', { value: sizes.clientWidth, configurable: true });
	return list;
}

async function mount() {
	registerViewerControls();
	vi.spyOn(window, 'requestAnimationFrame').mockImplementation((callback) => {
		callback(0);
		return 0;
	});
	const ribbon = document.body.appendChild(createRibbon(document));
	await (ribbon as HTMLElement & { updateComplete: Promise<unknown> }).updateComplete;
	return ribbon;
}

describe('ribbon tab cue', () => {
	it('shows nothing while every tab fits', async () => {
		const ribbon = await mount();
		const list = strip(ribbon, { scrollWidth: 300, clientWidth: 300 });
		const dispose = wireRibbonTabCue(ribbon);
		list.dispatchEvent(new Event('scroll'));
		expect(ribbon.hasAttribute('data-tabs-end')).toBe(false);
		expect(ribbon.querySelector<HTMLElement>('.tab-scroll-more')!.hidden).toBe(true);
		dispose();
	});

	it('fades the edges that have more tabs and pages through them', async () => {
		const ribbon = await mount();
		const list = strip(ribbon, { scrollWidth: 600, clientWidth: 300 });
		const scrollTo = vi.fn();
		list.scrollTo = scrollTo as unknown as typeof list.scrollTo;
		const dispose = wireRibbonTabCue(ribbon);
		list.dispatchEvent(new Event('scroll'));
		const button = ribbon.querySelector<HTMLButtonElement>('.tab-scroll-more')!;
		expect(button.hidden).toBe(false);
		expect(button.slot).toBe('end');
		expect(button.getAttribute('aria-label')).toBe('More tabs');
		expect(ribbon.hasAttribute('data-tabs-end')).toBe(true);
		expect(ribbon.hasAttribute('data-tabs-start')).toBe(false);
		button.click();
		expect(scrollTo).toHaveBeenLastCalledWith({ left: 240, behavior: 'smooth' });
		Object.defineProperty(list, 'scrollLeft', { value: 300, configurable: true });
		list.dispatchEvent(new Event('scroll'));
		expect(ribbon.hasAttribute('data-tabs-start')).toBe(true);
		expect(ribbon.hasAttribute('data-tabs-end')).toBe(false);
		expect(button.textContent).toBe('‹');
		button.click();
		expect(scrollTo).toHaveBeenLastCalledWith({ left: 0, behavior: 'smooth' });
		dispose();
		expect(ribbon.querySelector('.tab-scroll-more')).toBeNull();
	});

	it('names the button in the ribbon locale', async () => {
		const ribbon = await mount();
		const list = strip(ribbon, { scrollWidth: 600, clientWidth: 300 });
		localizeRibbon(ribbon, 'fr');
		const dispose = wireRibbonTabCue(ribbon);
		list.dispatchEvent(new Event('scroll'));
		expect(ribbon.querySelector('.tab-scroll-more')!.getAttribute('aria-label')).toBe(
			'Autres onglets',
		);
		dispose();
	});
});
