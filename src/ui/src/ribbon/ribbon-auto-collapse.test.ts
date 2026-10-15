// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { registerOfficeUi } from '../index';
import { syncScrollMore } from './scroll-more';

beforeAll(() => registerOfficeUi());
afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
	delete (window as { matchMedia?: unknown }).matchMedia;
	vi.restoreAllMocks();
});

/** A media query list the test can flip. */
function stubShortViewport(initial: boolean) {
	const listeners = new Set<() => void>();
	const query = {
		matches: initial,
		addEventListener: (_: string, fn: () => void) => listeners.add(fn),
		removeEventListener: (_: string, fn: () => void) => listeners.delete(fn),
	};
	vi.stubGlobal('matchMedia', () => query);
	Object.defineProperty(window, 'matchMedia', { value: () => query, configurable: true });
	return (matches: boolean) => {
		query.matches = matches;
		for (const fn of listeners) fn();
	};
}

function ribbon() {
	const el = document.createElement('office-ui-ribbon');
	el.toggleAttribute('collapsible', true);
	el.innerHTML =
		'<div data-ribbon-tab="home" data-label="Home"><button>Bold</button></div>' +
		'<div data-ribbon-tab="view" data-label="View"><button>Grid</button></div>';
	document.body.append(el);
	return el;
}

describe('ribbon on a short viewport', () => {
	it('starts collapsed to the tab row and opens a tab body as an overlay', async () => {
		stubShortViewport(true);
		const el = ribbon();
		await Promise.resolve();
		expect(el.hasAttribute('collapsed')).toBe(true);
		expect(el.querySelector<HTMLElement>('[data-ribbon-tab="home"]')!.hidden).toBe(true);
		(el as unknown as { tabButton(t: string): HTMLButtonElement }).tabButton('view').click();
		await Promise.resolve();
		expect(el.hasAttribute('peek')).toBe(true);
		expect(el.querySelector<HTMLElement>('[data-ribbon-tab="view"]')!.hidden).toBe(false);
	});

	it('follows the viewport until the user chooses, then the choice wins', async () => {
		const flip = stubShortViewport(false);
		const el = ribbon();
		await Promise.resolve();
		expect(el.hasAttribute('collapsed')).toBe(false);
		flip(true);
		await Promise.resolve();
		expect(el.hasAttribute('collapsed')).toBe(true);
		el.shadowRoot!.querySelector<HTMLButtonElement>('.collapse')!.click();
		await Promise.resolve();
		expect(el.hasAttribute('collapsed')).toBe(false);
		flip(false);
		flip(true);
		await Promise.resolve();
		expect(el.hasAttribute('collapsed')).toBe(false);
	});
});

describe('syncScrollMore', () => {
	function panel(overflowing: boolean) {
		const el = document.createElement('div');
		el.className = 'ribbon-panel';
		Object.defineProperty(el, 'scrollWidth', { get: () => (overflowing ? 600 : 300) });
		Object.defineProperty(el, 'clientWidth', { get: () => 300 });
		document.body.append(el);
		return el;
	}

	it('pins one chevron while the panel overflows and removes it when it fits', () => {
		const el = panel(true);
		syncScrollMore(el, 'More commands');
		syncScrollMore(el, 'More commands');
		const buttons = el.querySelectorAll('.ribbon-scroll-more');
		expect(buttons).toHaveLength(1);
		expect(buttons[0]?.getAttribute('aria-label')).toBe('More commands');
		const fits = panel(false);
		syncScrollMore(fits);
		expect(fits.querySelector('.ribbon-scroll-more')).toBeNull();
	});

	it('scrolls on by a page and wraps to the start at the end', () => {
		const el = panel(true);
		el.scrollTo = vi.fn() as unknown as typeof el.scrollTo;
		syncScrollMore(el);
		const button = el.querySelector<HTMLButtonElement>('.ribbon-scroll-more')!;
		button.click();
		expect(el.scrollTo).toHaveBeenLastCalledWith({ left: 240, behavior: 'smooth' });
		el.scrollLeft = 300;
		button.click();
		expect(el.scrollTo).toHaveBeenLastCalledWith({ left: 0, behavior: 'smooth' });
	});
});
