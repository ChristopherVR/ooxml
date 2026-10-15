import { afterEach, describe, expect, it } from 'vitest';
import { syncScrollMore } from './scroll-more';

afterEach(() => document.body.replaceChildren());

/** A panel whose content overflows its width, so the chevron is pinned. */
function overflowingPanel(): HTMLElement {
	const panel = document.createElement('div');
	document.body.append(panel);
	Object.defineProperty(panel, 'clientWidth', { get: () => 100 });
	Object.defineProperty(panel, 'scrollWidth', { get: () => 400 });
	return panel;
}

const chevron = (panel: HTMLElement) =>
	panel.querySelector<HTMLButtonElement>(':scope > .ribbon-scroll-more');

describe('syncScrollMore label', () => {
	it('defaults to the English "More"', () => {
		const panel = overflowingPanel();
		syncScrollMore(panel);
		expect(chevron(panel)?.getAttribute('aria-label')).toBe('More');
		expect(chevron(panel)?.title).toBe('More');
	});

	it('uses a string label as given', () => {
		const panel = overflowingPanel();
		syncScrollMore(panel, 'Weitere');
		expect(chevron(panel)?.getAttribute('aria-label')).toBe('Weitere');
	});

	it('calls a function label with the panel and re-reads it on every sync', () => {
		const panel = overflowingPanel();
		let locale: Record<string, string> = { en: 'More', de: 'Mehr' };
		let current = 'en';
		const seen: HTMLElement[] = [];
		const label = (target: HTMLElement) => {
			seen.push(target);
			return locale[current] ?? 'More';
		};

		syncScrollMore(panel, label);
		expect(chevron(panel)?.getAttribute('aria-label')).toBe('More');
		expect(seen).toEqual([panel]);

		current = 'de';
		syncScrollMore(panel, label);
		expect(chevron(panel)?.getAttribute('aria-label')).toBe('Mehr');
		expect(chevron(panel)?.title).toBe('Mehr');
		expect(panel.querySelectorAll('.ribbon-scroll-more')).toHaveLength(1);

		locale = { de: 'Weitere' };
		syncScrollMore(panel, label);
		expect(chevron(panel)?.getAttribute('aria-label')).toBe('Weitere');
	});
});
