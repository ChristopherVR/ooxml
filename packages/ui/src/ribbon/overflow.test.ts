// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createRibbonOverflow } from './overflow.js';

const overflow = createRibbonOverflow({
	icon: (doc, name) => Object.assign(doc.createElement('i'), { textContent: name }),
	groupIcons: { Font: 'font' },
	fallbackIcon: 'more',
	ribbon: '.ribbon',
	command: 'button[data-command]',
});

/** jsdom has no layout: a panel that is `room` wide holding groups of 100 that overflow it. */
function panel(room: number) {
	const ribbon = document.createElement('div');
	ribbon.className = 'ribbon';
	const el = document.createElement('div');
	el.className = 'ribbon-panel';
	const define = (node: HTMLElement, props: Record<string, () => number>) => {
		for (const [key, get] of Object.entries(props)) Object.defineProperty(node, key, { get });
	};
	const groups = ['Font', 'Styles'].map((label) => {
		const group = document.createElement('div');
		group.className = 'ribbon-group';
		group.dataset.label = label;
		group.dataset.caption = label;
		const run = document.createElement('button');
		run.dataset.command = 'x';
		group.append(run);
		define(group, { offsetWidth: () => (group.hasAttribute('data-collapsed') ? 40 : 100) });
		el.append(group);
		return group;
	});
	define(el, {
		clientWidth: () => room,
		scrollWidth: () =>
			groups.reduce((sum, group) => sum + (group.hasAttribute('data-collapsed') ? 40 : 100), 0),
	});
	ribbon.append(el);
	document.body.append(ribbon);
	return { el, groups };
}

afterEach(() => {
	overflow.close();
	document.body.replaceChildren();
});

describe('ribbon overflow', () => {
	it('folds groups from the right until the panel fits, naming the button after the group', () => {
		const { el, groups } = panel(150);
		overflow.fitPanel(el);
		expect(groups.map((group) => group.hasAttribute('data-collapsed'))).toEqual([false, true]);
		const button = groups[1]!.querySelector<HTMLButtonElement>('.ribbon-overflow-button')!;
		expect(button.getAttribute('aria-label')).toBe('Styles');
		expect(button.querySelector('i')?.textContent).toBe('more');
	});

	it('shows the folded controls in a dropdown and puts them back on close', () => {
		const { el, groups } = panel(150);
		overflow.fitPanel(el);
		const group = groups[1]!;
		const run = group.querySelector('[data-command]')!;
		group.querySelector<HTMLButtonElement>('.ribbon-overflow-button')!.click();
		const pop = document.querySelector('.ribbon-overflow-panel')!;
		expect(pop.contains(run)).toBe(true);
		overflow.close();
		expect(group.contains(run)).toBe(true);
		expect(document.querySelector('.ribbon-overflow-panel')).toBeNull();
	});

	it('leaves a panel that already fits alone', () => {
		const { el, groups } = panel(500);
		overflow.fitPanel(el);
		expect(groups.some((group) => group.hasAttribute('data-collapsed'))).toBe(false);
	});
});
