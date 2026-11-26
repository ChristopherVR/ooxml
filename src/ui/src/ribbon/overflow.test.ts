// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createRibbonOverflow, type RibbonOverflowOptions } from './overflow';

const options = {
	icon: (doc, name) => Object.assign(doc.createElement('i'), { textContent: name }),
	groupIcons: { Font: 'font' },
	fallbackIcon: 'more',
	ribbon: '.ribbon',
	command: 'button[data-command]',
} satisfies RibbonOverflowOptions;
const overflow = createRibbonOverflow(options);

/** jsdom has no layout: a panel that is `room` wide holding groups of 100 that overflow it. */
function panel(room: number, collapsedWidth = 40, widths = [100, 100]) {
	const ribbon = document.createElement('div');
	ribbon.className = 'ribbon';
	const el = document.createElement('div');
	el.className = 'ribbon-panel';
	const define = (node: HTMLElement, props: Record<string, () => number>) => {
		for (const [key, get] of Object.entries(props)) Object.defineProperty(node, key, { get });
	};
	const groups = ['Font', 'Styles'].map((label, index) => {
		const group = document.createElement('div');
		group.className = 'ribbon-group';
		group.dataset.label = label;
		group.dataset.caption = label;
		const run = document.createElement('button');
		run.dataset.command = 'x';
		group.append(run);
		define(group, {
			offsetWidth: () => (group.hasAttribute('data-collapsed') ? collapsedWidth : widths[index]!),
		});
		el.append(group);
		return group;
	});
	define(el, {
		clientWidth: () => room,
		scrollWidth: () => groups.reduce((sum, group) => sum + group.offsetWidth, 0),
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

	it('keeps a preferred group expanded while a sibling can fold to make room', () => {
		const { el, groups } = panel(150);
		const preferred = createRibbonOverflow({
			...options,
			preferExpanded: (group) => group.dataset.label === 'Styles',
		});
		preferred.fitPanel(el);
		expect(groups.map((group) => group.hasAttribute('data-collapsed'))).toEqual([true, false]);
	});

	it('folds a lower-priority group first even if overflow is only one pixel', () => {
		const { el, groups } = panel(199);
		createRibbonOverflow({
			...options,
			groupPriority: (group) => Number(group.dataset.label === 'Styles'),
		}).fitPanel(el);
		expect(groups.map((group) => group.hasAttribute('data-collapsed'))).toEqual([true, false]);
		expect(el.scrollWidth).toBeLessThanOrEqual(el.clientWidth);
	});

	it('keeps groups expanded when folding them saves no space', () => {
		const { el, groups } = panel(150, 100);
		overflow.fitPanel(el);
		expect(groups.some((group) => group.hasAttribute('data-collapsed'))).toBe(false);
	});

	it('restores a narrow group when folding a wide group frees enough room', () => {
		const { el, groups } = panel(170, 40, [160, 70]);
		overflow.fitPanel(el);
		expect(groups.map((group) => group.hasAttribute('data-collapsed'))).toEqual([true, false]);
	});

	it('consumes Escape while closing a dropdown from outside the ribbon', () => {
		const { el, groups } = panel(150);
		overflow.fitPanel(el);
		const toggle = groups[1]!.querySelector<HTMLButtonElement>('.ribbon-overflow-button')!;
		const editor = document.createElement('button');
		document.body.append(editor);
		editor.focus();
		let received = 0;
		editor.addEventListener('keydown', () => received++);
		toggle.click();
		const press = () =>
			editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', cancelable: true }));
		expect(press()).toBe(false);
		expect(document.querySelector('.ribbon-overflow-panel')).toBeNull();
		expect(document.activeElement).toBe(toggle);
		expect(received).toBe(0);
		expect(press()).toBe(true);
		expect(received).toBe(1);
	});

	it('ends the dropdown on a command but keeps it open for a button that opens a popup', async () => {
		const { el, groups } = panel(150);
		overflow.fitPanel(el);
		const group = groups[1]!;
		const run = group.querySelector<HTMLButtonElement>('[data-command]')!;
		const gallery = document.createElement('button');
		gallery.dataset.command = 'gallery';
		gallery.setAttribute('aria-haspopup', 'dialog');
		group.append(gallery);
		const toggle = group.querySelector<HTMLButtonElement>('.ribbon-overflow-button')!;
		const tick = () => new Promise((done) => setTimeout(done, 0));
		toggle.click();
		gallery.click();
		await tick();
		expect(document.querySelector('.ribbon-overflow-panel')?.contains(gallery)).toBe(true);
		run.click();
		await tick();
		expect(document.querySelector('.ribbon-overflow-panel')).toBeNull();
		expect(group.contains(gallery)).toBe(true);
	});

	it('leaves a panel that already fits alone', () => {
		const { el, groups } = panel(500);
		overflow.fitPanel(el);
		expect(groups.some((group) => group.hasAttribute('data-collapsed'))).toBe(false);
	});
});
