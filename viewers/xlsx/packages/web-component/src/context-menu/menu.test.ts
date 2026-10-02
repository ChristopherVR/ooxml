// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createTestContext } from '../grid/test-context';
import type { MenuEntry } from './items';
import { currentContextMenu, openContextMenu } from './menu';

const entries: MenuEntry[] = [
	{ id: 'one', label: 'One', command: 'test.one', shortcut: 'Ctrl+1' },
	{ id: 'two', label: 'Two', command: 'test.missing' },
	{ id: 'three', label: 'Three', action: () => {}, separatorBefore: true },
];

const key = (target: Element, k: string) =>
	target.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true }));

afterEach(() => document.body.replaceChildren());

describe('context menu DOM', () => {
	it('renders menu items, disables unknown commands and shows shortcuts', () => {
		const ctx = createTestContext();
		ctx.commands.register({ id: 'test.one', label: 'One', run() {} });
		const { element } = openContextMenu(ctx, entries, 20, 30);
		expect(element.getAttribute('role')).toBe('menu');
		expect(ctx.root.contains(element)).toBe(true);
		const items = [...element.querySelectorAll('[role="menuitem"]')];
		expect(items.map((i) => i.getAttribute('data-item'))).toEqual(['one', 'two', 'three']);
		expect(items[1]?.getAttribute('aria-disabled')).toBe('true');
		expect(element.querySelector('.xcm-shortcut')?.textContent).toBe('Ctrl+1');
		expect(element.querySelectorAll('[role="separator"]').length).toBe(1);
		expect(ctx.root.activeElement).toBe(items[0]);
	});

	it('moves focus with the keyboard, skipping disabled items', () => {
		const ctx = createTestContext();
		ctx.commands.register({ id: 'test.one', label: 'One', run() {} });
		const { element } = openContextMenu(ctx, entries, 0, 0);
		const items = [...element.querySelectorAll<HTMLElement>('[role="menuitem"]')];
		key(items[0]!, 'ArrowDown');
		expect(ctx.root.activeElement).toBe(items[2]);
		key(items[2]!, 'ArrowDown');
		expect(ctx.root.activeElement).toBe(items[0]);
		key(items[0]!, 'End');
		expect(ctx.root.activeElement).toBe(items[2]);
		key(items[2]!, 'Home');
		expect(ctx.root.activeElement).toBe(items[0]);
	});

	it('runs the command with its argument and closes', async () => {
		const ctx = createTestContext();
		const ran: unknown[] = [];
		ctx.commands.register({ id: 'test.one', label: 'One', run: (_c, arg) => void ran.push(arg) });
		let focused = 0;
		const menu = openContextMenu(
			ctx,
			[{ id: 'one', label: 'One', command: 'test.one', arg: 7 }],
			0,
			0,
			{ restoreFocus: () => void focused++ },
		);
		key(menu.element.querySelector('[data-item="one"]')!, 'Enter');
		await Promise.resolve();
		expect(ran).toEqual([7]);
		expect(menu.element.isConnected).toBe(false);
		expect(focused).toBe(1);
		expect(currentContextMenu(ctx)).toBeUndefined();
	});

	it('runs an action entry', () => {
		const ctx = createTestContext();
		let ran = 0;
		const menu = openContextMenu(ctx, [{ id: 'a', label: 'A', action: () => void ran++ }], 0, 0);
		menu.element.querySelector<HTMLElement>('[data-item="a"]')!.click();
		expect(ran).toBe(1);
	});

	it('closes on Escape and keeps one menu per context', () => {
		const ctx = createTestContext();
		let closed = 0;
		const first = openContextMenu(ctx, entries, 0, 0, { onClose: () => void closed++ });
		const second = openContextMenu(ctx, entries, 0, 0);
		expect(first.element.isConnected).toBe(false);
		expect(closed).toBe(1);
		expect(currentContextMenu(ctx)).toBe(second);
		key(second.element, 'Escape');
		expect(second.element.isConnected).toBe(false);
	});

	it('dismisses on an outside pointerdown', async () => {
		const ctx = createTestContext();
		const menu = openContextMenu(ctx, entries, 0, 0);
		await new Promise((resolve) => setTimeout(resolve, 0));
		menu.element.dispatchEvent(new Event('pointerdown', { bubbles: true, composed: true }));
		expect(menu.element.isConnected).toBe(true);
		document.body.dispatchEvent(new Event('pointerdown', { bubbles: true }));
		expect(menu.element.isConnected).toBe(false);
	});

	it('translates labels through ctx.t', () => {
		const ctx = createTestContext();
		Object.assign(ctx, { t: (k: string) => `[${k}]` });
		const menu = openContextMenu(ctx, entries, 0, 0);
		expect(menu.element.querySelector('.xcm-label')?.textContent).toBe('[One]');
	});
});
