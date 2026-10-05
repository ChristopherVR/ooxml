// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createTestContext } from '../grid/test-context';
import type { MenuEntry } from 'ooxml-core/xlsx/ui';
import { currentContextMenu, openContextMenu } from './menu';

const entries: MenuEntry[] = [
	{ id: 'one', label: 'One', command: 'test.one', shortcut: 'Ctrl+1' },
	{ id: 'two', label: 'Two', command: 'test.missing' },
	{ id: 'three', label: 'Three', action: () => {}, separatorBefore: true },
];

type Menu = HTMLElement & { updateComplete: Promise<unknown> };
/** The rows the shared element drew, once it has rendered. */
const rowsOf = async (menu: { element: HTMLElement }) => {
	await (menu.element as Menu).updateComplete;
	return [...menu.element.shadowRoot!.querySelectorAll<HTMLButtonElement>('[role="menuitem"]')];
};
const request = (menu: { element: HTMLElement }, id: string) =>
	menu.element.dispatchEvent(new CustomEvent('office-menu-request', { detail: { id } }));

afterEach(() => document.body.replaceChildren());

describe('context menu', () => {
	it('draws the entries with the shared element, disabling unknown commands', async () => {
		const ctx = createTestContext();
		ctx.commands.register({ id: 'test.one', label: 'One', run() {} });
		const menu = openContextMenu(ctx, entries, 20, 30);
		expect(menu.element.localName).toBe('office-ui-context-menu');
		expect(ctx.root.contains(menu.element)).toBe(true);
		const rows = await rowsOf(menu);
		expect(rows.map((row) => row.dataset.itemId)).toEqual(['one', 'two', 'three']);
		expect(rows.map((row) => row.disabled)).toEqual([false, true, false]);
		expect(rows[0]?.querySelector('.shortcut')?.textContent).toBe('Ctrl+1');
		expect(menu.element.shadowRoot!.querySelectorAll('[role="separator"]')).toHaveLength(1);
	});

	it('runs the command with its argument, closes and restores focus', async () => {
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
		request(menu, 'one');
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
		request(menu, 'a');
		expect(ran).toBe(1);
	});

	it('keeps one menu per context and closes on a dismissal event', () => {
		const ctx = createTestContext();
		let closed = 0;
		const first = openContextMenu(ctx, entries, 0, 0, { onClose: () => void closed++ });
		const second = openContextMenu(ctx, entries, 0, 0);
		expect(first.element.isConnected).toBe(false);
		expect(closed).toBe(1);
		expect(currentContextMenu(ctx)).toBe(second);
		second.element.dispatchEvent(
			new CustomEvent('office-menu-close', { detail: { reason: 'escape' } }),
		);
		expect(second.element.isConnected).toBe(false);
	});

	it('restores focus after Escape but not after an outside click', () => {
		const ctx = createTestContext();
		let focused = 0;
		const restoreFocus = () => void focused++;
		const escaped = openContextMenu(ctx, entries, 0, 0, { restoreFocus });
		escaped.element.dispatchEvent(
			new CustomEvent('office-menu-close', { detail: { reason: 'escape' } }),
		);
		expect(focused).toBe(1);
		const outside = openContextMenu(ctx, entries, 0, 0, { restoreFocus });
		outside.element.dispatchEvent(
			new CustomEvent('office-menu-close', { detail: { reason: 'outside' } }),
		);
		expect(focused).toBe(1);
	});

	it('translates labels through ctx.t', async () => {
		const ctx = createTestContext();
		Object.assign(ctx, { t: (k: string) => `[${k}]` });
		const menu = openContextMenu(ctx, entries, 0, 0);
		expect((await rowsOf(menu))[0]?.textContent).toContain('[One]');
	});
});
