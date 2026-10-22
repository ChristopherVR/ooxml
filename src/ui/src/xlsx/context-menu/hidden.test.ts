// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createCommandRegistry, type MenuEntry } from 'ooxml-core/xlsx/ui';
import { createTestContext } from '../grid/test-context';
import { menuEntryCommand, visibleMenuEntries } from './hidden';
import { currentContextMenu, openContextMenu } from './menu';

const hiddenIn = (ids: string[]) => (command: string) => ids.includes(command);

describe('visibleMenuEntries', () => {
	it('maps an entry to the command it runs, and an action to its ribbon command', () => {
		expect(menuEntryCommand({ id: 'cut', label: 'Cut', command: 'edit.cut' })).toBe('edit.cut');
		expect(menuEntryCommand({ id: 'rename', label: 'Rename', action: () => {} })).toBe(
			'sheet.rename',
		);
		expect(menuEntryCommand({ id: 'other', label: 'Other', action: () => {} })).toBeUndefined();
	});

	it('removes entries whose command is hidden and keeps the rest in order', () => {
		const entries: MenuEntry[] = [
			{ id: 'cut', label: 'Cut', command: 'edit.cut' },
			{ id: 'copy', label: 'Copy', command: 'edit.copy' },
			{ id: 'paste', label: 'Paste', command: 'edit.paste' },
		];
		const visible = visibleMenuEntries(entries, hiddenIn(['edit.copy']));
		expect(visible.map((entry) => entry.id)).toEqual(['cut', 'paste']);
	});

	it('removes a tab rename entry when Rename Sheet is hidden', () => {
		const entries: MenuEntry[] = [
			{ id: 'insert', label: 'Insert', command: 'sheet.insert' },
			{ id: 'rename', label: 'Rename', action: () => {} },
		];
		expect(visibleMenuEntries(entries, hiddenIn(['sheet.rename'])).map((e) => e.id)).toEqual([
			'insert',
		]);
	});

	it('leaves no separator at the top when the first group is hidden', () => {
		const entries: MenuEntry[] = [
			{ id: 'cut', label: 'Cut', command: 'edit.cut' },
			{ id: 'insert', label: 'Insert', command: 'cells.insert-rows', separatorBefore: true },
			{ id: 'format', label: 'Format', command: 'format.cells', separatorBefore: true },
		];
		const visible = visibleMenuEntries(entries, hiddenIn(['edit.cut', 'cells.insert-rows']));
		expect(visible.map((entry) => entry.id)).toEqual(['format']);
		expect(visible[0]?.separatorBefore).toBe(false);
	});

	it('keeps the separator before a group that still follows a visible entry', () => {
		const entries: MenuEntry[] = [
			{ id: 'cut', label: 'Cut', command: 'edit.cut' },
			{ id: 'insert', label: 'Insert', command: 'cells.insert-rows', separatorBefore: true },
			{ id: 'format', label: 'Format', command: 'format.cells', separatorBefore: true },
		];
		const visible = visibleMenuEntries(entries, hiddenIn(['cells.insert-rows']));
		expect(visible.map((entry) => entry.id)).toEqual(['cut', 'format']);
		expect(visible[1]?.separatorBefore).toBe(true);
	});

	it('returns an empty list when every entry is hidden', () => {
		const entries: MenuEntry[] = [{ id: 'cut', label: 'Cut', command: 'edit.cut' }];
		expect(visibleMenuEntries(entries, hiddenIn(['edit.cut']))).toEqual([]);
	});
});

afterEach(() => document.body.replaceChildren());

describe('context menu with hiddenActions', () => {
	const hiddenContext = (hidden: string[]) => {
		const ctx = createTestContext();
		const registry = createCommandRegistry(() => ctx, { isHidden: hiddenIn(hidden) });
		Object.assign(ctx, { commands: registry });
		return ctx;
	};

	it('does not draw a hidden command and cannot run it from the menu', () => {
		const ctx = hiddenContext(['test.one']);
		const ran: string[] = [];
		ctx.commands.register({ id: 'test.one', label: 'One', run: () => void ran.push('one') });
		ctx.commands.register({ id: 'test.two', label: 'Two', run: () => void ran.push('two') });
		const menu = openContextMenu(
			ctx,
			[
				{ id: 'one', label: 'One', command: 'test.one' },
				{ id: 'two', label: 'Two', command: 'test.two' },
			],
			0,
			0,
		);
		const { state } = menu.element as HTMLElement & { state: { items: { id: string }[] } };
		expect(state.items.map((item) => item.id)).toEqual(['two']);
		menu.element.dispatchEvent(new CustomEvent('office-menu-request', { detail: { id: 'one' } }));
		expect(ran).toEqual([]);
		expect(currentContextMenu(ctx)).toBe(menu);
		menu.close();
	});

	it('opens no menu when every entry is hidden', () => {
		const ctx = hiddenContext(['test.one']);
		ctx.commands.register({ id: 'test.one', label: 'One', run() {} });
		let restored = 0;
		const menu = openContextMenu(ctx, [{ id: 'one', label: 'One', command: 'test.one' }], 0, 0, {
			restoreFocus: () => void restored++,
		});
		expect(menu.element.isConnected).toBe(false);
		expect(ctx.root.querySelector('office-ui-context-menu')).toBeNull();
		expect(currentContextMenu(ctx)).toBeUndefined();
		expect(restored).toBe(1);
	});
});
