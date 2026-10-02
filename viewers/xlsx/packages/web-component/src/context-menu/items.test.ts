import { describe, expect, it } from 'vitest';
import {
	cellMenu,
	clampToViewport,
	columnHeaderMenu,
	rowHeaderMenu,
	tabMenu,
	type MenuEntry,
} from './items';

const ids = (entries: MenuEntry[]) => entries.map((e) => e.id);
const disabled = (entries: MenuEntry[]) => entries.filter((e) => e.disabled).map((e) => e.id);
const cell = { readOnly: false, hasComment: false, hasLink: false, hasList: false };

describe('cell menu', () => {
	it('lists the Excel cell items in groups', () => {
		const entries = cellMenu(cell);
		expect(ids(entries)).toEqual([
			'cut',
			'copy',
			'paste',
			'paste-special',
			'insert-cells',
			'delete-cells',
			'clear-contents',
			'filter',
			'sort-ascending',
			'sort-descending',
			'custom-sort',
			'new-comment',
			'format-cells',
			'pick-from-list',
			'define-name',
			'link',
		]);
		expect(entries.filter((e) => e.separatorBefore).map((e) => e.id)).toEqual([
			'insert-cells',
			'filter',
			'new-comment',
			'format-cells',
		]);
		expect(entries.find((e) => e.id === 'copy')?.command).toBe('edit.copy');
		expect(entries.find((e) => e.id === 'format-cells')?.shortcut).toBe('Ctrl+1');
	});

	it('offers edit and delete for comments and links', () => {
		const list = ids(cellMenu({ ...cell, hasComment: true, hasLink: true }));
		expect(list).toContain('edit-comment');
		expect(list).toContain('delete-comment');
		expect(list).not.toContain('new-comment');
		expect(list).toContain('remove-link');
		expect(list).not.toContain('link');
	});

	it('keeps only Copy enabled in read-only mode', () => {
		const entries = cellMenu({ ...cell, readOnly: true });
		expect(entries.filter((e) => !e.disabled).map((e) => e.id)).toEqual(['copy']);
	});
});

describe('header menus', () => {
	it('maps row items to row commands', () => {
		const entries = rowHeaderMenu({ readOnly: false, hiddenInSelection: true });
		expect(ids(entries)).toContain('row-height');
		expect(entries.find((e) => e.id === 'insert')?.command).toBe('cells.insert-rows');
		expect(entries.find((e) => e.id === 'hide')?.command).toBe('format.hide-rows');
		expect(disabled(entries)).toEqual([]);
	});

	it('maps column items to column commands and disables unhide without hidden columns', () => {
		const entries = columnHeaderMenu({ readOnly: false, hiddenInSelection: false });
		expect(entries.find((e) => e.id === 'column-width')?.command).toBe('format.column-width');
		expect(entries.find((e) => e.id === 'delete')?.command).toBe('cells.delete-columns');
		expect(disabled(entries)).toEqual(['unhide']);
	});
});

describe('tab menu', () => {
	const state = {
		readOnly: false,
		structureLocked: false,
		visibleSheets: 2,
		hiddenSheets: 0,
		protected: false,
	};
	it('lists the sheet items and wires rename to the caller', () => {
		let renamed = 0;
		const entries = tabMenu(state, { rename: () => void renamed++ });
		expect(ids(entries)).toEqual([
			'insert',
			'delete',
			'rename',
			'move-copy',
			'tab-color',
			'hide',
			'unhide',
			'protect',
		]);
		void entries.find((e) => e.id === 'rename')?.action?.();
		expect(renamed).toBe(1);
		expect(disabled(entries)).toEqual(['unhide']);
		expect(entries.find((e) => e.id === 'protect')?.label).toBe('Protect Sheet...');
	});

	it('protects the last visible sheet and honours locks', () => {
		const one = tabMenu({ ...state, visibleSheets: 1, hiddenSheets: 1 }, { rename() {} });
		expect(disabled(one)).toEqual(['delete', 'hide']);
		const locked = tabMenu({ ...state, structureLocked: true, protected: true }, { rename() {} });
		expect(disabled(locked)).toEqual(['insert', 'delete', 'rename', 'move-copy', 'hide', 'unhide']);
		expect(locked.find((e) => e.id === 'protect')?.label).toBe('Unprotect Sheet');
		const readOnly = tabMenu({ ...state, readOnly: true }, { rename() {} });
		expect(readOnly.every((e) => e.disabled)).toBe(true);
	});
});

describe('clampToViewport', () => {
	it('keeps the menu inside the viewport', () => {
		const viewport = { width: 1000, height: 800 };
		expect(clampToViewport(990, 750, { width: 200, height: 100 }, viewport)).toEqual({
			left: 796,
			top: 696,
		});
		expect(clampToViewport(10, 10, { width: 50, height: 50 }, viewport)).toEqual({
			left: 10,
			top: 10,
		});
	});
});
