// @vitest-environment jsdom
import { createWorkbook } from 'ooxml-core/xlsx';
import { afterEach, describe, expect, it } from 'vitest';
import { createTestContext } from '../commands/test-support';
import { clearUserSortLists, userSortLists } from '../dialogs/custom-lists';
import { renderCustomListsPane } from './options-custom-lists';
import type { BackstageHost } from './parts';

afterEach(() => {
	clearUserSortLists();
	document.body.innerHTML = '';
});

function makePane() {
	const wb = createWorkbook({ sheets: ['Sheet1'] });
	const ctx = createTestContext(wb);
	const session = ctx.session()!;
	session.setCellInput(0, 0, 0, 'High');
	session.setCellInput(0, 1, 0, 'Medium');
	session.setCellInput(0, 2, 0, 'Low');
	const pane = document.createElement('div');
	document.body.append(pane);
	const host = { ctx } as unknown as BackstageHost;
	renderCustomListsPane(pane, host);
	return { pane, ctx };
}

describe('Excel Options > Advanced > Edit Custom Lists', () => {
	it('adds, edits/renames, and deletes user custom lists while keeping built-in lists read-only', () => {
		const { pane } = makePane();
		const select = pane.querySelector<HTMLSelectElement>('select')!;
		const textarea = pane.querySelector<HTMLTextAreaElement>('textarea')!;
		const buttons = [...pane.querySelectorAll<HTMLButtonElement>('button')];
		const saveBtn = buttons[0]!;
		const deleteBtn = buttons[1]!;

		// Add a new list
		textarea.value = 'Bronze\nSilver\nGold';
		saveBtn.click();
		expect(userSortLists()).toEqual([['Bronze', 'Silver', 'Gold']]);
		expect(select.value).toBe('user:0');
		expect(deleteBtn.disabled).toBe(false);

		// Edit/rename the saved user list
		textarea.value = 'Gold, Platinum, Diamond';
		saveBtn.click();
		expect(userSortLists()).toEqual([['Gold', 'Platinum', 'Diamond']]);

		// Built-in list is read-only and cannot be deleted
		select.value = 'builtin:0';
		select.dispatchEvent(new Event('change'));
		expect(textarea.readOnly).toBe(true);
		expect(saveBtn.disabled).toBe(true);
		expect(deleteBtn.disabled).toBe(true);

		// Delete the user list
		select.value = 'user:0';
		select.dispatchEvent(new Event('change'));
		deleteBtn.click();
		expect(userSortLists()).toEqual([]);
		expect(select.value).toBe('new');
	});

	it('imports a custom list from worksheet cells and reports invalid or empty ranges', () => {
		const { pane } = makePane();
		const rangeInput = pane.querySelector<HTMLInputElement>('input[type="text"]')!;
		const importBtn = [...pane.querySelectorAll<HTMLButtonElement>('button')][2]!;
		const status = pane.querySelector<HTMLElement>('[role="status"]')!;

		// Invalid range
		rangeInput.value = 'not-a-range';
		importBtn.click();
		expect(status.textContent).toContain('valid cell range');
		expect(userSortLists()).toEqual([]);

		// Empty range
		rangeInput.value = 'D1:D3';
		importBtn.click();
		expect(status.textContent).toContain('no text or number entries');
		expect(userSortLists()).toEqual([]);

		// Valid populated range A1:A3
		rangeInput.value = '$A$1:$A$3';
		importBtn.click();
		expect(status.textContent).toBe('');
		expect(userSortLists()).toEqual([['High', 'Medium', 'Low']]);
	});
});
