// @vitest-environment jsdom
import { createWorkbook } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	pressKey,
	setValue,
	tick,
} from '../commands/test-support.js';
import { registerNavigationDialogs } from './register-navigation.js';

afterEach(() => (document.body.innerHTML = ''));

function setup() {
	const ctx = createTestContext(createWorkbook({ sheets: ['Sheet1', 'Other Sheet'] }));
	registerNavigationDialogs(ctx);
	return ctx;
}

describe('define name dialog', () => {
	it('defaults Refers to to the absolute selection and adds the name', async () => {
		const ctx = setup();
		ctx.select('A1:B2');
		const done = ctx.dialogs.open('define-name');
		const dialog = dialogEl(ctx, 'define-name');
		expect(inputByLabel(dialog, 'Refers to:').value).toBe('=Sheet1!$A$1:$B$2');
		setValue(inputByLabel(dialog, 'Name:'), 'Sales');
		setValue(inputByLabel<HTMLTextAreaElement>(dialog, 'Comment:'), 'Q1');
		clickButton(dialog, 'OK');
		await done;
		expect(ctx.workbook()!.definedNames).toEqual([
			{ name: 'Sales', formula: 'Sheet1!$A$1:$B$2', comment: 'Q1' },
		]);
	});

	it('rejects invalid and duplicate names', async () => {
		const ctx = setup();
		ctx.session()!.setDefinedName({ name: 'Taken', formula: '1' });
		void ctx.dialogs.open('define-name');
		const dialog = dialogEl(ctx, 'define-name');
		setValue(inputByLabel(dialog, 'Name:'), 'A1');
		clickButton(dialog, 'OK');
		await tick();
		setValue(inputByLabel(dialog, 'Name:'), 'taken');
		clickButton(dialog, 'OK');
		await tick();
		expect(ctx.toasts.map((t) => t.message)).toEqual([
			'A name cannot look like a cell reference.',
			'The name that you entered already exists. Enter a unique name.',
		]);
		expect(ctx.workbook()!.definedNames.length).toBe(1);
	});

	it('renames an existing name in one undo step; Escape closes', async () => {
		const ctx = setup();
		const s = ctx.session()!;
		s.setDefinedName({ name: 'Old', formula: 'Sheet1!$A$1' });
		const done = ctx.dialogs.open('define-name', { name: 'Old' });
		const dialog = dialogEl(ctx, 'define-name');
		setValue(inputByLabel(dialog, 'Name:'), 'New');
		clickButton(dialog, 'OK');
		await done;
		expect(ctx.workbook()!.definedNames.map((n) => n.name)).toEqual(['New']);
		s.undo();
		expect(ctx.workbook()!.definedNames.map((n) => n.name)).toEqual(['Old']);
		const again = ctx.dialogs.open('define-name');
		pressKey(inputByLabel(dialogEl(ctx, 'define-name'), 'Name:'), 'Escape');
		expect(await again).toBeUndefined();
	});
});

describe('name manager', () => {
	it('lists names with values and scope, filters and deletes', async () => {
		const ctx = setup();
		const s = ctx.session()!;
		s.setCellInput(0, 0, 0, '42');
		s.setDefinedName({ name: 'Answer', formula: 'Sheet1!$A$1' });
		s.setDefinedName({ name: 'Local', formula: '#REF!', localSheet: 1 });
		s.setDefinedName({ name: '_xlnm.Print_Area', formula: 'Sheet1!$A$1', localSheet: 0 });
		const done = ctx.dialogs.open('name-manager');
		const dialog = dialogEl(ctx, 'name-manager');
		const rows = () => [...dialog.querySelectorAll<HTMLTableRowElement>('tbody tr')];
		expect(rows().map((r) => [...r.cells].map((c) => c.textContent))).toEqual([
			['Answer', '42', '=Sheet1!$A$1', 'Workbook', ''],
			['Local', '#REF!', '=#REF!', 'Other Sheet', ''],
		]);
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Filter'), 'errors');
		expect(rows().map((r) => r.dataset.name)).toEqual(['Local']);
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Filter'), 'all');
		rows()[0]!.click();
		clickButton(dialog, 'Delete');
		expect(ctx.workbook()!.definedNames.map((n) => n.name)).toEqual(['Local', '_xlnm.Print_Area']);
		expect(rows().length).toBe(1);
		clickButton(dialog, 'Close');
		expect(await done).toBeUndefined();
	});

	it('New... opens Define Name and refreshes the list', async () => {
		const ctx = setup();
		void ctx.dialogs.open('name-manager');
		clickButton(dialogEl(ctx, 'name-manager'), 'New...');
		const define = dialogEl(ctx, 'define-name');
		setValue(inputByLabel(define, 'Name:'), 'Fresh');
		clickButton(define, 'OK');
		await tick();
		await tick();
		expect(dialogEl(ctx, 'name-manager').querySelector('tbody tr')?.getAttribute('data-name')).toBe(
			'Fresh',
		);
	});
});

describe('create names from selection', () => {
	it('creates one name per column from the top row', async () => {
		const ctx = setup();
		const s = ctx.session()!;
		s.setCellInput(0, 0, 0, 'Unit Price');
		s.setCellInput(0, 0, 1, '2024 Q');
		s.setCellInput(0, 1, 0, '5');
		s.setCellInput(0, 1, 1, '6');
		ctx.select('A1:B3');
		const done = ctx.dialogs.open('create-names');
		const dialog = dialogEl(ctx, 'create-names');
		expect(inputByLabel(dialog, 'Top row').checked).toBe(true);
		clickButton(dialog, 'OK');
		expect(await done).toBe(2);
		expect(ctx.workbook()!.definedNames).toEqual([
			{ name: 'Unit_Price', formula: 'Sheet1!$A$2:$A$3' },
			{ name: '_2024_Q', formula: 'Sheet1!$B$2:$B$3' },
		]);
		s.undo();
		expect(ctx.workbook()!.definedNames).toEqual([]);
	});

	it('Cancel creates nothing', async () => {
		const ctx = setup();
		ctx.session()!.setCellInput(0, 0, 0, 'Head');
		ctx.select('A1:A3');
		const done = ctx.dialogs.open('create-names');
		clickButton(dialogEl(ctx, 'create-names'), 'Cancel');
		expect(await done).toBeUndefined();
		expect(ctx.workbook()!.definedNames).toEqual([]);
	});
});
