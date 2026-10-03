// @vitest-environment jsdom
import { createWorkbook, getCell } from '@christophervr/xlsx-core';
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
import { safeAddress } from './hyperlink.js';
import { registerRuleDialogs } from './register-rules.js';

afterEach(() => (document.body.innerHTML = ''));

function setup() {
	const ctx = createTestContext(createWorkbook({ sheets: ['Sheet1', 'My Data'] }));
	registerRuleDialogs(ctx);
	ctx.select('B2');
	return ctx;
}
const links = (ctx: ReturnType<typeof setup>) => ctx.workbook()!.sheets[0]!.hyperlinks;
const linkTo = (d: HTMLElement, label: string) => {
	const radio = inputByLabel(d, label);
	radio.checked = true;
	radio.dispatchEvent(new Event('change', { bubbles: true }));
};

describe('safeAddress', () => {
	it('allows web, mail, ftp, file and paths, and rejects script schemes', () => {
		expect(safeAddress('www.example.com')).toBe('https://www.example.com');
		expect(safeAddress('https://x.org/a')).toBe('https://x.org/a');
		expect(safeAddress('mailto:a@b.c')).toBe('mailto:a@b.c');
		expect(safeAddress('C:\\docs\\a.xlsx')).toBe('C:\\docs\\a.xlsx');
		expect(safeAddress('report.xlsx')).toBe('report.xlsx');
		expect(safeAddress('javascript:alert(1)')).toBeUndefined();
		expect(safeAddress(' JavaScript:alert(1)')).toBeUndefined();
		expect(safeAddress('data:text/html,x')).toBeUndefined();
		expect(safeAddress('vbscript:x')).toBeUndefined();
		expect(safeAddress('java\tscript:alert(1)')).toBeUndefined();
	});
});

describe('hyperlink dialog', () => {
	it('inserts a web link and writes the display text', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('hyperlink');
		const d = dialogEl(ctx, 'hyperlink');
		setValue(inputByLabel(d, 'Address:'), 'www.example.com');
		setValue(inputByLabel(d, 'ScreenTip:'), 'Go');
		clickButton(d, 'OK');
		await result;
		expect(links(ctx)[0]).toMatchObject({
			target: 'https://www.example.com',
			tooltip: 'Go',
			display: 'https://www.example.com',
		});
		expect(getCell(ctx.workbook()!.sheets[0]!, 1, 1)?.value).toBe('https://www.example.com');
	});

	it('rejects javascript: and keeps the dialog open', async () => {
		const ctx = setup();
		void ctx.dialogs.open('hyperlink');
		const d = dialogEl(ctx, 'hyperlink');
		setValue(inputByLabel(d, 'Address:'), 'javascript:alert(1)');
		clickButton(d, 'OK');
		await tick();
		expect(ctx.toasts).toHaveLength(1);
		expect(links(ctx)).toEqual([]);
		expect(dialogEl(ctx, 'hyperlink')).toBeTruthy();
	});

	it('links to a place in the workbook and to an e-mail address', async () => {
		const ctx = setup();
		let result = ctx.dialogs.open('hyperlink');
		let d = dialogEl(ctx, 'hyperlink');
		linkTo(d, 'Place in This Document');
		setValue(inputByLabel<HTMLSelectElement>(d, 'Or select a place in this document:'), 'My Data');
		setValue(inputByLabel(d, 'Type the cell reference:'), 'c3');
		setValue(inputByLabel(d, 'Text to display:'), 'Jump');
		clickButton(d, 'OK');
		await result;
		expect(links(ctx)[0]).toMatchObject({ location: "'My Data'!C3", display: 'Jump' });
		ctx.select('D4');
		result = ctx.dialogs.open('hyperlink');
		d = dialogEl(ctx, 'hyperlink');
		linkTo(d, 'E-mail Address');
		setValue(inputByLabel(d, 'E-mail address:'), 'a@b.co');
		setValue(inputByLabel(d, 'Subject:'), 'Hi there');
		clickButton(d, 'OK');
		await result;
		expect(links(ctx)[1]?.target).toBe('mailto:a@b.co?subject=Hi%20there');
	});

	it('edits and removes an existing link; Cancel and Escape change nothing', async () => {
		const ctx = setup();
		ctx
			.session()!
			.setHyperlink(
				0,
				{ start: { row: 1, col: 1 }, end: { row: 1, col: 1 } },
				{ target: 'https://a.b' },
			);
		const cancel = ctx.dialogs.open('hyperlink');
		let d = dialogEl(ctx, 'hyperlink');
		expect(inputByLabel(d, 'Address:').value).toBe('https://a.b');
		pressKey(inputByLabel(d, 'Address:'), 'Escape');
		expect(await cancel).toBeUndefined();
		void ctx.dialogs.open('hyperlink');
		d = dialogEl(ctx, 'hyperlink');
		clickButton(d, 'Remove Link');
		await tick();
		expect(links(ctx)).toEqual([]);
		expect(ctx.root.querySelector('[data-dialog="hyperlink"]')).toBeNull();
	});
});
