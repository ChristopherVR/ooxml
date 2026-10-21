import { expect, test } from '@playwright/test';
import type { NativeHarness } from '../../demos/xlsx/demo-vanilla/native-building-blocks';
for (const framework of ['react', 'angular', 'vue', 'svelte', 'solid', 'vanilla']) {
	test(`${framework}: native composed controls retain drafts, permission and saved edits`, async ({
		page,
	}) => {
		const pageErrors: string[] = [];
		page.on('pageerror', (error) => pageErrors.push(error.message));
		await page.goto(`/?framework=${framework}&blocks=1`);
		await page.getByRole('button', { name: 'or create a New Workbook', exact: true }).click();
		await expect(page.getByRole('button', { name: 'Bold', exact: true })).toBeVisible();
		const canvas = page.locator('xlsx-editor');
		if (framework === 'vue') {
			await page.evaluate(() => window.xlsxNativeBlocks.setCanvasClass?.('updated-canvas'));
			await expect(page.locator('.updated-canvas')).toHaveCount(1);
		}
		await expect(canvas).toHaveCount(1);
		await expect(
			canvas.locator('office-ui-ribbon, office-ui-title-bar, office-ui-backstage'),
		).toHaveCount(0);
		await page.evaluate(() => window.xlsxNativeBlocks.handle.select('B20'));
		await expect(page.getByRole('textbox', { name: 'Name Box', exact: true })).toHaveCount(1);
		const formula = page.getByRole('textbox', { name: 'Formula Bar', exact: true });
		await formula.fill('=1+2');
		await formula.press('Enter');
		await page.getByRole('button', { name: 'Bold', exact: true }).click();
		await expect(page.getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute(
			'aria-pressed',
			'true',
		);
		await page.getByRole('button', { name: 'Host italic' }).click();
		await expect(page.getByRole('button', { name: 'Italic', exact: true })).toHaveAttribute(
			'aria-pressed',
			'true',
		);
		// Model updates must preserve the focused native control rather than replacing it.
		const bold = page.getByRole('button', { name: 'Bold', exact: true });
		await bold.focus();
		await page.evaluate(() => window.xlsxNativeBlocks.snapshot().toolbarProps!.onUnderline());
		await expect(bold).toBeFocused();
		const name = page.getByRole('textbox', { name: 'Name Box', exact: true });
		await name.fill('C30');
		await page.evaluate(() => window.xlsxNativeBlocks.snapshot().toolbarProps!.onUnderline());
		await expect(name).toHaveValue('C30');
		await name.press('Enter');
		await expect
			.poll(() => page.evaluate(() => window.xlsxNativeBlocks.handle.getSelection()))
			.toBe('C30');
		await page.evaluate(() => window.xlsxNativeBlocks.handle.select('B20'));
		await page.evaluate(async () => {
			const b: NativeHarness = window.xlsxNativeBlocks;
			const bytes = await b.handle.saveBytes();
			await b.handle.load(bytes);
			b.handle.select('B20');
			return b.handle.element.getToolbarProps().formula;
		});
		await expect
			.poll(() => page.evaluate(() => window.xlsxNativeBlocks.snapshot().toolbarProps?.formula))
			.toBe('=1+2');
		await page.evaluate(() => window.xlsxNativeBlocks.update({ canEdit: false }));
		await expect(bold).toBeDisabled();
		await expect(formula).toHaveAttribute('readonly', '');
		await page.evaluate(async () => {
			const b = window.xlsxNativeBlocks;
			const retained = b.snapshot().toolbarProps!;
			await b.destroy();
			if (await retained.onBold()) throw new Error('Removed canvas accepted editing');
		});
		await expect(canvas).toHaveCount(0);
		expect(pageErrors).toEqual([]);
	});

	test(`${framework}: formula controls reject, correct and cancel validation drafts`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (e) => errors.push(e.message));
		await page.goto(`/?framework=${framework}&blocks=1`);
		await page.getByRole('button', { name: 'or create a New Workbook', exact: true }).click();
		await expect(page.getByRole('button', { name: 'Bold', exact: true })).toBeVisible();
		await page.evaluate(() => {
			const element = window.xlsxNativeBlocks.handle.element;
			const book = element.workbook!;
			const first = book.sheets[0]!;
			first.dataValidations = [
				{
					ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }],
					type: 'whole',
					operator: 'between',
					formula1: '1',
					formula2: '2',
					showErrorMessage: true,
					errorStyle: 'stop',
				},
			];
			element.workbook = {
				...book,
				sheets: [first, { ...first, name: 'Second', rows: new Map(), dataValidations: [] }],
			};
			element.select('A1');
		});
		const formula = page.getByRole('textbox', { name: 'Formula Bar', exact: true });
		await formula.fill('9');
		await formula.press('Enter');
		await expect(formula).toHaveAttribute('aria-invalid', 'true');
		const alert = page.locator('xlsx-editor').getByRole('alertdialog');
		await expect(alert).toBeVisible();
		await expect
			.poll(() => page.evaluate(() => window.xlsxNativeBlocks.snapshot().editError))
			.toBeTruthy();
		await expect(page.locator('span[role="alert"]')).toBeVisible();
		const blocked = await page.evaluate(async () => {
			const p = window.xlsxNativeBlocks.snapshot().toolbarProps!;
			return { format: await p.onBold(), sheet: p.onSheetChange(1), selection: p.onSelect('B2') };
		});
		expect(blocked).toEqual({ format: false, sheet: false, selection: false });
		await alert.getByRole('button', { name: 'Retry', exact: true }).click();
		await formula.fill('2');
		await formula.press('Enter');
		await expect(formula).toHaveAttribute('aria-invalid', 'false');
		await expect
			.poll(() => page.evaluate(() => window.xlsxNativeBlocks.snapshot().editError))
			.toBeNull();
		await expect
			.poll(() => page.evaluate(() => window.xlsxNativeBlocks.snapshot().toolbarProps?.editing))
			.toBe(false);
		await formula.fill('1');
		await formula.press('Escape');
		await expect(formula).toHaveValue('2');
		await page.evaluate(async () => {
			const b = window.xlsxNativeBlocks;
			const bytes = await b.handle.saveBytes();
			await b.handle.load(bytes);
			b.handle.select('A1');
		});
		await expect(formula).toHaveValue('2');
		expect(errors).toEqual([]);
	});

	test(`${framework}: library toolbar controls format, and dialogs go to host callbacks`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (e) => errors.push(e.message));
		await page.goto(`/?framework=${framework}&blocks=1`);
		await page.getByRole('button', { name: 'or create a New Workbook', exact: true }).click();
		await expect(page.getByRole('button', { name: 'Bold', exact: true })).toBeVisible();
		await page.getByRole('button', { name: 'Center', exact: true }).click();
		await expect(page.getByRole('button', { name: 'Center', exact: true })).toHaveAttribute(
			'aria-pressed',
			'true',
		);
		await page.getByRole('combobox', { name: 'Font Size', exact: true }).selectOption('18');
		await expect
			.poll(() => page.evaluate(() => window.xlsxNativeBlocks.snapshot().toolbarProps?.fontSize))
			.toBe(18);
		await page.getByRole('combobox', { name: 'Font', exact: true }).selectOption('Arial');
		await expect
			.poll(() => page.evaluate(() => window.xlsxNativeBlocks.snapshot().toolbarProps?.fontFamily))
			.toBe('Arial');
		expect(await page.evaluate(() => window.xlsxNativeBlocks.snapshot().dirty)).toBe(true);
		// Sharing is a host workflow here: not enabled, and a silent no-op, until the host provides it.
		expect(
			await page.evaluate(() => window.xlsxNativeBlocks.snapshot().toolbarProps?.enabled.share),
		).toBe(false);
		await page.evaluate(() => {
			const b = window.xlsxNativeBlocks;
			const calls: string[] = [];
			(window as unknown as { shareCalls: string[] }).shareCalls = calls;
			b.update({ onOpenShareDialog: (context) => calls.push(context.activeCell) });
		});
		await expect
			.poll(() =>
				page.evaluate(() => window.xlsxNativeBlocks.snapshot().toolbarProps?.enabled.share),
			)
			.toBe(true);
		await page.evaluate(() => window.xlsxNativeBlocks.handle.share());
		expect(
			await page.evaluate(() => (window as unknown as { shareCalls: string[] }).shareCalls),
		).toEqual(['A1']);
		await expect(page.locator('xlsx-editor').locator('office-ui-backstage')).toHaveCount(0);
		expect(errors).toEqual([]);
	});
}
