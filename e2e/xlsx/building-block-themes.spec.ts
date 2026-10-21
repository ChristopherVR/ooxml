import { expect, test, type Page } from '@playwright/test';
type XlsxEditorElement = HTMLElementTagNameMap['xlsx-editor'];

async function colours(page: Page) {
	return page.evaluate(() => {
		const scope = document.getElementById('editor')!;
		const probe = document.createElement('span');
		probe.style.cssText =
			'background:var(--office-surface);color:var(--office-foreground);border:1px solid var(--office-border)';
		scope.append(probe);
		const expected = getComputedStyle(probe);
		const palette = {
			surface: expected.backgroundColor,
			foreground: expected.color,
			border: expected.borderColor,
		};
		probe.style.background = 'var(--office-background)';
		const background = getComputedStyle(probe).backgroundColor;
		probe.remove();
		const editor = scope.querySelector('xlsx-editor')!;
		const toolbar = scope.querySelector('[data-office-block="toolbar"]')!;
		const formula = scope.querySelector<HTMLInputElement>(
			'[data-office-block="formula-bar"] input:last-child',
		)!;
		const header = editor.shadowRoot!.querySelector('.xg-hbox')!;
		const grid = editor.shadowRoot!.querySelector('.xg-root')!;
		return {
			palette,
			background,
			toolbar: getComputedStyle(toolbar).backgroundColor,
			input: getComputedStyle(formula).backgroundColor,
			foreground: getComputedStyle(formula).color,
			border: getComputedStyle(formula).borderColor,
			header: getComputedStyle(header).backgroundColor,
			paper: getComputedStyle(grid).backgroundColor,
			dirty: editor.dirty,
			selection: editor.getSelection(),
			inputIdentity: formula.dataset.identity,
			font: getComputedStyle(formula).fontFamily,
			headerFont: getComputedStyle(header).fontFamily,
			headerSize: getComputedStyle(editor.shadowRoot!.querySelector('.xg-hd')!).fontSize,
			canvasFont: getComputedStyle(editor).fontFamily,
			fontSize: getComputedStyle(formula).fontSize,
			radius: getComputedStyle(formula).borderRadius,
			ring: getComputedStyle(formula).outlineColor,
			rows: editor.workbook?.sheets[0]?.rows.size,
		};
	});
}
for (const framework of ['react', 'angular', 'vue', 'svelte', 'solid']) {
	test(`${framework}: canonical, inherited, custom and legacy themes preserve live formula controls`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.emulateMedia({ colorScheme: 'light' });
		await page.goto(`/?framework=${framework}&blocks=1`);
		await page.getByRole('button', { name: 'or create a New Workbook', exact: true }).click();
		await expect(page.locator('[data-office-block="toolbar"]')).toBeVisible();
		await expect(page.locator('[data-office-block]')).toHaveCount(3);
		await page.evaluate(() => {
			document.getElementById('editor')!.dataset.officeTheme = 'light';
			window.xlsxNativeBlocks.handle.element.theme = 'auto';
		});
		const light = await colours(page);
		expect(light.toolbar).toBe(light.palette.surface);
		expect(light.header).toBe(light.palette.surface);
		expect(light.input).toBe(light.background);
		expect(light.foreground).toBe(light.palette.foreground);
		expect(light.paper).toBe('rgb(255, 255, 255)');
		const input = page.getByRole('textbox', { name: 'Formula Bar', exact: true });
		await input.fill('=1+');
		await input.evaluate((element) => {
			element.dataset.identity = 'original';
		});
		for (const mode of ['dark', 'light', 'dark']) {
			await page.evaluate((mode) => {
				document.getElementById('editor')!.dataset.officeTheme = mode;
			}, mode);
			const next = await colours(page);
			expect(next.toolbar).toBe(next.palette.surface);
			expect(next.header).toBe(next.palette.surface);
			expect(next.input).toBe(next.background);
			expect(next.foreground).toBe(next.palette.foreground);
			if (mode === 'dark') expect(next.toolbar).not.toBe(light.toolbar);
			expect(next.paper).toBe(light.paper);
			expect(next.dirty).toBe(false);
			expect(next.rows).toBe(0);
			expect(next.selection).toBe('A1');
			expect(next.inputIdentity).toBe('original');
			await expect(input).toBeFocused();
			await expect(input).toHaveValue('=1+');
		}
		await page.evaluate(() => {
			const scope = document.getElementById('editor')!;
			scope.dataset.officeTheme = 'auto';
		});
		await page.emulateMedia({ colorScheme: 'dark' });
		expect((await colours(page)).toolbar).not.toBe(light.toolbar);
		await page.emulateMedia({ colorScheme: 'light' });
		expect((await colours(page)).toolbar).toBe(light.toolbar);
		await page.evaluate(() => {
			const scope = document.getElementById('editor')!;
			scope.style.setProperty('--office-surface', '#17324d');
			scope.style.setProperty('--office-background', '#09111a');
			scope.style.setProperty('--office-foreground', '#eef5fc');
			scope.style.setProperty('--office-border', '#6a8bad');
			scope.style.setProperty('--office-font', '"Courier New",monospace');
			scope.style.setProperty('--office-font-size', '15px');
			scope.style.setProperty('--office-font-size-xs', '14px');
			scope.style.setProperty('--office-radius', '9px');
			scope.style.setProperty('--office-ring', '#fd8a55');
		});
		const custom = await colours(page);
		expect(custom.toolbar).toBe('rgb(23, 50, 77)');
		expect(custom.header).toBe(custom.toolbar);
		expect(custom.input).toBe('rgb(9, 17, 26)');
		expect(custom.foreground).toBe('rgb(238, 245, 252)');
		expect(custom.font).toContain('Courier New');
		expect(custom.headerFont).toContain('Courier New');
		expect(custom.canvasFont).toContain('Courier New');
		expect(custom.headerSize).toBe('14px');
		expect(custom.fontSize).toBe('15px');
		expect(custom.radius).toBe('9px');
		expect(custom.ring).toBe('rgb(253, 138, 85)');
		await page.evaluate(() => {
			const scope = document.getElementById('editor')!;
			scope.style.setProperty('--xve-card', '#435a71');
			scope.style.setProperty('--xve-foreground', '#fef8ed');
		});
		const legacy = await colours(page);
		expect(legacy.toolbar).toBe('rgb(67, 90, 113)');
		expect(legacy.header).toBe(legacy.toolbar);
		expect(legacy.foreground).toBe('rgb(254, 248, 237)');
		await expect(input).toBeFocused();
		await expect(input).toHaveValue('=1+');
		await input.press('Escape');
		await page.evaluate(() => window.xlsxNativeBlocks.update({ canEdit: false }));
		const bold = page.getByRole('button', { name: 'Bold', exact: true });
		await expect(bold).toBeDisabled();
		await expect(bold).toHaveCSS('opacity', '0.55');
		expect(errors).toEqual([]);
	});
	test(`${framework}: native controls inherit a theme scoped inside a shadow root`, async ({
		page,
	}) => {
		await page.emulateMedia({ colorScheme: 'light' });
		await page.goto(`/?framework=${framework}&blocks=1&shadow=1`);
		await page.getByRole('button', { name: 'or create a New Workbook', exact: true }).click();
		const toolbar = page.locator('[data-office-block="toolbar"]');
		await expect(toolbar).toBeVisible();
		await expect(toolbar).toHaveCSS('background-color', 'rgb(31, 41, 55)');
		const input = page.getByRole('textbox', { name: 'Formula Bar', exact: true });
		await input.fill('=7+');
		await page.evaluate(() => {
			const element = window.xlsxNativeBlocks.handle.element;
			const root = element.getRootNode() as ShadowRoot;
			(root.getElementById('shadow-scope') as HTMLElement).dataset.officeTheme = 'light';
		});
		await expect(toolbar).toHaveCSS('background-color', 'rgb(243, 244, 246)');
		await expect(input).toBeFocused();
		await expect(input).toHaveValue('=7+');
	});
}

test('full Excel chrome keeps theme and themeColors compatibility without changing workbook data', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await page.getByRole('button', { name: 'or create a New Workbook', exact: true }).click();
	const editor = page.locator('xlsx-editor');
	await expect(editor.locator('.xve-frame')).toBeVisible();
	const result = await editor.evaluate(async (node) => {
		const element = node as XlsxEditorElement;
		element.getToolbarProps().onFormulaChange('Document text');
		await new Promise(requestAnimationFrame);
		element.markClean();
		const model = element.workbook;
		const cell = element.shadowRoot!.querySelector('.xg-t')!;
		const documentFont = getComputedStyle(cell).fontFamily;
		const scope = document.getElementById('editor')!;
		scope.style.setProperty('--office-font', '"Courier New",monospace');
		scope.style.setProperty('--office-font-size', '16px');
		scope.style.setProperty('--office-font-size-xs', '14px');
		element.theme = 'dark';
		const mode = element.dataset.officeTheme;
		element.themeColors = { card: '#123456', primary: '#ab2468', headerBg: '#234567' };
		const frame = element.shadowRoot!.querySelector('.xve-frame')!;
		const shared = element.shadowRoot!.querySelector('office-ui-ribbon')!;
		const colour = getComputedStyle(frame).backgroundColor;
		const primary = getComputedStyle(shared).getPropertyValue('--office-accent').trim();
		const header = getComputedStyle(element.shadowRoot!.querySelector('.xg-hbox')!).backgroundColor;
		element.theme = 'auto';
		return {
			mode,
			marker: element.dataset.officeTheme,
			colour,
			primary,
			header,
			font: getComputedStyle(frame).fontFamily,
			fontSize: getComputedStyle(frame).fontSize,
			formulaFont: getComputedStyle(
				element.shadowRoot!.querySelector('.xfb-input .xg-field-input')!,
			).fontFamily,
			documentFont,
			documentFontAfter: getComputedStyle(cell).fontFamily,
			sameModel: model === element.workbook,
			rows: element.workbook?.sheets[0]?.rows.size,
			dirty: element.dirty,
		};
	});
	expect(result.mode).toBe('dark');
	expect(result.marker).toBeUndefined();
	expect(result.colour).toBe('rgb(18, 52, 86)');
	expect(result.primary).toBe('#ab2468');
	expect(result.header).toBe('rgb(35, 69, 103)');
	expect(result.font).toContain('Courier New');
	expect(result.formulaFont).toContain('Courier New');
	expect(result.fontSize).toBe('16px');
	expect(result.documentFontAfter).toBe(result.documentFont);
	expect(result.documentFontAfter).not.toContain('Courier New');
	expect(result.sameModel).toBe(true);
	expect(result.rows).toBe(1);
	expect(result.dirty).toBe(false);
	await editor.locator('[data-id="pageLayout"]').click();
	const pageHeader = editor.locator('.xpl-colhdr').first();
	await expect(pageHeader).toBeVisible();
	await expect(pageHeader).toHaveCSS('font-size', '14px');
	expect(await pageHeader.evaluate((node) => getComputedStyle(node).fontFamily)).toContain(
		'Courier New',
	);
	const documentCell = editor.locator('.xpl-inner td').first();
	expect(await documentCell.evaluate((node) => getComputedStyle(node).fontFamily)).not.toContain(
		'Courier New',
	);
	await editor.locator('[data-id="normal"]').click();
	await editor.evaluate((node) => (node as XlsxEditorElement).focusGrid());
	await page.keyboard.press('Control+1');
	const dialog = editor.locator('.xve-cmd-dialog');
	await expect(dialog).toBeVisible();
	const dialogTheme = await dialog.evaluate((node) => ({
		border: getComputedStyle(node).getPropertyValue('--office-border').trim(),
		ring: getComputedStyle(node).getPropertyValue('--office-ring').trim(),
	}));
	expect(dialogTheme.border).not.toBe('');
	expect(dialogTheme.ring).not.toBe('');
});
