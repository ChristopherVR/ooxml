import { expect, test, type Page } from '@playwright/test';
import { dialogByHeading, fileInput, newDocument, openSample } from './helpers';
import { multiPageDocx } from './print-click-fixture';

test.use({ viewport: { width: 1366, height: 1500 } });

/** A click anywhere on a Print Layout page puts a visible caret at the nearest text, like Word. */
async function enterPrint(page: Page) {
	const editor = page.locator('docx-editor');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('combobox', { name: 'Layout view', exact: true }).selectOption('print');
	await expect(editor.locator('.dve-print-line').first()).toBeVisible();
	return editor;
}

async function clickAndType(page: Page, x: number, y: number, text: string) {
	await page.mouse.click(x, y);
	await expect(page.locator('docx-editor .dve-print-caret')).toHaveCount(1);
	await expect
		.poll(
			async () => (await page.locator('docx-editor .dve-print-caret').boundingBox())?.height ?? 0,
		)
		.toBeGreaterThan(2);
	await page.keyboard.type(text);
}

test('blank page: clicks in the margins and empty space place a caret and take typing', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = await enterPrint(page);
	const sheet = (await editor.locator('.dve-print-page').first().boundingBox())!;
	const spots: Array<[string, number, number]> = [
		['bottom blank area', sheet.x + sheet.width / 2, sheet.y + sheet.height * 0.6],
		['right margin', sheet.x + sheet.width - 8, sheet.y + 200],
		['left margin', sheet.x + 8, sheet.y + 300],
		['bottom margin', sheet.x + sheet.width / 2, sheet.y + sheet.height - 10],
		['top margin', sheet.x + sheet.width / 2, sheet.y + 8],
	];
	for (const [label, x, y] of spots) {
		await clickAndType(page, x, y, '.');
		await expect(editor.locator('.ProseMirror').first(), label).toContainText('.');
		await page.keyboard.press('Control+z');
	}
});

test('sample document: clicks right of a line and below the text land on the nearest text', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await openSample(page);
	const editor = await enterPrint(page);
	const line = editor.locator('.dve-print-line').first();
	const box = (await line.boundingBox())!;
	const sheet = (await editor.locator('.dve-print-page').first().boundingBox())!;
	// Right of the end of the first line, inside the page margin.
	await clickAndType(page, sheet.x + sheet.width - 6, box.y + box.height / 2, 'Z');
	await expect(editor.locator('.ProseMirror p').first()).toContainText('Z');
	await page.keyboard.press('Control+z');
	// Immediately after typing (relayout pending), then a quick double click in the bottom margin.
	await page.keyboard.type('Q');
	await page.mouse.click(sheet.x + sheet.width / 2, sheet.y + sheet.height - 12);
	await page.mouse.click(sheet.x + sheet.width / 2, sheet.y + sheet.height - 12);
	await expect(editor.locator('.dve-print-caret')).toHaveCount(1);
});

async function openMultiPage(page: Page, view: 'print' | 'draft' = 'print') {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'multi.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await multiPageDocx(),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror').first()).toContainText('Page one intro');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor
		.getByRole('combobox', { name: 'Layout view', exact: true })
		.selectOption(view, { force: true });
	if (view === 'print') await expect(editor.locator('.dve-print-line').first()).toBeVisible();
	return editor;
}

/** Typing after the click lands in the document, proving the editor owns focus. */
async function typesIntoEditor(page: Page, editor: ReturnType<Page['locator']>, mark: string) {
	await page.keyboard.type(mark);
	await expect(editor.locator('.ProseMirror').first()).toContainText(mark);
	await page.keyboard.press('Control+z');
}

test('multi-page: table cells, header area and later pages show a caret', async ({ page }) => {
	const editor = await openMultiPage(page);
	const cell = (await editor.locator('.dve-print-line', { hasText: 'Cell A' }).boundingBox())!;
	for (const [dx, dy] of [
		[5, 5],
		[150, 80],
	] as const) {
		await clickAndType(page, cell.x + dx, cell.y + dy, '#');
		await expect(editor.locator('.ProseMirror table').first()).toContainText('#');
		await page.keyboard.press('Control+z');
	}
	const sheet = (await editor.locator('.dve-print-page').first().boundingBox())!;
	// The header area keeps a visible body caret (double-click edits the header itself).
	await clickAndType(page, sheet.x + 400, sheet.y + 30, '#');
	await page.keyboard.press('Control+z');
	await page.mouse.move(600, 400);
	await page.mouse.wheel(0, 2300);
	const thirdPage = editor.locator('.dve-print-page').nth(2);
	await expect
		.poll(async () => {
			const box = await thirdPage.boundingBox().catch(() => null);
			return box ? box.y < 900 && box.y + box.height > 400 : false;
		})
		.toBe(true);
	const third = (await thirdPage.boundingBox())!;
	await clickAndType(page, third.x + 20, Math.max(third.y, 0) + 300, '#');
	await expect(editor.locator('.ProseMirror').first()).toContainText('#Page three first');
});

test('multi-page: a ribbon command keeps the caret, and one click back shows it', async ({
	page,
}) => {
	const editor = await openMultiPage(page);
	const sheet = (await editor.locator('.dve-print-page').first().boundingBox())!;
	await clickAndType(page, sheet.x + 300, sheet.y + 120, '');
	await editor.getByRole('tab', { name: 'Home', exact: true }).click();
	for (const name of ['Bold', 'Align center', 'Grow font']) {
		await editor.getByRole('button', { name, exact: true }).first().click();
		await expect(editor.locator('.dve-print-caret')).toHaveCount(1);
		await typesIntoEditor(page, editor, '#');
	}
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await page.mouse.click(sheet.x + 300, sheet.y + 300);
	await expect(editor.locator('.dve-print-caret')).toHaveCount(1);
	await typesIntoEditor(page, editor, '#');
});

test('multi-page: Draft view clicks beside the paper and below the text take typing', async ({
	page,
}) => {
	const editor = await openMultiPage(page, 'draft');
	const paper = (await editor.locator('.ProseMirror').first().boundingBox())!;
	for (const [x, y] of [
		[paper.x - 20, paper.y + 40],
		[20, paper.y + 100],
		[paper.x + 100, Math.min(paper.y + paper.height + 30, 1400)],
	] as const) {
		await page.mouse.click(x, y);
		await typesIntoEditor(page, editor, '#');
	}
});

test.describe('150% zoom at 1366x768', () => {
	test.use({ viewport: { width: 1366, height: 768 } });
	test('clicks beside and below the text still place the caret', async ({ page }) => {
		const editor = await openMultiPage(page);
		await editor.getByRole('button', { name: 'Zoom dialog' }).click();
		const dialog = dialogByHeading(editor, 'Zoom');
		await dialog.getByLabel('Percent', { exact: true }).fill('150');
		await dialog.getByRole('button', { name: 'OK' }).click();
		await expect(dialog).toBeHidden();
		// 816px at 100%: zoom applies once, not once per nested pages container.
		const sheet = (await editor.locator('.dve-print-page').first().boundingBox())!;
		expect(sheet.width).toBeCloseTo(1224, 0);
		const line = (await editor.locator('.dve-print-line').first().boundingBox())!;
		for (const [x, y] of [
			[line.x + 5, line.y + 5],
			[line.x + 300, line.y + 400],
			[8, line.y + 5],
		] as const) {
			await clickAndType(page, x, Math.min(y, 680), '#');
			await page.keyboard.press('Control+z');
		}
	});
});

test.describe('touch at 375x812 with the fit-to-width zoom', () => {
	test.use({ viewport: { width: 375, height: 812 }, hasTouch: true, isMobile: true });
	test('taps on text, below it and in the gutter keep a caret and focus', async ({ page }) => {
		const editor = await openMultiPage(page);
		const sheet = (await editor.locator('.dve-print-page').first().boundingBox())!;
		// The fit-to-width zoom fills the phone width instead of shrinking the page twice.
		expect(sheet.width).toBeGreaterThan(300);
		expect(sheet.width).toBeLessThanOrEqual(375);
		const line = (await editor.locator('.dve-print-line').first().boundingBox())!;
		for (const [x, y] of [
			[line.x + 10, line.y + 5],
			[line.x + 40, line.y + 300],
			[2, line.y + 5],
		] as const) {
			await page.touchscreen.tap(x, y);
			await expect(editor.locator('.dve-print-caret')).toHaveCount(1);
			await typesIntoEditor(page, editor, '#');
		}
	});
});
