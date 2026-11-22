import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { newDocument, reopenDocument, reveal, saveButton, dialogByHeading } from './helpers';

test('Insert > Text Box inserts, edits and exports an inline text box', async ({ page }) => {
	await page.setViewportSize({ width: 1700, height: 720 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('Before ');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const insertTextBox = editor.getByRole('button', { name: 'Text box', exact: true });
	await reveal(editor, insertTextBox);
	await insertTextBox.click();
	const dialog = dialogByHeading(editor, 'Text box');
	await expect(dialog).toBeVisible();
	await dialog.getByRole('textbox', { name: 'Text', exact: true }).fill('Hello box\nSecond line');
	await dialog.getByRole('spinbutton', { name: 'Width (inches)', exact: true }).fill('0');
	await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeDisabled();
	await dialog.getByRole('spinbutton', { name: 'Width (inches)', exact: true }).fill('2');
	expect(
		await dialog
			.getByRole('button', { name: 'OK', exact: true })
			.evaluate((button) => button.getBoundingClientRect().bottom <= window.innerHeight),
	).toBe(true);
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	// Drawn like any shape: its outline, and its lines as paragraphs over it.
	const box = editor.locator('.ProseMirror .dve-shape');
	await expect(box).toHaveCount(1);
	await expect(box.locator('.dve-text-box-content p')).toHaveText(['Hello box', 'Second line']);
	await expect(box.locator('svg path')).toHaveAttribute('stroke', /^#000000$/i);
	await expect(editor.locator('.ProseMirror .dve-image-placeholder')).toHaveCount(0);
	// Selecting the box shows Shape Format, as in Word; Insert > Text Box then edits the box
	// instead of inserting another.
	await box.click();
	await expect(editor.getByRole('tab', { name: 'Shape Format', exact: true })).toBeVisible();
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await reveal(editor, insertTextBox);
	await insertTextBox.click();
	await expect(dialog.getByRole('textbox', { name: 'Text', exact: true })).toHaveValue(
		'Hello box\nSecond line',
	);
	await dialog.getByRole('textbox', { name: 'Text', exact: true }).fill('Edited');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(box).toHaveCount(1);
	await expect(box.locator('.dve-text-box-content p')).toHaveText(['Edited']);
	// Shape Format > Add Text edits the same box in the Text Box dialog.
	await editor.getByRole('tab', { name: 'Shape Format', exact: true }).click();
	await editor.getByRole('button', { name: 'Add text', exact: true }).click();
	await expect(dialog.getByRole('textbox', { name: 'Text', exact: true })).toHaveValue('Edited');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	const downloadPromise = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await downloadPromise).path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toContain('txBox="1"');
	expect(xml).toContain('<w:t>Edited</w:t>');
	expect(xml.match(/<wp:inline/g)).toHaveLength(1);
});

test('WordArt stays styled and editable after reopening a DOCX', async ({ page }, testInfo) => {
	await page.setViewportSize({ width: 1700, height: 800 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const insertWordArt = editor.getByRole('button', { name: 'WordArt', exact: true });
	await reveal(editor, insertWordArt);
	await insertWordArt.click();
	const picker = editor.locator('office-ui-dialog[open]');
	await picker.getByRole('textbox', { name: 'WordArt text' }).fill('Annual review');
	await picker.getByRole('button', { name: 'WordArt style 1', exact: true }).press('End');
	await expect(
		picker.getByRole('button', { name: 'WordArt style 36', pressed: true }),
	).toBeFocused();
	await picker.getByRole('button', { name: 'WordArt style 30', exact: true }).click();
	await picker.getByRole('combobox', { name: 'Transform' }).selectOption('textArchUp');
	await expect(picker.locator('[data-word-art-transform="textArchUp"] textPath')).toHaveText(
		'Annual review',
	);
	await picker.getByRole('button', { name: 'Insert', exact: true }).click();
	const box = editor.locator('.ProseMirror .dve-shape');
	await expect(box).toContainText('Annual review');
	await expect(box.locator('[data-word-art-transform="textArchUp"] textPath')).toHaveText(
		'Annual review',
	);
	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const path = testInfo.outputPath('word-art.docx');
	await (await pending).saveAs(path);
	const artifact = await readFile(path);
	const zip = await JSZip.loadAsync(artifact);
	const xml = await zip.file('word/document.xml')?.async('string');
	expect(xml).toContain('w14:textOutline');
	expect(xml).toContain('w14:shadow');
	expect(xml).toContain('prst="textArchUp"');
	await reopenDocument(page, 'word-art.docx', artifact);
	await box.click();
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await reveal(editor, insertWordArt);
	await insertWordArt.click();
	await expect(picker.getByRole('combobox', { name: 'Transform' })).toHaveValue('textArchUp');
	await picker.getByRole('combobox', { name: 'Transform' }).selectOption('textWave1');
	await picker.getByRole('button', { name: 'OK', exact: true }).click();
	await editor.getByRole('tab', { name: 'Shape Format', exact: true }).click();
	await editor.getByRole('button', { name: 'Add text', exact: true }).click();
	const textBox = dialogByHeading(editor, 'Text box');
	await textBox.getByRole('textbox', { name: 'Text', exact: true }).fill('Annual results');
	await textBox.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(box).toContainText('Annual results');
	await expect(box.locator('[data-word-art-transform="textWave1"] textPath')).toHaveText(
		'Annual results',
	);
});

test.describe('phone Insert ribbon', () => {
	test.use({ viewport: { width: 375, height: 812 }, hasTouch: true, isMobile: true });

	test('inserts a symbol at the text caret through the folded ribbon', async ({ page }) => {
		await page.goto('/?framework=vanilla');
		await newDocument(page);
		const editor = page.locator('docx-editor');
		const body = editor.locator('.ProseMirror');
		await body.tap();
		await page.keyboard.type('abcd');
		await page.keyboard.press('ArrowLeft');
		await page.keyboard.press('ArrowLeft');

		await editor.getByRole('tab', { name: 'Insert', exact: true }).tap();
		const symbol = editor.getByRole('combobox', { name: 'Symbol', exact: true });
		await reveal(editor, symbol);
		await symbol.selectOption('©');
		await expect(body).toHaveText('ab©cd');
		await reveal(editor, symbol);
		await symbol.selectOption('more');
		const gallery = dialogByHeading(editor, 'Symbol');
		await gallery.getByRole('combobox', { name: 'Font', exact: true }).selectOption('Wingdings 3');
		await gallery.getByRole('button', { name: /U\+2B60/ }).tap();
		await expect(body).toHaveText('ab©⭠cd');
	});

	test('inserts WordArt and an icon through folded groups', async ({ page }) => {
		await page.goto('/?framework=vanilla');
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.getByRole('tab', { name: 'Insert', exact: true }).tap();
		const wordArt = editor.getByRole('button', { name: 'WordArt', exact: true });
		await reveal(editor, wordArt);
		await wordArt.tap();
		const artPicker = dialogByHeading(editor, 'WordArt');
		await artPicker.getByRole('textbox', { name: 'WordArt text' }).fill('Phone WordArt');
		await artPicker.getByRole('button', { name: 'WordArt style 31', exact: true }).tap();
		await artPicker.getByRole('button', { name: 'Insert', exact: true }).tap();
		await expect(editor.locator('.ProseMirror .dve-shape')).toContainText('Phone WordArt');
		await expect(editor.locator('.ProseMirror .dve-shape linearGradient')).toHaveCount(1);

		const icons = editor.getByRole('button', { name: 'Icons', exact: true });
		await reveal(editor, icons);
		await icons.tap();
		const iconPicker = dialogByHeading(editor, 'Icons');
		await iconPicker.getByRole('searchbox', { name: 'Search icons' }).fill('Light bulb');
		await iconPicker.getByRole('button', { name: 'Light bulb', exact: true }).tap();
		await iconPicker.getByRole('searchbox', { name: 'Search icons' }).fill('Check mark');
		await iconPicker.getByRole('button', { name: 'Check mark', exact: true }).tap();
		await iconPicker.getByRole('button', { name: 'Insert', exact: true }).tap();
		await expect(editor.locator('.ProseMirror img[alt="Light bulb"]')).toBeVisible();
		await expect(editor.locator('.ProseMirror img[alt="Check mark"]')).toBeVisible();
	});
});

test('inserts a captured screenshot as a native Word picture', async ({ page }) => {
	await page.addInitScript(() => {
		Object.defineProperty(navigator.mediaDevices, 'getDisplayMedia', {
			configurable: true,
			value: () => {
				const canvas = document.createElement('canvas');
				canvas.width = 320;
				canvas.height = 180;
				const paint = canvas.getContext('2d')!;
				const stream = canvas.captureStream(20);
				let frame = 0;
				const timer = setInterval(() => {
					if (stream.getVideoTracks()[0]?.readyState === 'ended') clearInterval(timer);
					else {
						paint.fillStyle = frame++ % 2 ? '#0451a5' : '#00aa77';
						paint.fillRect(0, 0, 320, 180);
					}
				}, 50);
				return Promise.resolve(stream);
			},
		});
	});
	await page.setViewportSize({ width: 1700, height: 800 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('Before screenshot');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const menu = editor.getByRole('combobox', { name: 'Screenshot', exact: true });
	await reveal(editor, menu);
	await menu.selectOption('full');
	await expect(editor.getByRole('img', { name: 'Screenshot', exact: true })).toBeVisible();
	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile(await (await pending).path()));
	expect(Object.keys(zip.files)).toEqual(
		expect.arrayContaining([expect.stringMatching(/^word[/]media[/].+[.]png$/u)]),
	);
});
