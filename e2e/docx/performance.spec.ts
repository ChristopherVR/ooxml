import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';

test('mobile Print Layout bounds page DOM while typing, jumping, scrolling and printing', async ({
	page,
}) => {
	await page.setViewportSize({ width: 390, height: 844 });
	const cdp = await page.context().newCDPSession(page);
	await cdp.send('Emulation.setCPUThrottlingRate', { rate: 4 });
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto('/?framework=vanilla');
	const zip = new JSZip();
	const paragraphs = Array.from(
		{ length: 60 },
		(_, index) =>
			`<w:p><w:pPr>${index ? '<w:pageBreakBefore/>' : ''}</w:pPr><w:r><w:t>Page ${index} editable text</w:t></w:r></w:p>`,
	).join('');
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs}<w:sectPr><w:pgSz w:w="4500" w:h="3000"/><w:pgMar w:top="300" w:right="300" w:bottom="300" w:left="300"/></w:sectPr></w:body></w:document>`,
	);
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'large-mobile.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await zip.generateAsync({ type: 'nodebuffer' }),
	});
	const editor = page.locator('docx-editor');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	const sheets = editor.locator('.dve-print-page');
	const lines = editor.locator('.dve-print-column .dve-print-line');
	await expect(sheets).toHaveCount(60);
	await expect(lines.first()).toContainText('Page 0');
	await expect.poll(() => lines.count()).toBeLessThanOrEqual(6);
	await lines
		.first()
		.locator('span')
		.first()
		.click({ position: { x: 0, y: 5 } });
	await page.keyboard.type('Typed ');
	await expect(sheets.first()).toContainText('Typed Page 0');
	await page.keyboard.press('Control+End');
	await expect(sheets.last().locator('.dve-print-line')).toBeVisible();
	await page.keyboard.type(' at end');
	await expect(sheets.last()).toContainText('Page 59 editable text at end');
	await expect(editor.locator('.dve-print-caret')).toHaveCount(1);
	await sheets.nth(30).scrollIntoViewIfNeeded();
	await expect(sheets.nth(30).locator('.dve-print-line')).toBeVisible();
	await expect.poll(() => lines.count()).toBeLessThanOrEqual(6);
	await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
	await expect(lines).toHaveCount(60);
	await page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
	await expect.poll(() => lines.count()).toBeLessThanOrEqual(6);
	await cdp.send('Emulation.setCPUThrottlingRate', { rate: 1 });
});
