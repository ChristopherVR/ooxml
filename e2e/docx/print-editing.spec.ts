import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';
const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function fixture() {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t>Editable print text</w:t><w:br/><w:t>Second line</w:t><w:br/><w:t>Third line</w:t></w:r></w:p><w:p><w:r><w:t>Next paragraph</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="4500" w:h="3000"/><w:pgMar w:top="300" w:right="300" w:bottom="300" w:left="300"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}
async function open(page: import('@playwright/test').Page, framework = 'vanilla') {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto(`/?framework=${framework}`);
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'print-edit.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await fixture(),
	});
	const editor = page.locator('docx-editor');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	const layout = editor.getByRole('combobox', { name: 'Layout view', exact: true });
	await layout.selectOption('print');
	const lines = editor.locator('.dve-print-column .dve-print-line');
	await expect(lines.first()).toContainText('Editable');
	return { editor, layout, lines };
}
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: native editing remains in Print Layout with undo and save/reopen`, async ({
		page,
	}) => {
		const { editor, layout, lines } = await open(page, framework);
		await lines
			.first()
			.locator('span')
			.first()
			.click({ position: { x: 0, y: 5 } });
		await page.keyboard.type('Changed ');
		await expect(layout).toHaveValue('print');
		await expect(editor.locator('.dve-print-pages').first()).toContainText('Changed Editable');
		await expect(editor.locator('.dve-print-caret')).toHaveCount(1);
		await page.keyboard.press('Shift+ArrowRight');
		await expect(editor.locator('.dve-print-selection')).not.toHaveCount(0);
		await page.keyboard.press('ArrowLeft');
		await page.keyboard.press('Control+z');
		await expect(editor.locator('.dve-print-pages').first()).not.toContainText('Changed');
		await page.keyboard.type('Saved ');
		await expect(editor.locator('.dve-print-pages').first()).toContainText('Saved Editable');
		const bytes = await editor.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'reopened.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(bytes),
		});
		await layout.selectOption('print');
		await expect(editor.locator('.dve-print-pages').first()).toContainText('Saved Editable');
	});
}
test('Print Layout supports visual line navigation, drag selection and native IME composition', async ({
	page,
}) => {
	const { editor, layout, lines } = await open(page);
	await lines
		.first()
		.locator('span')
		.first()
		.click({ position: { x: 0, y: 5 } });
	await page.keyboard.press('ArrowDown');
	await page.keyboard.type('On second ');
	await expect(lines.nth(1)).toContainText('On second Second');
	await page.keyboard.press('Home');
	const session = await page.context().newCDPSession(page);
	await session.send('Input.imeSetComposition', {
		text: '日本',
		selectionStart: 2,
		selectionEnd: 2,
	});
	await expect(editor.locator('.dve-print-pages').first()).toContainText('日本');
	await session.send('Input.insertText', { text: '日本語' });
	await expect(editor.locator('.dve-print-pages').first()).toContainText('日本語On second');
	await expect(layout).toHaveValue('print');
	await page.keyboard.press('Shift+End');
	await expect(editor.locator('.dve-print-selection')).not.toHaveCount(0);
	const first = await lines.first().boundingBox(),
		last = await lines.nth(2).boundingBox();
	if (!first || !last) throw new Error('Missing printed line geometry');
	await page.mouse.move(first.x + 1, first.y + first.height / 2);
	await page.mouse.down();
	await page.mouse.move(last.x + 30, last.y + last.height / 2, { steps: 6 });
	await page.mouse.up();
	await expect(editor.locator('.dve-print-selection')).not.toHaveCount(0);
	await page.keyboard.type('Replacement');
	await expect(editor.locator('.dve-print-pages').first()).toContainText('Replacement');
	await expect(editor.locator('.dve-print-pages').first()).not.toContainText('Editable');
	await session.detach();
});

test('Print Layout preserves native paste and prevents edits in read-only mode', async ({
	page,
}) => {
	const { editor, layout, lines } = await open(page);
	await lines
		.first()
		.locator('span')
		.first()
		.click({ position: { x: 0, y: 5 } });
	await editor
		.locator('.ProseMirror')
		.first()
		.evaluate((node) => {
			const data = new DataTransfer();
			data.setData('text/plain', 'Pasted ');
			node.dispatchEvent(
				new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
			);
		});
	await expect(editor.locator('.dve-print-pages').first()).toContainText('Pasted Editable');
	await expect(layout).toHaveValue('print');
	await lines.first().locator('span').first().dblclick();
	await page.keyboard.type('Copied');
	await expect(editor.locator('.dve-print-pages').first()).toContainText('Copied Editable');
	await editor.evaluate((node) => ((node as unknown as { readOnly: boolean }).readOnly = true));
	await lines
		.first()
		.locator('span')
		.first()
		.click({ position: { x: 0, y: 5 } });
	await page.keyboard.type('Forbidden');
	await expect(editor.locator('.dve-print-pages').first()).not.toContainText('Forbidden');
	await expect(editor.locator('.dve-print-pages').first()).toContainText('Copied Editable');
});
for (const fontFamily of [undefined, 'Courier New'])
	test(`Print Layout edits across column/page boundaries at 150 percent zoom${fontFamily ? ' after wrapped insertion' : ''}`, async ({
		page,
	}) => {
		const { editor, layout } = await open(page);
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${w}"><w:body>${Array.from({ length: 12 }, (_, index) => `<w:p><w:r>${fontFamily ? `<w:rPr><w:rFonts w:ascii="${fontFamily}" w:hAnsi="${fontFamily}"/></w:rPr>` : ''}<w:t>Paragraph ${index} content words</w:t></w:r></w:p>`).join('')}<w:sectPr><w:pgSz w:w="4500" w:h="2250"/><w:pgMar w:top="300" w:right="300" w:bottom="300" w:left="300"/><w:cols w:num="2" w:space="300"/></w:sectPr></w:body></w:document>`,
		);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'columns.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: await zip.generateAsync({ type: 'nodebuffer' }),
		});
		await layout.selectOption('print');
		await editor.getByRole('combobox', { name: 'Zoom', exact: true }).selectOption('150');
		const sheets = editor.locator('.dve-print-page');
		await expect(sheets).not.toHaveCount(1);
		const line = sheets.nth(1).locator('.dve-print-column .dve-print-line').first();
		await line
			.locator('span')
			.first()
			.click({ position: { x: 0, y: 5 } });
		const selection = () =>
			editor.evaluate((node) => {
				const view = (
					node as unknown as {
						core: {
							view: {
								state: {
									selection: {
										head: number;
										$head: { parentOffset: number; parent: { attrs: { id: string } } };
									};
								};
							};
						};
					}
				).core.view;
				const { head, $head } = view.state.selection;
				return { head, id: String($head.parent.attrs.id), offset: $head.parentOffset };
			});
		const before = await selection();
		expect(before.id).toBe(await line.getAttribute('data-block-id'));
		await page.keyboard.type('Zoomed ');
		await expect(sheets.nth(1)).toContainText('Zoomed');
		await expect(layout).toHaveValue('print');
		const after = await selection();
		expect(after).toEqual({ head: before.head + 7, id: before.id, offset: before.offset + 7 });
		const sourceLines = editor.locator(
			`.dve-print-column .dve-print-line[data-block-id="${after.id}"]`,
		);
		// Insertion can rewrap and repaginate a paragraph. Its source character at the caret,
		// rather than the original page's first line, identifies the correct projected line.
		const currentIndex = await sourceLines.evaluateAll(
			(lines, offset) =>
				lines.findIndex(
					(line) =>
						Number((line as HTMLElement).dataset.sourceStart) <= offset &&
						Number((line as HTMLElement).dataset.sourceEnd) > offset,
				),
			after.offset,
		);
		expect(currentIndex).toBeGreaterThanOrEqual(0);
		const currentLine = sourceLines.nth(currentIndex);
		await expect(editor.locator('.dve-print-caret')).toHaveCount(1);
		await expect
			.poll(async () => {
				const caret = await editor.locator('.dve-print-caret').boundingBox(),
					lineRect = await currentLine.boundingBox();
				return caret && lineRect ? Math.abs(caret.y - lineRect.y) : Infinity;
			})
			.toBeLessThan(2);
		const caret = await editor.locator('.dve-print-caret').boundingBox(),
			lineRect = await currentLine.boundingBox();
		if (!caret || !lineRect) throw new Error('Missing caret geometry');
		expect(caret.x).toBeGreaterThanOrEqual(lineRect.x);
		expect(caret.x).toBeLessThan(lineRect.x + lineRect.width);
		if (fontFamily) {
			await expect(sourceLines.first()).toHaveText('Zoomed ');
			await expect(currentLine).toContainText('Paragraph');
			const firstRect = await sourceLines.first().boundingBox();
			expect(caret.y - firstRect!.y).toBeGreaterThan(20);
		}
		await expect
			.poll(async () =>
				editor.evaluate((node) => {
					const view = (
						node as unknown as {
							core: {
								view: {
									state: { selection: { head: number } };
									coordsAtPos(pos: number): { left: number; top: number };
								};
							};
						}
					).core.view;
					const printed = node
						.shadowRoot!.querySelector('.dve-print-caret')!
						.getBoundingClientRect();
					const native = view.coordsAtPos(view.state.selection.head);
					return Math.max(Math.abs(native.left - printed.left), Math.abs(native.top - printed.top));
				}),
			)
			.toBeLessThan(2);
		await page.keyboard.press('Control+End');
		await expect(editor.locator('.dve-print-caret')).toBeInViewport();
	});

test('Print Layout disposes an active drag and reconnects its native input bridge', async ({
	page,
}) => {
	const { editor, layout, lines } = await open(page);
	const rect = await lines.first().boundingBox();
	if (!rect) throw new Error('Missing print geometry');
	await page.mouse.move(rect.x + 1, rect.y + rect.height / 2);
	await page.mouse.down();
	await editor.evaluate((node) => {
		const parent = node.parentNode!;
		node.remove();
		parent.appendChild(node);
	});
	await page.mouse.move(rect.x + 50, rect.y + 40);
	await page.mouse.up();
	await expect(layout).toHaveValue('print');
	await lines
		.first()
		.locator('span')
		.first()
		.click({ position: { x: 0, y: 5 } });
	await page.keyboard.type('Reconnected ');
	await expect(editor.locator('.dve-print-pages').first()).toContainText('Reconnected Editable');
	await expect(editor.locator('.dve-print-caret')).toHaveCount(1);
	await editor.evaluate((node) => {
		const parent = node.parentNode!;
		node.remove();
		parent.appendChild(node);
	});
	await lines
		.first()
		.locator('span')
		.first()
		.click({ position: { x: 0, y: 5 } });
	await page.keyboard.type('Again ');
	await expect(editor.locator('.dve-print-pages').first()).toContainText(
		'Again Reconnected Editable',
	);
	await expect(editor.locator('.dve-print-pages').first()).not.toContainText('Again Again');
});
