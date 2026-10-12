import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';
const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function openLong(page: import('@playwright/test').Page, zoom: string) {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto('/?framework=vanilla');
	const editor = page.locator('docx-editor');
	const zip = new JSZip();
	const blocks = Array.from(
		{ length: 12 },
		(_, index) =>
			`${index ? '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' : ''}<w:p><w:r><w:t>Page ${index} first line</w:t></w:r></w:p>${Array.from({ length: 7 }, (_, line) => `<w:p><w:r><w:t>Page ${index} body line ${line}</w:t></w:r></w:p>`).join('')}`,
	).join('');
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>${blocks}<w:sectPr><w:pgSz w:w="6000" w:h="4500"/><w:pgMar w:top="300" w:right="300" w:bottom="300" w:left="300"/></w:sectPr></w:body></w:document>`,
	);
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'page-navigation.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await zip.generateAsync({ type: 'nodebuffer' }),
	});
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	const layout = editor.getByRole('combobox', { name: 'Layout view', exact: true });
	await layout.selectOption('print');
	await editor.getByRole('combobox', { name: 'Zoom', exact: true }).selectOption(zoom);
	await editor.evaluate((node) => {
		const canvas = node.shadowRoot!.querySelector<HTMLElement>('.dve-canvas')!;
		canvas.style.height = '280px';
		canvas.style.maxHeight = '280px';
		canvas.style.flex = 'none';
	});
	const lines = editor.locator('.dve-print-column .dve-print-line');
	await expect(lines.first()).toContainText('Page 0 first line');
	await lines
		.first()
		.locator('span')
		.first()
		.click({ position: { x: 0, y: 5 } });
	return { editor, layout, lines, canvas: editor.locator('.dve-canvas') };
}
for (const zoom of ['100', '150']) {
	test(`Print Layout uses native screen/page navigation at ${zoom} percent zoom`, async ({
		page,
	}) => {
		const { editor, lines, canvas } = await openLong(page, zoom);
		const head = () =>
			editor.evaluate(
				(node) =>
					(node as unknown as { core: { view: { state: { selection: { head: number } } } } }).core
						.view.state.selection.head,
			);
		const initial = await head();
		await page.keyboard.press('PageDown');
		await expect.poll(head).toBeGreaterThan(initial);
		expect(await canvas.evaluate((node) => node.scrollTop)).toBeGreaterThan(0);
		const screenHead = await head();
		await page.keyboard.press('PageUp');
		await expect.poll(head).toBeLessThan(screenHead);
		await page.keyboard.press('Control+Home');
		await expect.poll(head).toBe(1);
		await page.keyboard.press('Control+PageDown');
		await page.keyboard.type('Next ');
		await expect(editor.locator('.dve-print-page').nth(1)).toContainText('Next Page 1 first line');
		await page.keyboard.press('Control+z');
		await page.keyboard.press('Control+PageUp');
		await page.keyboard.type('Previous ');
		await expect(lines.first()).toContainText('Previous Page 0 first line');
		await page.keyboard.press('Control+Home');
		await page.keyboard.press('Shift+PageDown');
		await expect(editor.locator('.dve-print-selection')).not.toHaveCount(0);
	});
}
test('Print Layout drag scrolling extends selection and stops on release and disposal', async ({
	page,
}) => {
	const { editor, canvas, lines } = await openLong(page, '150');
	const rect = await canvas.boundingBox(),
		line = await lines.first().boundingBox();
	if (!rect || !line) throw new Error('Missing scroll geometry');
	await page.mouse.move(line.x + 2, line.y + line.height / 2);
	await page.mouse.down();
	await page.mouse.move(line.x + 40, rect.y + rect.height + 15);
	const top = () => canvas.evaluate((node) => node.scrollTop);
	await expect.poll(top).toBeGreaterThan(100);
	await expect(editor.locator('.dve-print-selection')).not.toHaveCount(0);
	await page.mouse.up();
	const released = await top();
	await page.waitForTimeout(100);
	expect(await top()).toBe(released);
	const visible = await lines.evaluateAll(
		(nodes, bounds) => {
			const boxes = nodes
				.map((node) => node.getBoundingClientRect())
				.filter((box) => box.height && box.top > bounds.top + 5 && box.bottom < bounds.bottom - 5);
			const box = boxes.at(-1);
			return box ? { x: box.left + 2, y: box.top + box.height / 2 } : undefined;
		},
		{ top: rect.y, bottom: rect.y + rect.height },
	);
	if (!visible) throw new Error('Missing visible body line');
	await page.mouse.move(visible.x, visible.y);
	await page.mouse.down();
	await page.mouse.move(line.x + 40, rect.y + rect.height + 15);
	await expect.poll(top).toBeGreaterThan(released + 50);
	const handle = await editor.elementHandle();
	if (!handle) throw new Error('Missing editor');
	await handle.evaluate((node) => node.remove());
	const stopped = await handle.evaluate(
		(node) => node.shadowRoot!.querySelector('.dve-canvas')!.scrollTop,
	);
	await page.waitForTimeout(100);
	expect(
		await handle.evaluate((node) => node.shadowRoot!.querySelector('.dve-canvas')!.scrollTop),
	).toBe(stopped);
	await page.mouse.up();
});
