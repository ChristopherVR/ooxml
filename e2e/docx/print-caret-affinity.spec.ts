import { expect, test, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';

test.use({ viewport: { width: 1366, height: 1500 } });

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const words = (n: number, tag: string) =>
	Array.from({ length: n }, (_, i) => `${tag}${i}`).join(' ');
const para = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

async function fixture(): Promise<Buffer> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	const softBreak =
		'<w:p><w:r><w:t>Before</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>After</w:t></w:r></w:p>';
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>${para(words(60, 'wrap'))}${softBreak}${para(words(1400, 'long'))}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page): Promise<Locator> {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'affinity.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await fixture(),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror').first()).toContainText('wrap0');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await expect(editor.locator('.dve-print-line').first()).toBeVisible();
	return editor;
}

/** The index of the rendered line the caret is drawn on, among the lines matching a selector. */
async function caretLine(editor: Locator, selector = '.dve-print-line'): Promise<number> {
	await expect(editor.locator('.dve-print-caret')).toHaveCount(1);
	await editor.page().evaluate(() => new Promise((r) => requestAnimationFrame(() => r(0))));
	const caret = (await editor.locator('.dve-print-caret').boundingBox())!;
	const lines = editor.locator(selector);
	const count = await lines.count();
	const mid = caret.y + caret.height / 2;
	for (let i = 0; i < count; i++) {
		const box = await lines.nth(i).boundingBox();
		if (box && mid >= box.y && mid <= box.y + box.height) return i;
	}
	return -1;
}

test('a click right of a wrapped line stays at its end; End and Home follow the line', async ({
	page,
}) => {
	const editor = await open(page);
	const first = editor.locator('.dve-print-line').first();
	const box = (await first.boundingBox())!;
	const sheet = (await editor.locator('.dve-print-page').first().boundingBox())!;
	await page.mouse.click(sheet.x + sheet.width - 6, box.y + box.height / 2);
	await expect(editor.locator('.dve-print-caret')).toHaveCount(1);
	// Upstream: the caret is on the first line, not wrapped down to the start of the second.
	expect(await caretLine(editor)).toBe(0);
	await page.keyboard.press('ArrowDown');
	expect(await caretLine(editor)).toBe(1);
	await page.keyboard.press('End');
	expect(await caretLine(editor)).toBe(1);
	const endCaret = (await editor.locator('.dve-print-caret').boundingBox())!;
	await page.keyboard.press('Home');
	expect(await caretLine(editor)).toBe(1);
	const homeCaret = (await editor.locator('.dve-print-caret').boundingBox())!;
	expect(homeCaret.x).toBeLessThan(endCaret.x);
	// End on line 1, then typing lands at the end of that line; the caret stays on the line.
	await page.keyboard.press('ArrowUp');
	await page.keyboard.press('End');
	expect(await caretLine(editor)).toBe(0);
	// A click at the start of the next line is downstream.
	const second = (await editor.locator('.dve-print-line').nth(1).boundingBox())!;
	await page.mouse.click(second.x + 1, second.y + second.height / 2);
	expect(await caretLine(editor)).toBe(1);
});

test('the caret after a soft line break sits at the start of the new line', async ({ page }) => {
	const editor = await open(page);
	const after = editor.locator('.dve-print-line', { hasText: 'After' }).first();
	const before = editor.locator('.dve-print-line', { hasText: 'Before' }).first();
	const box = (await after.boundingBox())!;
	await page.mouse.click(box.x + 1, box.y + box.height / 2);
	const caret = editor.locator('.dve-print-caret');
	await expect.poll(async () => (await caret.boundingBox())?.y).toBeGreaterThanOrEqual(box.y - 1);
	await page.keyboard.press('ArrowLeft');
	const beforeBox = (await before.boundingBox())!;
	// Selection overlays are replaced on an animation frame after keyboard navigation.
	await expect(async () => {
		const moved = await caret.boundingBox();
		expect(moved).not.toBeNull();
		expect(moved!.y).toBeLessThan(box.y);
		expect(moved!.y).toBeGreaterThanOrEqual(beforeBox.y - 1);
	}).toPass({ timeout: 5000 });
});

test('arrow keys carry the caret across a page boundary and back', async ({ page }) => {
	const editor = await open(page);
	const pageOf = () =>
		editor.evaluate(async (el) => {
			await new Promise((r) => requestAnimationFrame(() => r(0)));
			const caret =
				el.querySelector('.dve-print-caret') ?? el.shadowRoot?.querySelector('.dve-print-caret');
			if (!caret) return -2;
			const rect = caret.getBoundingClientRect();
			const root = caret.getRootNode() as Document | ShadowRoot;
			const hit = root
				.elementsFromPoint(rect.left + 1, rect.top + rect.height / 2)
				.find((e) => e.classList.contains('dve-print-page')) as HTMLElement | undefined;
			return hit ? Number(hit.dataset.pageIndex) : -1;
		});
	const long = editor.locator('.dve-print-line', { hasText: 'long0' }).first();
	await long.scrollIntoViewIfNeeded();
	const box = (await long.boundingBox())!;
	await page.mouse.click(box.x + 4, box.y + box.height / 2);
	expect(await pageOf()).toBe(0);
	for (let i = 0; i < 80; i++) await page.keyboard.press('ArrowDown');
	expect(await pageOf()).toBeGreaterThanOrEqual(1);
	for (let i = 0; i < 80; i++) await page.keyboard.press('ArrowUp');
	expect(await pageOf()).toBe(0);
});
