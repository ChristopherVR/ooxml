import { expect, test, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';

test.use({ viewport: { width: 1366, height: 1400 } });

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const png = Buffer.from(
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
	'base64',
);
const EMU = 9525;

const anchor = (id: number, behind: boolean, offsetX: number, offsetY: number, name: string) =>
	`<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id}" behindDoc="${behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>${offsetX * EMU}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${offsetY * EMU}</wp:posOffset></wp:positionV><wp:extent cx="${200 * EMU}" cy="${100 * EMU}"/>${behind ? '<wp:wrapNone/>' : '<wp:wrapSquare wrapText="bothSides"/>'}<wp:docPr id="${id}" name="${name}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${200 * EMU}" cy="${100 * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

async function fixture(): Promise<Buffer> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	zip.file(
		'word/_rels/document.xml.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
	);
	zip.file('word/media/image1.png', png);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<w:p><w:r><w:t>Intro line</w:t></w:r></w:p>
<w:p>${anchor(1, false, 300, 0, 'Front')}<w:r><w:t>Text beside the front picture</w:t></w:r></w:p>
<w:p><w:r><w:t>Spacer</w:t></w:r></w:p><w:p><w:r><w:t>Spacer</w:t></w:r></w:p><w:p><w:r><w:t>Spacer</w:t></w:r></w:p><w:p><w:r><w:t>Spacer</w:t></w:r></w:p><w:p><w:r><w:t>Spacer</w:t></w:r></w:p>
<w:p>${anchor(2, true, 0, 0, 'Behind')}<w:r><w:t>Text drawn over the back picture</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page): Promise<Locator> {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'floats.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await fixture(),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror').first()).toContainText('Intro line');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await expect(editor.locator('.dve-print-float')).toHaveCount(2);
	return editor;
}

const front = (editor: Locator) =>
	editor.locator('.dve-print-float[data-block-id]:not(.dve-print-float-behind)');
const frame = (editor: Locator) => editor.locator('.dve-print-float-frame');

/**
 * The element's box. A relayout swaps the page DOM in one step, so a locator resolved just before
 * it can find its element detached; read again until the new element answers.
 */
async function rect(locator: Locator) {
	let box = await locator.boundingBox();
	for (let attempt = 0; !box && attempt < 20; attempt++) {
		await locator.page().waitForTimeout(25);
		box = await locator.boundingBox();
	}
	return box!;
}

async function savedXml(editor: Locator): Promise<string> {
	const bytes = Buffer.from(
		await editor.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		),
	);
	return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
}

const offsets = (xml: string, name: string) => {
	const at = xml.indexOf(`name="${name}"`);
	const drawing = xml.slice(Math.max(0, at - 1200), at);
	const x = /<wp:positionH[^>]*><wp:posOffset>(-?\d+)</.exec(drawing)?.[1];
	const y = /<wp:positionV[^>]*><wp:posOffset>(-?\d+)</.exec(drawing)?.[1];
	return { x: Number(x) / EMU, y: Number(y) / EMU };
};

test('clicking a floating picture selects it with a frame and eight handles', async ({ page }) => {
	const editor = await open(page);
	await expect(frame(editor)).toHaveCount(0);
	const box = await rect(front(editor));
	await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
	await expect(frame(editor)).toHaveCount(1);
	await expect(frame(editor).locator('[data-float-handle]')).toHaveCount(8);
	const framed = await rect(frame(editor));
	expect(Math.abs(framed.x - box.x)).toBeLessThan(2);
	expect(Math.abs(framed.width - box.width)).toBeLessThan(2);
	// A press on text elsewhere drops the selection and places the caret.
	const intro = editor.locator('.dve-print-line', { hasText: 'Intro line' }).first();
	const line = await rect(intro);
	await page.mouse.click(line.x + 2, line.y + line.height / 2);
	await expect(frame(editor)).toHaveCount(0);
	await page.keyboard.press('Home');
	await page.keyboard.type('Q');
	await expect(editor.locator('.ProseMirror').first()).toContainText('QIntro line');
});

test('Delete removes the selected picture in one undo step', async ({ page }) => {
	const editor = await open(page);
	const box = await rect(front(editor));
	await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
	await expect(frame(editor)).toHaveCount(1);
	await page.keyboard.press('Delete');
	await expect(editor.locator('.dve-print-float')).toHaveCount(1);
	await expect(editor.locator('.ProseMirror').first().locator('img[data-docx-image]')).toHaveCount(
		1,
	);
	await page.keyboard.press('Control+z');
	await expect(editor.locator('.dve-print-float')).toHaveCount(2);
	await expect(editor.locator('.ProseMirror').first().locator('img[data-docx-image]')).toHaveCount(
		2,
	);
});

test('arrow keys nudge the picture and Ctrl+arrow moves one pixel', async ({ page }) => {
	const editor = await open(page);
	const before = await rect(front(editor));
	await page.mouse.click(before.x + before.width / 2, before.y + before.height / 2);
	await page.keyboard.press('ArrowRight');
	await page.keyboard.press('ArrowDown');
	await expect.poll(async () => Math.round((await rect(front(editor))).x - before.x)).toBe(6);
	expect(Math.round((await rect(front(editor))).y - before.y)).toBe(6);
	await page.keyboard.press('Control+ArrowLeft');
	await expect.poll(async () => Math.round((await rect(front(editor))).x - before.x)).toBe(5);
	await expect(frame(editor)).toHaveCount(1);
	const moved = offsets(await savedXml(editor), 'Front');
	expect(moved).toEqual({ x: 305, y: 6 });
	// The caret never moved with the arrows: typing replaces nothing and lands in text after Escape.
	await page.keyboard.press('Escape');
	await expect(frame(editor)).toHaveCount(0);
});

test('dragging moves the picture by its anchor offset in one undo step', async ({ page }) => {
	const editor = await open(page);
	const before = await rect(front(editor));
	const startX = before.x + before.width / 2;
	const startY = before.y + before.height / 2;
	await page.mouse.move(startX, startY);
	await page.mouse.down();
	await page.mouse.move(startX + 20, startY + 10, { steps: 4 });
	await page.mouse.move(startX + 60, startY + 40, { steps: 6 });
	await page.mouse.up();
	await expect.poll(async () => Math.round((await rect(front(editor))).x - before.x)).toBe(60);
	expect(Math.round((await rect(front(editor))).y - before.y)).toBe(40);
	expect(offsets(await savedXml(editor), 'Front')).toEqual({ x: 360, y: 40 });
	await page.keyboard.press('Control+z');
	await expect.poll(async () => Math.round((await rect(front(editor))).x - before.x)).toBe(0);
	expect(offsets(await savedXml(editor), 'Front')).toEqual({ x: 300, y: 0 });
});

test('a corner handle resizes with the aspect ratio kept', async ({ page }) => {
	const editor = await open(page);
	const before = await rect(front(editor));
	await page.mouse.click(before.x + before.width / 2, before.y + before.height / 2);
	const handle = await rect(frame(editor).locator('[data-float-handle="se"]'));
	const hx = handle.x + handle.width / 2;
	const hy = handle.y + handle.height / 2;
	await page.mouse.move(hx, hy);
	await page.mouse.down();
	await page.mouse.move(hx + 40, hy + 5, { steps: 5 });
	await page.mouse.move(hx + 100, hy + 10, { steps: 5 });
	await page.mouse.up();
	await expect.poll(async () => Math.round((await rect(front(editor))).width)).toBeGreaterThan(240);
	const after = await rect(front(editor));
	expect(after.width / after.height).toBeCloseTo(2, 1);
	expect(Math.abs(after.x - before.x)).toBeLessThan(1);
	const xml = await savedXml(editor);
	const extent = /<wp:extent cx="(\d+)" cy="(\d+)"\/>/.exec(
		xml.slice(0, xml.indexOf('name="Front"') + 10),
	);
	expect(Number(extent![1]) / Number(extent![2])).toBeCloseTo(2, 1);
});

test('a picture behind the text does not steal presses on its text', async ({ page }) => {
	const editor = await open(page);
	const line = editor.locator('.dve-print-line', { hasText: 'drawn over the back' }).first();
	const box = await rect(line);
	await page.mouse.click(box.x + 4, box.y + box.height / 2);
	await expect(frame(editor)).toHaveCount(0);
	await page.keyboard.press('Home');
	await page.keyboard.type('Z');
	await expect(editor.locator('.ProseMirror').first()).toContainText('ZText drawn over');
});
