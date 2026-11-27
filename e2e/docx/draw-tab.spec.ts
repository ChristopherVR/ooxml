import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { expect, test, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { savedDocument } from './saved-document';
import { openSample, reopenDocument } from './helpers';

const editor = (page: Page) => page.locator('docx-editor');
const inkFloats = (page: Page) => editor(page).locator('.dve-print-float[data-ink]');

/** Drags across the first page from (x, y), page coordinates, through `points` offsets. */
async function drawOnPage(page: Page, x: number, y: number, points: [number, number][]) {
	const sheet = await editor(page).locator('.dve-print-page').first().boundingBox();
	await page.mouse.move(sheet!.x + x, sheet!.y + y);
	await page.mouse.down();
	for (const [dx, dy] of points)
		await page.mouse.move(sheet!.x + x + dx, sheet!.y + y + dy, { steps: 3 });
	await page.mouse.up();
}

const wave: [number, number][] = Array.from({ length: 16 }, (_, i) => [
	i * 10,
	Math.round(Math.sin(i / 2) * 20),
]);

/** Opens the Draw tab, which builds its panel when it first loads. */
async function openDraw(page: Page) {
	await editor(page).getByRole('tab', { name: 'Draw', exact: true }).click();
	const panel = editor(page).locator('[data-ribbon-tab="draw"]');
	await expect(panel.locator('.ribbon-group')).toHaveCount(4);
	return panel;
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`Draw: a pen stroke becomes Word ink that saves as w14:contentPart (${framework})`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 1600, height: 900 });
		await openSample(page, framework);
		const draw = await openDraw(page);
		// Microsoft 365's groups, in order.
		expect(
			await draw
				.locator('.ribbon-group')
				.evaluateAll((groups) => groups.map((g) => (g as HTMLElement).dataset.label)),
		).toEqual(['Drawing Tools', 'Stencils', 'Convert', 'Replay']);
		await expect(draw.getByRole('button', { name: 'Ink to Math', exact: true })).toBeDisabled();
		await expect(draw.getByRole('button', { name: 'Ink to Math', exact: true })).toHaveAttribute(
			'aria-description',
			/handwriting recognizer/,
		);
		const red = draw.getByRole('button', { name: 'Pen: 0.5 mm, Red' });
		await red.click();
		await expect(red).toHaveAttribute('aria-pressed', 'true');
		await expect(editor(page).locator('office-ui-ink-canvas[tool="pen"]')).toHaveCount(1);
		await drawOnPage(page, 150, 220, wave);
		await expect(inkFloats(page)).toHaveCount(1);
		// A second stroke right away joins the same ink drawing.
		await drawOnPage(page, 150, 250, [
			[60, 0],
			[120, 5],
		]);
		await expect(inkFloats(page)).toHaveCount(1);
		await expect(inkFloats(page).first().locator('path')).toHaveCount(2);
		await expect(editor(page).getByRole('tab', { name: 'Draw', exact: true })).toHaveAttribute(
			'aria-selected',
			'true',
		);

		const { bytes, xml } = await savedDocument(page, `draw-${framework}.docx`);
		expect(xml).toContain('wordprocessingInk');
		expect(xml).toContain('<w14:contentPart');
		const zip = await JSZip.loadAsync(bytes);
		const ink = await zip.file('word/ink/ink1.xml')!.async('string');
		expect(ink).toContain('value="#E81123"');
		expect(ink.match(/<ink:trace /g)).toHaveLength(2);

		// It reopens as ink.
		await reopenDocument(page, `draw-${framework}.docx`, bytes);
		await expect(inkFloats(page)).toHaveCount(1);
	});
}

test('Draw: pen settings, eraser, lasso, Ink to Shape and Ink Replay on Word’s own ink', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1600, height: 900 });
	await openSample(page);
	await reopenDocument(
		page,
		'ink-floating.docx',
		readFileSync(join(import.meta.dirname, 'support', 'ink-floating.docx')),
	);
	const draw = await openDraw(page);
	// Drawing tools work on the Print Layout pages, where Word's floating ink is drawn.
	await draw.getByRole('button', { name: 'Lasso Select', exact: true }).click();
	await draw.getByRole('button', { name: 'Select', exact: true }).click();
	await expect(inkFloats(page)).toHaveCount(1);

	// The pressed pen opens its colour and thickness menu, kept inside the window.
	const blue = draw.getByRole('button', { name: 'Pen: 0.5 mm, Blue' });
	await blue.click();
	await blue.click();
	const menu = draw.getByRole('dialog', { name: 'Pen: 0.5 mm, Blue' });
	await expect(menu).toBeVisible();
	const box = await menu.boundingBox();
	expect(box!.x).toBeGreaterThanOrEqual(0);
	expect(box!.x + box!.width).toBeLessThanOrEqual(1600);
	await menu.getByRole('radio', { name: '2 mm' }).click();
	await expect(draw.getByRole('button', { name: 'Pen: 2 mm, Blue' })).toBeVisible();
	await page.keyboard.press('Escape');

	// Ink Replay hides the ink while it redraws it.
	await draw.getByRole('button', { name: 'Ink Replay', exact: true }).click();
	await expect(editor(page).locator('.dve-ink-replay path').first()).toBeAttached();
	await page.keyboard.press('Escape');
	await expect(editor(page).locator('.dve-ink-replay')).toHaveCount(0);

	// A straight stroke, lassoed, becomes a line shape outlined in the ink's colour.
	await draw.getByRole('button', { name: 'Pen: 2 mm, Blue' }).click();
	await drawOnPage(page, 120, 560, [
		[80, 0],
		[160, 0],
		[240, 0],
	]);
	await expect(inkFloats(page)).toHaveCount(2);
	await page.waitForTimeout(2100);
	await draw.getByRole('button', { name: 'Lasso Select', exact: true }).click();
	await drawOnPage(page, 100, 530, [
		[290, 0],
		[290, 60],
		[0, 60],
		[0, 0],
	]);
	await expect(editor(page).locator('.dve-ink-lasso-box')).toHaveCount(1);
	const toShape = draw.getByRole('button', { name: 'Ink to Shape', exact: true });
	await expect(toShape).toBeEnabled();
	const shapes = await editor(page).locator('.dve-print-float:not([data-ink])').count();
	await toShape.click();
	await expect(inkFloats(page)).toHaveCount(1);
	await expect(editor(page).locator('.dve-print-float:not([data-ink])')).toHaveCount(shapes + 1);

	// The eraser removes what it touches; the last stroke takes the drawing with it.
	await draw.getByRole('button', { name: 'Eraser', exact: true }).click();
	for (let pass = 0; pass < 6 && (await inkFloats(page).count()); pass++) {
		const left = (await inkFloats(page).boundingBox())!;
		await page.mouse.move(left.x - 4, left.y + left.height * (pass / 5));
		await page.mouse.down();
		await page.mouse.move(left.x + left.width + 4, left.y + left.height * (1 - pass / 5), {
			steps: 30,
		});
		await page.mouse.up();
		await page.waitForTimeout(200);
	}
	await expect(inkFloats(page)).toHaveCount(0);
	const { xml } = await savedDocument(page, 'draw-erased.docx');
	expect(xml).not.toContain('<w14:contentPart');
});

test('Draw tab folds at a narrow width and reads in the dark theme', async ({ page }) => {
	await page.setViewportSize({ width: 700, height: 800 });
	await page.emulateMedia({ colorScheme: 'dark' });
	await openSample(page);
	const draw = await openDraw(page);
	const overflow = await draw.evaluate((panel) => panel.scrollWidth <= panel.clientWidth + 1);
	expect(overflow).toBe(true);
	// Each folded group shows its own icon, and its pop-out holds the group's controls.
	const icons = await draw
		.locator('.ribbon-group')
		.evaluateAll((groups) =>
			groups.map(
				(group) =>
					group.getAttribute('icon') ??
					group.querySelector('[data-icon]')?.getAttribute('data-icon'),
			),
		);
	expect(new Set(icons).size).toBe(icons.length);
	const pens: Locator = draw.locator('office-ui-ink-pens');
	await expect(pens).toBeAttached();
});

/** Scrolls the pages down by `top` pixels, so a point low on the first page is in the window. */
const scrollPages = (page: Page, top: number) =>
	editor(page)
		.locator('.dve-canvas')
		.evaluate((canvas, value) => (canvas.scrollTop = value), top);

/** The focused element, through the shadow roots. */
const focused = (page: Page) =>
	page.evaluate(() => {
		let active = document.activeElement;
		while (active?.shadowRoot?.activeElement) active = active.shadowRoot.activeElement;
		return active
			? `${active.tagName.toLowerCase()} ${active.getAttribute('aria-label') ?? ''}`
			: '';
	});

test('Draw: icons, Add Pen and the pen menu work from the keyboard', async ({ page }) => {
	await page.setViewportSize({ width: 1600, height: 900 });
	await openSample(page);
	const draw = await openDraw(page);
	// The built tab draws every icon (a dropped side-effect import once left them blank).
	const blank = await draw
		.locator('button[data-draw]')
		.evaluateAll((buttons) =>
			buttons
				.filter((button) => !button.querySelector('svg path')?.getAttribute('d'))
				.map((button) => button.getAttribute('aria-label')),
		);
	expect(blank).toEqual([]);
	// One caption size across the tab.
	const sizes = await draw
		.locator('button[data-draw] > span')
		.evaluateAll((spans) => [...new Set(spans.map((span) => getComputedStyle(span).fontSize))]);
	expect(sizes).toHaveLength(1);
	// Add Pen is a dropdown command with a menu, like Excel's.
	const addPen = draw.getByRole('button', { name: 'Add Pen', exact: true });
	await expect(addPen).toHaveAttribute('aria-haspopup', 'menu');
	await addPen.click();
	const menu = editor(page).getByRole('menu', { name: 'Add Pen' });
	await expect(menu.getByRole('menuitem')).toHaveCount(3);
	await menu.getByRole('menuitem', { name: 'Pencil' }).click();
	await expect(menu).toHaveCount(0);
	const pens = draw.locator('office-ui-ink-pens').locator('button.pen');
	await expect(pens).toHaveCount(6);
	// The new pen is in use; Escape closes its menu and gives focus back to it, which reopens it.
	const pencil = draw.getByRole('button', { name: 'Pencil: 1 mm, #404040' }).last();
	await expect(pencil).toHaveAttribute('aria-pressed', 'true');
	await pencil.click();
	const dialog = draw.getByRole('dialog', { name: 'Pencil: 1 mm, #404040' });
	await expect(dialog).toBeVisible();
	await page.keyboard.press('Escape');
	await expect(dialog).toHaveCount(0);
	expect(await focused(page)).toBe('button Pencil: 1 mm, #404040');
	await page.keyboard.press('Enter');
	await expect(dialog).toBeVisible();
	await page.keyboard.press('Escape');
	await expect(editor(page).getByRole('tab', { name: 'Draw', exact: true })).toHaveAttribute(
		'aria-selected',
		'true',
	);
});

test('Draw: ink stays where it was drawn, below the text and after a save', async ({ page }) => {
	await page.setViewportSize({ width: 1600, height: 900 });
	await openSample(page);
	const draw = await openDraw(page);
	await draw.getByRole('button', { name: 'Pen: 0.5 mm, Black' }).click();
	await scrollPages(page, 400);
	const sheet = editor(page).locator('.dve-print-page').first();
	const page0 = (await sheet.boundingBox())!;
	// Below the last block of the page (a table): the ink anchors to a body paragraph, not a cell.
	await drawOnPage(page, 200, 900, [
		[60, 10],
		[120, 20],
	]);
	await expect(inkFloats(page)).toHaveCount(1);
	const drawnBox = (await inkFloats(page).boundingBox())!;
	const pageNow = (await sheet.boundingBox())!;
	// Drawing does not scroll the pages, and the ink is where the pen went.
	expect(pageNow.y).toBe(page0.y);
	expect(Math.abs(drawnBox.x - (page0.x + 200))).toBeLessThan(4);
	expect(Math.abs(drawnBox.y - (page0.y + 900))).toBeLessThan(4);
	const offset = async () => {
		const ink = (await inkFloats(page).boundingBox())!;
		const at = (await editor(page).locator('.dve-print-page').first().boundingBox())!;
		return { x: Math.round(ink.x - at.x), y: Math.round(ink.y - at.y) };
	};
	const before = await offset();
	const { bytes } = await savedDocument(page, 'draw-position.docx');
	await reopenDocument(page, 'draw-position.docx', bytes);
	await expect(inkFloats(page)).toHaveCount(1);
	const after = await offset();
	expect(Math.abs(after.x - before.x)).toBeLessThanOrEqual(1);
	expect(Math.abs(after.y - before.y)).toBeLessThanOrEqual(1);
});

test('Draw: Undo from the ribbon after Ink to Shape, and the eraser on ink in line with text', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1600, height: 900 });
	await openSample(page);
	const draw = await openDraw(page);
	await draw.getByRole('button', { name: 'Pen: 0.5 mm, Blue' }).click();
	await scrollPages(page, 400);
	await drawOnPage(page, 120, 760, [
		[80, 0],
		[160, 0],
		[240, 0],
	]);
	await expect(inkFloats(page)).toHaveCount(1);
	const shapes = () => editor(page).locator('.dve-print-float:not([data-ink])').count();
	const before = await shapes();
	await draw.getByRole('button', { name: 'Lasso Select', exact: true }).click();
	await drawOnPage(page, 100, 730, [
		[290, 0],
		[290, 60],
		[0, 60],
		[0, 0],
	]);
	const toShape = draw.getByRole('button', { name: 'Ink to Shape', exact: true });
	await toShape.click();
	await expect(inkFloats(page)).toHaveCount(0);
	expect(await shapes()).toBe(before + 1);
	// Focus stays on the ribbon (Ink to Shape turns itself off): Ctrl+Z undoes it in one step.
	await expect.poll(() => focused(page)).toBe('button Lasso Select');
	await page.keyboard.press('Control+z');
	await expect(inkFloats(page)).toHaveCount(1);
	expect(await shapes()).toBe(before);

	// Word's own inline ink takes the eraser.
	await reopenDocument(
		page,
		'ink-inline.docx',
		readFileSync(join(import.meta.dirname, 'support', 'ink-inline.docx')),
	);
	const inline = editor(page).locator('.dve-print-line [data-ink]');
	await expect(inline).toHaveCount(1);
	const paths = await inline.locator('path').count();
	await draw.getByRole('button', { name: 'Eraser', exact: true }).click();
	await scrollPages(page, 0);
	const box = (await inline.boundingBox())!;
	await page.mouse.move(box.x - 4, box.y + box.height / 2);
	await page.mouse.down();
	await page.mouse.move(box.x + box.width + 4, box.y + box.height / 2, { steps: 30 });
	await page.mouse.up();
	await expect
		.poll(async () => (await inline.count()) && (await inline.locator('path').count()))
		.toBeLessThan(paths);
});
