import { expect, test, type Route } from '@playwright/test';
import {
	createEditSession,
	createWorkbook,
	loadXlsx,
	pictureAnchorAt,
	saveXlsx,
} from 'ooxml-core/xlsx';
import { inspectRasterImage } from 'ooxml-core/drawingml';
import { encodePng } from '../../src/core/drawingml/raster-encode';
import { editor, goToCell, openLanding, pageErrors } from './helpers';

test('formats pictures and shapes through their contextual ribbon tabs', async ({ page }) => {
	const errors = pageErrors(page);
	const book = createWorkbook();
	const edit = createEditSession(book);
	const original = edit.addImage(
		0,
		encodePng(480, 240, new Uint8Array(480 * 240 * 3).fill(180)),
		'image/png',
		pictureAnchorAt({ row: 1, col: 1 }, 120, 60),
		'Logo',
	);
	edit.addShape(0, pictureAnchorAt({ row: 8, col: 1 }, 120, 60), {
		kind: 'preset',
		preset: 'rect',
		name: 'Box',
	});
	await page.setViewportSize({ width: 1800, height: 900 });
	await openLanding(page, 'vanilla');
	await page.locator('#landing-file').setInputFiles({
		name: 'drawings.xlsx',
		mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		buffer: Buffer.from(await saveXlsx(book)),
	});
	const host = editor(page);
	const picture = host.getByRole('img', { name: 'Logo', exact: true }).first();
	const pictureTab = host.getByRole('tab', { name: 'Picture Format' });
	const shapeTab = host.getByRole('tab', { name: 'Shape Format' });
	await expect(pictureTab).toHaveCount(0);
	await expect(shapeTab).toHaveCount(0);
	await expect(picture).toBeVisible();
	await host
		.locator('.xg-obj')
		.filter({ has: page.getByRole('img', { name: 'Logo', exact: true }) })
		.first()
		.click();
	await pictureTab.focus();
	await pictureTab.press('Enter');
	const artistic = host.getByRole('button', { name: 'Artistic Effects' });
	await expect(artistic).toBeDisabled();
	await expect(artistic).toHaveAttribute('aria-description', /baked image cache/);
	const logo = host.locator('img[alt="Logo"]').first();
	await expect(logo).toHaveJSProperty('naturalWidth', 480);
	await host.getByRole('button', { name: 'Compress Pictures', exact: true }).click();
	await host.locator('[data-dialog=picture-compress]').getByRole('button', { name: 'OK' }).click();
	await expect(logo).toHaveJSProperty('naturalWidth', 120);
	await host.getByRole('button', { name: 'Undo', exact: true }).click();
	await expect(logo).toHaveJSProperty('naturalWidth', 480);
	await host.getByRole('button', { name: 'Redo', exact: true }).click();
	await expect(logo).toHaveJSProperty('naturalWidth', 120);
	await host.getByRole('button', { name: 'Corrections', exact: true }).click();
	const corrections = host.locator('[data-dialog=picture-corrections]');
	await corrections.getByRole('spinbutton', { name: 'Brightness (%)' }).fill('20');
	await corrections.getByRole('button', { name: 'OK' }).click();
	await host.getByRole('button', { name: 'Picture Styles', exact: true }).click();
	const whiteFrame = host.getByRole('menuitem', { name: 'Simple Frame, White' });
	await expect(whiteFrame.locator('svg')).toBeVisible();
	await whiteFrame.click();

	await host.locator('.xg-shape[aria-label="Box"]').first().click();
	await expect(pictureTab).toHaveCount(0);
	await shapeTab.click();
	await host.getByRole('button', { name: 'Shape Styles', exact: true }).click();
	await host.getByRole('menuitem', { name: 'Accent 2', exact: true }).click();
	await host.getByRole('button', { name: 'Alt Text', exact: true }).click();
	const alt = host.locator('[data-dialog=drawing-alt-text]');
	await alt.getByRole('textbox', { name: 'Description' }).fill('Quarterly target');
	await alt.getByRole('button', { name: 'OK' }).click();
	await expect(host.locator('.xg-shape').first()).toHaveAttribute('aria-label', 'Quarterly target');

	const saved = await host.evaluate(async (element) => {
		const blob = await (element as unknown as { save(): Promise<Blob> }).save();
		return Array.from(new Uint8Array(await blob.arrayBuffer()));
	});
	const loaded = await loadXlsx(Buffer.from(saved));
	expect(loaded.sheets[0]?.drawings).toMatchObject([
		{
			kind: 'image',
			effects: [{ kind: 'luminance', brightness: 0.2 }],
			line: { fill: { color: { value: 'FFFFFF' } } },
		},
		{
			shape: {
				fill: { color: { kind: 'scheme', value: 'accent2' } },
				description: 'Quarterly target',
			},
		},
	]);
	const compressed = loaded.sheets[0]?.drawings[0];
	const bytes = compressed?.kind === 'image' && loaded.source?.parts.get(compressed.partName);
	if (!bytes) throw new Error('Compressed picture media is missing');
	expect(inspectRasterImage(bytes)).toMatchObject({ pixelWidth: 120, pixelHeight: 60 });
	expect(loaded.source?.parts.has(original.partName)).toBe(false);
	await host
		.locator('.xg-obj')
		.filter({ has: page.getByRole('img', { name: 'Logo', exact: true }) })
		.first()
		.click();
	await pictureTab.click();
	await host.getByRole('button', { name: 'Reset Picture', exact: true }).click();
	const resetBytes = await host.evaluate(async (element) => {
		const blob = await (element as unknown as { save(): Promise<Blob> }).save();
		return Array.from(new Uint8Array(await blob.arrayBuffer()));
	});
	const reset = (await loadXlsx(Buffer.from(resetBytes))).sheets[0]?.drawings[0];
	expect(reset).not.toHaveProperty('line');
	expect(reset).not.toHaveProperty('effects');
	await goToCell(page, 'A1');
	await expect(shapeTab).toHaveCount(0);
	expect(errors).toEqual([]);
});

test('reports a lazy gallery failure and ignores clicks while it loads', async ({ page }) => {
	const errors = pageErrors(page);
	const book = createWorkbook();
	createEditSession(book).addShape(0, pictureAnchorAt({ row: 2, col: 1 }, 120, 60), {
		kind: 'preset',
		preset: 'rect',
		name: 'Box',
	});
	let resolveChunk!: (route: Route) => void;
	const chunk = new Promise<Route>((resolve) => {
		resolveChunk = resolve;
	});
	await page.route('**/drawing-format-styles-*.js', resolveChunk);
	await page.setViewportSize({ width: 1800, height: 900 });
	await openLanding(page, 'vanilla');
	await page.locator('#landing-file').setInputFiles({
		name: 'drawings.xlsx',
		mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		buffer: Buffer.from(await saveXlsx(book)),
	});
	const host = editor(page);
	await host.locator('.xg-shape[aria-label="Box"]').first().click();
	await host.getByRole('tab', { name: 'Shape Format' }).click();
	const trigger = host.getByRole('button', { name: 'Shape Styles', exact: true });
	await trigger.dblclick();
	await expect(trigger).toHaveAttribute('aria-busy', 'true');
	await (await chunk).abort();
	await expect(host.locator('.xve-toast')).toContainText('Could not load the gallery.');
	await expect(trigger).not.toHaveAttribute('aria-busy', 'true');
	await expect(host.getByRole('menu', { name: 'Shape Styles' })).toHaveCount(0);
	expect(errors).toEqual([]);
});

test('formats multiple shapes and undoes a size change for both together', async ({ page }) => {
	const book = createWorkbook();
	const edit = createEditSession(book);
	edit.addShape(0, pictureAnchorAt({ row: 2, col: 1 }, 120, 60), {
		kind: 'preset',
		preset: 'rect',
		name: 'Box',
	});
	edit.addShape(0, pictureAnchorAt({ row: 8, col: 1 }, 180, 80), {
		kind: 'preset',
		preset: 'ellipse',
		name: 'Oval',
	});
	await page.setViewportSize({ width: 1800, height: 900 });
	await openLanding(page, 'vanilla');
	await page.locator('#landing-file').setInputFiles({
		name: 'drawings.xlsx',
		mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		buffer: Buffer.from(await saveXlsx(book)),
	});
	const host = editor(page);
	await host.locator('.xg-shape[aria-label="Box"]').first().click();
	await host
		.locator('.xg-shape[aria-label="Oval"]')
		.first()
		.click({ modifiers: ['Control'] });
	await host.getByRole('tab', { name: 'Shape Format' }).click();
	for (const [command, color] of [
		['Shape Fill', '#cc0000'],
		['Shape Outline', '#0000cc'],
	] as const) {
		await host.getByRole('button', { name: command, exact: true }).click();
		const dialog = host.locator('[data-dialog=drawing-color]');
		await dialog.locator('input[type=color]').fill(color);
		await dialog.getByRole('button', { name: 'OK' }).click();
	}
	await host.getByRole('button', { name: 'Size', exact: true }).click();
	const size = host.locator('[data-dialog=drawing-size]');
	await size.getByRole('spinbutton', { name: 'Width (px)' }).fill('150');
	await size.getByRole('spinbutton', { name: 'Height (px)' }).fill('72');
	await size.getByRole('button', { name: 'OK' }).click();

	const read = async () => {
		const bytes = await host.evaluate(async (element) => {
			const blob = await (element as unknown as { save(): Promise<Blob> }).save();
			return Array.from(new Uint8Array(await blob.arrayBuffer()));
		});
		return (await loadXlsx(Buffer.from(bytes))).sheets[0]?.drawings;
	};
	const formatted = {
		anchor: { ext: { cx: 1428750, cy: 685800 } },
		shape: { fill: { color: { value: 'CC0000' } }, line: { fill: { color: { value: '0000CC' } } } },
	};
	expect(await read()).toMatchObject([formatted, formatted]);
	await host.getByRole('button', { name: 'Undo', exact: true }).click();
	expect(await read()).toMatchObject([
		{ ...formatted, anchor: { ext: { cx: 1143000, cy: 571500 } } },
		{ ...formatted, anchor: { ext: { cx: 1714500, cy: 762000 } } },
	]);
});
