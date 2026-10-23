import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { dialogByHeading, fileInput, newDocument } from './helpers';

const FRAMEWORKS = ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'] as const;

const editorOf = (page: Page) => page.locator('docx-editor');

async function saveBytes(page: Page): Promise<Uint8Array> {
	const list = await editorOf(page).evaluate(async (element) =>
		Array.from(await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
	);
	return Uint8Array.from(list);
}

const pixels = (style: string | null) => Number(/width:\s*(\d+)px/.exec(style ?? '')?.[1]);

for (const framework of FRAMEWORKS) {
	test(`${framework}: Format pane and resizing a chart save and reload`, async ({ page }) => {
		test.setTimeout(120_000);
		await page.setViewportSize({ width: 1400, height: 900 });
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = editorOf(page);
		await editor.locator('.ProseMirror').click();
		await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
		await editor.getByRole('button', { name: 'Insert chart', exact: true }).click();
		await dialogByHeading(editor, 'Insert Chart').getByRole('button', { name: 'OK' }).click();
		const chart = editor.locator('.ProseMirror [data-docx-chart="1"]');
		await expect(chart).toHaveCount(1);

		// Chart Design > Format Selection docks the shared Format pane beside the page.
		await editor.getByRole('button', { name: 'Format Selection', exact: true }).click();
		const pane = editor.locator('office-ui-chart-format-pane');
		await expect(pane).toBeVisible();
		await expect(pane.locator('h3').first()).toHaveText('Format Chart Area');

		// Series 1: a red solid fill repaints the first series at once.
		await pane.locator('[data-chart-format-target]').selectOption('series:0');
		await pane.getByRole('radio', { name: 'Solid fill', exact: true }).click();
		await pane.locator('[data-chart-format-field="fill.color"] input').fill('#ff0000');
		await expect(chart.locator('g[data-chart-series="0"] > *').first()).toHaveAttribute(
			'fill',
			/#FF0000/i,
		);

		// Legend: moved to the top from Legend Options.
		await pane.locator('[data-chart-format-target]').selectOption('legend');
		await pane.locator('[data-chart-format-tab="options"]').click();
		await pane.getByRole('radio', { name: 'Top', exact: true }).click();
		await pane.locator('[data-chart-format-close]').click();
		await expect(pane).toBeHidden();

		// The corner handle resizes the selected chart; the chart repaints at its new size.
		await chart.click();
		const before = pixels(await chart.getAttribute('style'));
		const handle = editor.locator('.ProseMirror .dve-chart-frame .dve-picture-handle');
		await expect(handle).toBeVisible();
		const box = (await handle.boundingBox())!;
		await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
		await page.mouse.down();
		await page.mouse.move(box.x - 150, box.y - 80, { steps: 5 });
		await page.mouse.up();
		await expect
			.poll(async () => pixels(await chart.getAttribute('style')))
			.toBeLessThan(before - 50);
		const resized = pixels(await chart.getAttribute('style'));

		const bytes = await saveBytes(page);
		const zip = await JSZip.loadAsync(bytes);
		const part = await zip.file('word/charts/chart1.xml')!.async('string');
		expect(part).toContain('<a:srgbClr val="FF0000"/>');
		expect(part).toContain('<c:legendPos val="t"/>');
		const documentXml = await zip.file('word/document.xml')!.async('string');
		expect(documentXml).toContain(`<wp:extent cx="${resized * 9525}"`);

		// Reloaded: the formatted chart comes back at its new size.
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'chart-formatted.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(bytes),
		});
		const reloaded = editor.locator('.ProseMirror [data-docx-chart="1"]');
		await expect(reloaded).toHaveCount(1);
		expect(pixels(await reloaded.getAttribute('style'))).toBe(resized);
		await expect(reloaded.locator('g[data-chart-series="0"] > *').first()).toHaveAttribute(
			'fill',
			/#FF0000/i,
		);
	});
}
