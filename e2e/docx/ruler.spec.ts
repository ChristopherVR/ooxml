import { expect, test, type Page } from '@playwright/test';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { newDocument } from './helpers';

// View > Ruler draws the shared `office-ui-ruler` above the page; its indent markers live in its
// shadow root and apply the paragraph's indent on release as one undoable step.
const firstIndent = (page: Page) =>
	page.locator('docx-editor').evaluate((element) => {
		const block = (element as DocxEditorElement).documentModel?.blocks[0];
		return block?.type === 'paragraph' ? (block.indentLeftTwips ?? 0) : null;
	});

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: the ruler shows in Print Layout and a left-indent drag is one undo step`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 1600, height: 900 });
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror');
		await surface.click();
		await page.keyboard.type('Ruled paragraph');

		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
		await expect(editor.locator('.dve-print-line').first()).toBeVisible();
		await editor.getByRole('button', { name: 'Ruler', exact: true }).click();
		const ruler = editor.getByRole('img', { name: 'Ruler', exact: true });
		await expect(ruler).toBeVisible();
		await expect(ruler).toHaveCount(1);

		const left = ruler.locator('[data-marker="left"]');
		await expect(left).toBeVisible();
		expect(await firstIndent(page)).toBe(0);
		// One inch on screen is 96 CSS pixels times the ruler's zoom.
		const scale = await ruler.evaluate(
			(element) =>
				element.getBoundingClientRect().width /
				(element as HTMLElement & { extent: number }).extent,
		);
		const box = (await left.boundingBox())!;
		const offset = async () => Math.round(((await left.boundingBox())!.x - box.x) / scale);
		const x = box.x + box.width / 2;
		const y = box.y + box.height / 2;
		await page.mouse.move(x, y);
		await page.mouse.down();
		await page.mouse.move(x + 48 * scale, y, { steps: 4 });
		await page.mouse.move(x + 96 * scale, y, { steps: 4 });
		await page.mouse.up();
		await expect.poll(() => firstIndent(page)).toBe(1440);
		await expect(surface.locator('p').first()).toHaveText('Ruled paragraph');
		// The marker follows the committed indent.
		await expect.poll(offset).toBe(96);

		// Undo removes the indent and nothing else: the typing is still there.
		await editor.getByRole('tab', { name: 'Home', exact: true }).click();
		await editor.getByRole('button', { name: 'Undo', exact: true }).click();
		await expect.poll(() => firstIndent(page)).toBe(0);
		await expect(surface.locator('p').first()).toHaveText('Ruled paragraph');
		await expect.poll(offset).toBe(0);
	});
}
