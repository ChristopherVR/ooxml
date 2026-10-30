import { expect, test } from '@playwright/test';
import { downloadCopy } from './ribbon';
import { openDemo } from './demo-page';
import { fixture, inventory, handlePoint } from './selection-resize-fixture';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: multi-selection resize is atomic across preview, undo, save and reopen`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.setViewportSize({ width: 1600, height: 1000 });
		await openDemo(
			page,
			framework === 'vanilla' ? '/demo/?sample=1' : `/demo-${framework}/?sample=1`,
		);
		const viewer = page.locator('visio-viewer');
		const load = async (locked = false) => {
			await page.locator('#file').setInputFiles({
				name: 'selection-resize.vsdx',
				mimeType: 'application/vnd.ms-visio.drawing',
				buffer: await fixture(locked),
			});
			await expect(page.locator('#file-name')).toHaveText('selection-resize.vsdx');
			await expect(viewer.locator('svg.paper')).toContainText('Selection resize');
			await expect.poll(async () => (await inventory(viewer)).busy).toBe(false);
			await viewer.locator('svg.paper [data-shape-id="1"]').click();
			await viewer.locator('svg.paper [data-shape-id="2"]').click({ modifiers: ['Shift'] });
			await expect.poll(async () => (await inventory(viewer)).selection).toEqual(['1', '2']);
			await expect(viewer.locator('[data-resize-handle]')).toHaveCount(8);
		};
		await load();
		const before = await inventory(viewer);
		const start = await handlePoint(viewer);
		await page.mouse.move(start.x, start.y);
		await page.mouse.down();
		await page.mouse.move(start.x + start.dx, start.y + start.dy, { steps: 4 });
		await expect
			.poll(async () =>
				Number(await viewer.locator('[data-resize-frame-preview]').getAttribute('width')),
			)
			.toBeCloseTo(6, 5);
		await expect
			.poll(async () =>
				Number(await viewer.locator('[data-resize-frame-preview]').getAttribute('height')),
			)
			.toBeCloseTo(1.5, 5);
		// Browser pointer/CTM coordinates have finite precision. Compare saved geometry to the
		// actual preview more strictly than the nominal two-inch, half-inch pointer displacement.
		const projected = await viewer.locator('[data-resize-frame-preview]').evaluate((node) => ({
			width: Number(node.getAttribute('width')),
			height: Number(node.getAttribute('height')),
		}));
		expect((await inventory(viewer)).bytes).toEqual(before.bytes);
		expect((await inventory(viewer)).shapes).toEqual(before.shapes);
		await page.mouse.up();
		await expect
			.poll(async () => (await inventory(viewer)).shapes[0]!.width)
			.toBeCloseTo(projected.width / 4, 9);
		const resized = await inventory(viewer);
		expect(resized.selection).toEqual(['1', '2']);
		for (const shape of resized.shapes) {
			expect(shape.width).toBeCloseTo(projected.width / 4, 9);
			expect(shape.height).toBeCloseTo(projected.height, 9);
			expect(shape.pin!.pinY).toBeCloseTo(6.5 + projected.height / 2, 9);
		}
		expect(resized.shapes[0]!.pin!.pinX).toBeCloseTo(1.5 + projected.width / 8, 9);
		expect(resized.shapes[1]!.pin!.pinX).toBeCloseTo(1.5 + (3.5 * projected.width) / 4, 9);
		await viewer.locator('.qat [data-command="undo"]').click();
		await expect.poll(async () => (await inventory(viewer)).bytes).toEqual(before.bytes);
		await expect(viewer.locator('.qat [data-command="undo"]')).toBeDisabled();
		await viewer.locator('.qat [data-command="redo"]').click();
		await expect.poll(async () => (await inventory(viewer)).bytes).toEqual(resized.bytes);
		const cancel = await handlePoint(viewer);
		await page.mouse.move(cancel.x, cancel.y);
		await page.mouse.down();
		await page.mouse.move(cancel.x + 25, cancel.y - 20);
		await page.keyboard.press('Escape');
		await page.mouse.up();
		expect((await inventory(viewer)).bytes).toEqual(resized.bytes);
		const downloading = page.waitForEvent('download');
		await (await downloadCopy(viewer)).click();
		const stream = await (await downloading).createReadStream(),
			chunks: Buffer[] = [];
		for await (const chunk of stream!) chunks.push(Buffer.from(chunk));
		await page.locator('#file').setInputFiles({
			name: 'reopened.vsdx',
			mimeType: 'application/vnd.ms-visio.drawing',
			buffer: Buffer.concat(chunks),
		});
		await expect(page.locator('#file-name')).toHaveText('reopened.vsdx');
		await expect.poll(async () => (await inventory(viewer)).shapes).toEqual(resized.shapes);
		await load(true);
		const protectedBefore = await inventory(viewer),
			locked = await handlePoint(viewer);
		await page.mouse.move(locked.x, locked.y);
		await page.mouse.down();
		await page.mouse.move(locked.x + locked.dx, locked.y + locked.dy);
		await page.mouse.up();
		// A common scale would change the LockWidth member, so the whole selection holds still.
		expect((await inventory(viewer)).error ?? '').toBe('');
		expect((await inventory(viewer)).bytes).toEqual(protectedBefore.bytes);
		await expect(viewer.locator('.qat [data-command="undo"]')).toBeDisabled();
		expect(errors).toEqual([]);
	});
}
