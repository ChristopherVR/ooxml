import { expect, test } from '@playwright/test';
import { downloadCopy } from './ribbon';
import { openDemo } from './demo-page';
import { fixture, inventory, handlePoint } from './selection-resize-fixture';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: local group resize preserves members through undo, reopen and ungroup`, async ({
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
		await page.locator('#file').setInputFiles({
			name: 'group-resize.vsdx',
			mimeType: 'application/vnd.ms-visio.drawing',
			buffer: await fixture(),
		});
		await expect(viewer.locator('svg.paper')).toContainText('Selection resize');
		await inventory(viewer);
		await viewer.locator('svg.paper [data-shape-id="1"]').click();
		await viewer.locator('svg.paper [data-shape-id="2"]').click({ modifiers: ['Shift'] });
		await page.keyboard.press('Control+Shift+G');
		await expect.poll(async () => (await inventory(viewer)).shapes.length).toBe(1);
		await expect(viewer.locator('[data-resize-handle]')).toHaveCount(8);
		const before = await inventory(viewer),
			start = await handlePoint(viewer);
		expect(before.shapes[0]!.children).toHaveLength(2);
		for (const cancel of [true, false]) {
			await page.mouse.move(start.x, start.y);
			await page.mouse.down();
			await page.mouse.move(start.x + start.dx, start.y + start.dy, { steps: 4 });
			await expect(viewer.locator('[data-resize-frame-preview]')).toHaveCount(1);
			expect((await inventory(viewer)).bytes).toEqual(before.bytes);
			if (cancel) await page.keyboard.press('Escape');
			await page.mouse.up();
			if (cancel) expect((await inventory(viewer)).bytes).toEqual(before.bytes);
		}
		await expect.poll(async () => (await inventory(viewer)).shapes[0]!.width).toBeCloseTo(5, 5);
		const resized = await inventory(viewer),
			group = resized.shapes[0]!;
		expect(group.height).toBeCloseTo(1.5, 5);
		group.transform.forEach((value, i) =>
			expect(value).toBeCloseTo(before.shapes[0]!.transform[i]!, 9),
		);
		for (const [i, member] of group.children.entries()) {
			const previous = before.shapes[0]!.children[i]!;
			expect(member.width).toBeCloseTo((previous.width * group.width) / 4, 9);
			expect(member.height).toBeCloseTo(previous.height * group.height, 9);
			expect(member.pin!.pinX).toBeCloseTo((previous.pin!.pinX * group.width) / 4, 9);
			expect(member.pin!.pinY).toBeCloseTo(previous.pin!.pinY * group.height, 9);
		}
		expect(resized.error).toBe('');
		await viewer.locator('.qat [data-command="undo"]').click();
		await expect.poll(async () => (await inventory(viewer)).bytes).toEqual(before.bytes);
		await viewer.locator('.qat [data-command="redo"]').click();
		await expect.poll(async () => (await inventory(viewer)).bytes).toEqual(resized.bytes);
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
		// A group can have empty space at its bounding-box center. Select its accessible
		// group target through the keyboard rather than clicking that empty SVG point.
		const reopenedGroup = viewer.locator('svg.paper [data-shape-id="3"]');
		await reopenedGroup.focus();
		await reopenedGroup.press('Enter');
		await expect(viewer.locator('[data-resize-handle]')).toHaveCount(8);
		await page.keyboard.press('Control+Shift+U');
		await expect.poll(async () => (await inventory(viewer)).shapes.length).toBe(2);
		const ungrouped = await inventory(viewer);
		ungrouped.shapes.forEach((member, i) => {
			expect(member.width).toBe(group.children[i]!.width);
			expect(member.height).toBe(group.children[i]!.height);
			expect(member.transform[4]).toBeCloseTo(
				group.transform[4]! + group.children[i]!.transform[4]!,
				9,
			);
			expect(member.transform[5]).toBeCloseTo(
				group.transform[5]! + group.children[i]!.transform[5]!,
				9,
			);
		});
		expect(errors).toEqual([]);
	});
}
