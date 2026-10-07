import { test, expect } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { parseVsdx } from 'ooxml-core/visio';
import { downloadCopy } from './ribbon';

const directory = process.env.VISIO_NATIVE_LINE_MOVEMENT_DIR;
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: native lines move, undo, redo, download and reload through the shared editor`, async ({
		page,
	}) => {
		test.skip(!directory, 'Set VISIO_NATIVE_LINE_MOVEMENT_DIR to the native line capture.');
		const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
			cases: { shapeId: string; after: Record<string, { value: number }> }[];
		};
		const nativeMoved = await parseVsdx(await readFile(join(directory!, 'moved.vsdx')));
		await page.goto(framework === 'vanilla' ? '/demo/?sample=1' : `/demo-${framework}/?sample=1`);
		await page.locator('#file').setInputFiles(join(directory!, 'original.vsdx'));
		await expect(page.locator('#file-name')).toHaveText('original.vsdx');
		const viewer = page.locator('visio-viewer');
		await viewer.locator('.edit-controls summary').click();
		const transforms = new Map<string, string>();
		for (const item of evidence.cases) {
			const line = viewer.locator(`[data-shape-id="${item.shapeId}"]`);
			await line.focus();
			await line.press('Enter');
			await expect(line).toHaveAttribute('data-selected', 'true');
			const original = (await line.getAttribute('transform'))!;
			await page.getByLabel('Pin X (inches)').fill(String(item.after.PinX!.value));
			await page.getByLabel('Pin Y (inches)').fill(String(item.after.PinY!.value));
			await page.getByRole('button', { name: 'Move selected', exact: true }).click();
			await expect(page.getByLabel('Pin X (inches)')).toHaveValue('');
			await expect(viewer.locator('[data-geometry-error]')).toBeHidden();
			await expect(line).not.toHaveAttribute('transform', original);
			const moved = (await line.getAttribute('transform'))!;
			await viewer
				.locator('.edit-controls')
				.getByRole('button', { name: 'Undo', exact: true })
				.click();
			await expect(line).toHaveAttribute('transform', original);
			await viewer
				.locator('.edit-controls')
				.getByRole('button', { name: 'Redo', exact: true })
				.click();
			await expect(line).toHaveAttribute('transform', moved);
			transforms.set(item.shapeId, moved);
		}
		const command = await downloadCopy(viewer);
		const pending = page.waitForEvent('download');
		await command.click();
		const download = await pending;
		const stream = await download.createReadStream();
		const chunks: Buffer[] = [];
		for await (const chunk of stream!) chunks.push(Buffer.from(chunk));
		const bytes = Buffer.concat(chunks);
		const saved = await parseVsdx(bytes);
		expect(saved.pages[0]!.shapes).toHaveLength(4);
		for (const actual of saved.pages[0]!.shapes) {
			const expected = nativeMoved.pages[0]!.shapes.find((shape) => shape.id === actual.id)!;
			for (let i = 0; i < 6; i++)
				expect(actual.transform[i]).toBeCloseTo(expected.transform[i]!, 12);
			expect(actual.width).toBeCloseTo(expected.width, 12);
			expect(actual.height).toBe(expected.height);
			expect(actual.geometry).toEqual(expected.geometry);
			expect(actual.style).toEqual(expected.style);
		}
		await page
			.locator('#file')
			.setInputFiles({
				name: 'core-moved.vsdx',
				mimeType: 'application/vnd.ms-visio.drawing',
				buffer: bytes,
			});
		await expect(page.locator('#file-name')).toHaveText('core-moved.vsdx');
		for (const [id, transform] of transforms)
			await expect(viewer.locator(`[data-shape-id="${id}"]`)).toHaveAttribute(
				'transform',
				transform,
			);
	});
}
