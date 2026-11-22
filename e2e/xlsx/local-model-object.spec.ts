import { expect, test } from '@playwright/test';
import { createWorkbook, loadXlsx, saveXlsx } from 'ooxml-core/xlsx';
import { clickRibbonControl, editor, goToCell, newWorkbook, openLanding } from './helpers';

function triangleGlb() {
	const scene = JSON.stringify({
		asset: { version: '2.0' },
		scenes: [{ nodes: [0] }],
		nodes: [{ mesh: 0 }],
		meshes: [{ primitives: [{ attributes: { POSITION: 0 } }] }],
		accessors: [
			{
				bufferView: 0,
				componentType: 5126,
				count: 3,
				type: 'VEC3',
				min: [0, 0, 0],
				max: [1, 1, 0],
			},
		],
		bufferViews: [{ buffer: 0, byteLength: 36 }],
		buffers: [{ byteLength: 36 }],
	});
	const json = Buffer.from(scene.padEnd(Math.ceil(scene.length / 4) * 4, ' '));
	const glb = Buffer.alloc(28 + json.length + 36);
	glb.write('glTF');
	glb.writeUInt32LE(2, 4);
	glb.writeUInt32LE(glb.length, 8);
	glb.writeUInt32LE(json.length, 12);
	glb.writeUInt32LE(0x4e4f534a, 16);
	json.copy(glb, 20);
	glb.writeUInt32LE(36, 20 + json.length);
	glb.writeUInt32LE(0x004e4942, 24 + json.length);
	[0, 0, 0, 1, 0, 0, 0, 1, 0].forEach((point, i) =>
		glb.writeFloatLE(point, 28 + json.length + i * 4),
	);
	return glb;
}

test.use({ viewport: { width: 375, height: 812 }, hasTouch: true, isMobile: true });

test('inserts and reopens a 3D model and Office object from the phone ribbon', async ({ page }) => {
	await newWorkbook(page);
	const host = editor(page);
	const insert = host.getByRole('tab', { name: 'Insert', exact: true });
	const picker = host.locator('office-ui-dialog[open]');
	await insert.tap();
	await host.getByRole('button', { name: 'Illustrations', exact: true }).tap();
	await host.getByRole('button', { name: '3D Models', exact: true }).tap();
	await picker.locator('input[type="file"]').setInputFiles({
		name: 'Triangle.glb',
		mimeType: 'model/gltf-binary',
		buffer: triangleGlb(),
	});
	await picker.getByRole('button', { name: 'Insert', exact: true }).click();
	await expect(host.getByRole('img', { name: 'Triangle.glb' })).toBeVisible();

	await goToCell(page, 'B12');
	await insert.tap();
	await host.getByRole('button', { name: 'Symbols', exact: true }).tap();
	await host.getByRole('button', { name: 'Object', exact: true }).tap();
	const source = await saveXlsx(createWorkbook());
	await picker.locator('input[type="file"]').setInputFiles({
		name: 'Source.xlsx',
		mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		buffer: Buffer.from(source),
	});
	await picker.getByRole('button', { name: 'Insert', exact: true }).click();
	await expect(host.getByRole('img', { name: 'Source.xlsx' })).toBeVisible();
	const bytes = Uint8Array.from(
		await host.evaluate(async (node) =>
			Array.from(await (node as HTMLElement & { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		),
	);
	const saved = await loadXlsx(bytes);
	const [model, embedded] = saved.sheets[0]?.drawings ?? [];
	expect(model?.kind).toBe('model3d');
	expect(embedded).toMatchObject({ kind: 'embeddedObject', anchor: { from: { row: 11, col: 1 } } });
	if (embedded?.kind !== 'embeddedObject') throw new Error('Missing embedded object');
	expect(saved.source?.parts.get(embedded.partName)).toEqual(source);
	await openLanding(page);
	await page.locator('#landing-file').setInputFiles({
		name: 'Reopened.xlsx',
		mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		buffer: Buffer.from(bytes),
	});
	await expect(host.getByRole('img', { name: 'Triangle.glb' })).toBeVisible();
	await host.getByRole('img', { name: 'Source.xlsx' }).click({ force: true });
	await host.getByRole('tab', { name: 'Embedded Object', exact: true }).click();
	const downloadReady = page.waitForEvent('download');
	await clickRibbonControl(
		page,
		host.getByRole('button', { name: 'Download Embedded File', exact: true }),
	);
	expect((await downloadReady).suggestedFilename()).toBe('Source.xlsx');
});
