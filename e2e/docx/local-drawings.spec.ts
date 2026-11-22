import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { dialogByHeading, newDocument, reopenDocument, reveal, saveButton } from './helpers';

test.use({ viewport: { width: 375, height: 812 }, hasTouch: true, isMobile: true });

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: inserts local 3D Models and Office Objects from the phone ribbon`, async ({
		page,
	}, testInfo) => {
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		const body = editor.locator('.ProseMirror');
		await body.tap();
		await page.keyboard.type('abcd');
		await page.keyboard.press('ArrowLeft');
		await page.keyboard.press('ArrowLeft');
		await editor.getByRole('tab', { name: 'Insert', exact: true }).tap();
		const child = await readFile(
			new URL('../../src/core/docx/__fixtures__/field-result-replace/empty.docx', import.meta.url),
		);
		for (const { heading, file } of [
			{
				heading: '3D Models',
				file: { name: 'Triangle.glb', mimeType: 'model/gltf-binary', buffer: triangleGlb() },
			},
			{
				heading: 'Object',
				file: {
					name: 'Attachment.docx',
					mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
					buffer: child,
				},
			},
		]) {
			const button = editor.getByRole('button', { name: heading, exact: true });
			await reveal(editor, button);
			await button.tap();
			const picker = dialogByHeading(editor, heading);
			await picker.locator('input[type=file]').setInputFiles(file);
			await picker.getByRole('button', { name: 'Insert', exact: true }).tap();
			await expect(body.locator(`img[alt="${file.name}"]`)).toBeVisible();
		}
		const pending = page.waitForEvent('download');
		await saveButton(page).tap();
		const path = testInfo.outputPath('native-local-drawings.docx');
		await (await pending).saveAs(path);
		const artifact = await readFile(path);
		const zip = await JSZip.loadAsync(artifact);
		const xml = await zip.file('word/document.xml')?.async('string');
		expect(xml).toMatch(
			/<w:t>ab<\/w:t>[\s\S]*<am3d:model3d[\s\S]*<o:OLEObject[\s\S]*<w:t>cd<\/w:t>/,
		);
		const embedded = zip.file(/^word\/embeddings\/[^/]+\.docx$/)[0];
		expect(await embedded?.async('nodebuffer')).toEqual(child);
		await reopenDocument(page, 'native-local-drawings.docx', artifact);
		await expect(body.locator('img[alt="Triangle.glb"]')).toBeVisible();
		await expect(body.locator('img[alt="Attachment.docx"]')).toBeVisible();
	});
}

function triangleGlb(): Buffer {
	const scene = JSON.stringify({
		asset: { version: '2.0' },
		scene: 0,
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
	const bytes = Buffer.alloc(28 + json.length + 36);
	bytes.write('glTF');
	bytes.writeUInt32LE(2, 4);
	bytes.writeUInt32LE(bytes.length, 8);
	bytes.writeUInt32LE(json.length, 12);
	bytes.write('JSON', 16);
	json.copy(bytes, 20);
	bytes.writeUInt32LE(36, 20 + json.length);
	bytes.write('BIN\0', 24 + json.length);
	[0, 0, 0, 1, 0, 0, 0, 1, 0].forEach((value, index) =>
		bytes.writeFloatLE(value, 28 + json.length + index * 4),
	);
	return bytes;
}
