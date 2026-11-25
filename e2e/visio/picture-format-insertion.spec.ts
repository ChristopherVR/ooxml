import { expect, test } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';
import { openDemo } from './demo-page';

function bitmap(): Buffer {
	const bytes = Buffer.alloc(62);
	bytes.write('BM');
	bytes.writeUInt32LE(bytes.length, 2);
	bytes.writeUInt32LE(54, 10);
	bytes.writeUInt32LE(40, 14);
	bytes.writeInt32LE(2, 18);
	bytes.writeInt32LE(1, 22);
	bytes.writeUInt16LE(1, 26);
	bytes.writeUInt16LE(24, 28);
	bytes[56] = bytes[59] = 255;
	return bytes;
}

function tiff(): Buffer {
	const fields = [
		[256, 2],
		[257, 1],
		[258, 8],
		[259, 1],
		[262, 1],
		[273, 8],
		[278, 1],
		[279, 2],
	];
	const bytes = Buffer.alloc(16 + 12 * fields.length);
	bytes.write('II');
	bytes.writeUInt16LE(42, 2);
	bytes.writeUInt32LE(10, 4);
	bytes.fill(128, 8, 10);
	bytes.writeUInt16LE(fields.length, 10);
	for (const [index, [tag, value]] of fields.entries()) {
		const offset = 12 + 12 * index;
		bytes.writeUInt16LE(tag!, offset);
		bytes.writeUInt16LE(4, offset + 2);
		bytes.writeUInt32LE(1, offset + 4);
		bytes.writeUInt32LE(value!, offset + 8);
	}
	return bytes;
}

test('Visio picks SVG and BMP files and decodes dropped TIFF to visible pixels', async ({
	page,
}) => {
	const decoderRequests: string[] = [];
	page.context().on('request', (request) => {
		if (/\/UTIF-[\w-]+\.js$/.test(new URL(request.url()).pathname))
			decoderRequests.push(request.url());
	});
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await expect(viewer.locator('svg.paper')).toContainText('Release workflow');
	await page.locator('#file').setInputFiles({
		name: 'Pictures.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Picture formats'),
	});
	await expect(viewer.locator('svg.paper')).toContainText('Picture formats');
	await viewer.getByRole('tab', { name: 'Insert', exact: true }).click();
	const picker = viewer.locator('[data-insert-picture]');
	const images = viewer.locator('svg.paper image');
	await picker.setInputFiles({
		name: 'red.svg',
		mimeType: 'image/svg+xml',
		buffer: Buffer.from(
			'<svg xmlns="http://www.w3.org/2000/svg" width="2" height="1"><rect width="2" height="1" fill="red"/></svg>',
		),
	});
	await expect(images).toHaveCount(1);
	await picker.setInputFiles({ name: 'red.bmp', mimeType: 'image/bmp', buffer: bitmap() });
	await expect(images).toHaveCount(2);
	expect(decoderRequests).toEqual([]);
	await viewer.locator('svg.paper').evaluate((paper, bytes) => {
		const dataTransfer = new DataTransfer();
		dataTransfer.items.add(new File([new Uint8Array(bytes)], 'gray.tif', { type: 'image/tiff' }));
		for (const type of ['dragenter', 'dragover', 'drop'])
			paper.dispatchEvent(
				new DragEvent(type, { dataTransfer, bubbles: true, composed: true, cancelable: true }),
			);
	}, Array.from(tiff()));
	await expect(images).toHaveCount(3);
	await expect.poll(() => decoderRequests.length).toBeGreaterThan(0);
	expect([...new Set(decoderRequests.map((url) => new URL(url).origin))]).toEqual([
		new URL(page.url()).origin,
	]);
	const pixels = await images.evaluateAll(async (elements) =>
		Promise.all(
			elements.map(async (element) => {
				const src =
					element.getAttribute('href') ??
					element.getAttributeNS('http://www.w3.org/1999/xlink', 'href');
				if (!src) throw new Error('Picture source is missing.');
				const image = new Image();
				image.src = src;
				await image.decode();
				const canvas = document.createElement('canvas');
				canvas.width = 1;
				canvas.height = 1;
				const context = canvas.getContext('2d')!;
				context.drawImage(image, 0, 0);
				return [...context.getImageData(0, 0, 1, 1).data];
			}),
		),
	);
	expect(pixels).toEqual([
		[255, 0, 0, 255],
		[255, 0, 0, 255],
		[128, 128, 128, 255],
	]);
	await expect(page.locator('#drop-overlay')).toBeHidden();
});
