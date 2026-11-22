import { readFile } from 'node:fs/promises';

import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';

import { loadDeck, ribbonTab, SAMPLE_DECK, slideElements, slideStage } from './support/deck';

declare global {
	interface Window {
		screenRecordingTestStream?: MediaStream;
	}
}

test.use({ viewport: { width: 375, height: 812 }, hasTouch: true, isMobile: true });

const recordingDock = (page: Page) => page.getByRole('group', { name: 'Screen Recording' });

async function openRecordingPicker(page: Page) {
	await page
		.getByRole('navigation', { name: 'Editor actions' })
		.getByRole('button', { name: 'Insert' })
		.tap();
	await page
		.locator('pptx-ui-mobile-insert')
		.getByRole('button', { name: 'Screen Recording', exact: true })
		.tap();
	return picker(page, 'Screen Recording');
}

const picker = (page: Page, name: string) =>
	page.locator('office-ui-dialog[open]').filter({
		has: page.getByRole('heading', { name, exact: true }),
	});

test.beforeEach(async ({ page }) => {
	await page.addInitScript(() => {
		const devices = navigator.mediaDevices;
		Object.defineProperty(devices, 'getDisplayMedia', {
			configurable: true,
			value: () => {
				const canvas = document.createElement('canvas');
				canvas.width = 320;
				canvas.height = 180;
				const context = canvas.getContext('2d')!;
				let frame = 0;
				const tick = () => {
					context.fillStyle = frame++ % 2 ? '#0451a5' : '#00aa77';
					context.fillRect(0, 0, canvas.width, canvas.height);
				};
				tick();
				const stream = canvas.captureStream(20);
				window.screenRecordingTestStream = stream;
				const timer = setInterval(() => {
					if (stream.getVideoTracks()[0]?.readyState === 'ended') clearInterval(timer);
					else tick();
				}, 50);
				return Promise.resolve(stream);
			},
		});
	});
	await loadDeck(page);
});

test('captures a screen clipping as an image on the slide', async ({ page }) => {
	const before = await slideElements(page).count();
	await page
		.getByRole('navigation', { name: 'Editor actions' })
		.getByRole('button', { name: 'Insert' })
		.tap();
	await page
		.locator('pptx-ui-mobile-insert')
		.getByRole('button', { name: 'Screenshot', exact: true })
		.tap();
	await page
		.locator('pptx-ui-mobile-insert')
		.getByRole('button', { name: 'Screen Clipping', exact: true })
		.tap();
	const dialog = picker(page, 'Screen Clipping');
	await dialog.getByRole('spinbutton', { name: /^(W|Width)$/ }).fill('100');
	await dialog.getByRole('button', { name: 'Insert', exact: true }).tap();
	await expect(slideElements(page)).toHaveCount(before + 1);
	await expect(slideStage(page).getByRole('img', { name: 'Screenshot' })).toBeVisible();
});

test('embeds a local screen recording from the floating dock', async ({ page }) => {
	const before = await slideElements(page).count();
	const dialog = await openRecordingPicker(page);
	await dialog.getByRole('checkbox', { name: 'Record pointer' }).uncheck();
	await dialog.getByRole('checkbox', { name: 'Include shared audio' }).uncheck();
	await dialog.getByRole('button', { name: 'Select Area' }).tap();
	const width = dialog.getByRole('spinbutton', { name: /^(W|Width)$/ });
	await width.fill('40');
	await expect(dialog.getByRole('button', { name: 'Record', exact: true })).toBeDisabled();
	await width.fill('96');
	await dialog.getByRole('spinbutton', { name: /^(H|Height)$/ }).fill('64');
	await dialog.getByRole('button', { name: 'Record', exact: true }).tap();
	await expect(dialog).toHaveCount(0);
	const dock = recordingDock(page);
	const controls = dock.getByRole('button', { name: 'Recording controls' });
	const pause = dock.getByRole('button', { name: 'Pause', exact: true });
	await expect(controls).toBeFocused();
	await controls.tap();
	await expect(pause).toBeHidden();
	await controls.press('Enter');
	await expect(pause).toBeVisible();
	const pin = dock.getByRole('button', { name: 'Pin dock' });
	await pin.tap();
	await expect(pin).toHaveAttribute('aria-pressed', 'true');
	await pause.tap();
	await dock.getByRole('button', { name: 'Resume', exact: true }).tap();
	await dock.getByRole('button', { name: 'Stop', exact: true }).tap();
	await expect(dock).toHaveCount(0);
	await expect(slideElements(page)).toHaveCount(before + 1);
	const box = await page.evaluate(() =>
		window.__pptxViewer.getSlides()[0]!.elements.find((element) => element.type === 'media'),
	);
	expect(box).toMatchObject({ width: 96, height: 64 });
	const pending = page.waitForEvent('download');
	await page
		.getByRole('toolbar', { name: 'Toolbar', exact: true })
		.getByRole('button', { name: 'Save', exact: true })
		.tap();
	const archive = await JSZip.loadAsync(await readFile(await (await pending).path()));
	expect(Object.keys(archive.files)).toEqual(
		expect.arrayContaining([expect.stringMatching(/^ppt[/]media[/]video\d+\.(mp4|webm)$/u)]),
	);
});

test('cancels a recording from the dock, releases capture and allows a retry', async ({ page }) => {
	const before = await slideElements(page).count();
	const dialog = await openRecordingPicker(page);
	await dialog.getByRole('button', { name: 'Record', exact: true }).tap();
	const dock = recordingDock(page);
	await dock.getByRole('button', { name: 'Cancel', exact: true }).tap();
	await expect(dock).toHaveCount(0);
	await expect(slideElements(page)).toHaveCount(before);
	const canRecordAgain = page
		.locator('pptx-ui-mobile-insert')
		.getByRole('button', { name: 'Screen Recording', exact: true });
	await canRecordAgain.tap();
	await expect(dialog).toBeVisible();
	const state = await page.evaluate(
		() => window.screenRecordingTestStream?.getVideoTracks()[0]?.readyState,
	);
	expect(state).toBe('ended');
});

test.describe('recording dock on desktop', () => {
	test.use({ viewport: { width: 1920, height: 900 }, hasTouch: false, isMobile: false });

	test('hides unpinned controls and inserts the video when sharing stops', async ({ page }) => {
		const before = await slideElements(page).count();
		await ribbonTab(page, 'Insert').click();
		await page.getByRole('button', { name: 'Screen Recording', exact: true }).click();
		await picker(page, 'Screen Recording')
			.getByRole('button', { name: 'Record', exact: true })
			.click();
		const dock = recordingDock(page);
		const controls = dock.getByRole('button', { name: 'Recording controls' });
		await expect(controls).toHaveAttribute('aria-expanded', 'false');
		await controls.hover();
		await expect(dock.getByRole('button', { name: 'Stop', exact: true })).toBeVisible();
		await page.evaluate(() => {
			window.screenRecordingTestStream?.getVideoTracks()[0]?.dispatchEvent(new Event('ended'));
		});
		await expect(dock).toHaveCount(0);
		await expect(slideElements(page)).toHaveCount(before + 1);
	});
});

test('creates and downloads a separate Photo Album', async ({ page }) => {
	const count = await page.evaluate(() => window.__pptxViewer.getSlides().length);
	await page
		.getByRole('navigation', { name: 'Editor actions' })
		.getByRole('button', { name: 'Insert' })
		.tap();
	await page
		.locator('pptx-ui-mobile-insert')
		.getByRole('button', { name: 'Photo Album', exact: true })
		.tap();
	const dialog = picker(page, 'Photo Album');
	await dialog.getByRole('textbox', { name: 'Album title' }).fill('Field trip');
	const image = Buffer.from(
		await page.evaluate(() => {
			const canvas = document.createElement('canvas');
			canvas.width = 50;
			canvas.height = 40;
			const context = canvas.getContext('2d')!;
			context.fillStyle = '#0451a5';
			context.fillRect(0, 0, 50, 40);
			return canvas.toDataURL('image/png').split(',')[1]!;
		}),
		'base64',
	);
	await dialog.locator('input[type=file]').setInputFiles([
		{ name: 'Terrace.png', mimeType: 'image/png', buffer: image },
		{ name: 'Waterfall.png', mimeType: 'image/png', buffer: image },
	]);
	await dialog.getByRole('combobox', { name: 'Pictures per slide' }).click();
	await dialog.getByRole('option', { name: '2', exact: true }).click();
	await dialog.getByRole('checkbox', { name: 'Captions below pictures' }).check();
	const pending = page.waitForEvent('download');
	await dialog.getByRole('button', { name: 'Create Album', exact: true }).tap();
	const download = await pending;
	expect(download.suggestedFilename()).toBe('Field trip.pptx');
	const archive = await JSZip.loadAsync(await readFile(await download.path()));
	expect(await archive.file('ppt/slides/slide2.xml')?.async('string')).toContain('Waterfall');
	expect(await page.evaluate(() => window.__pptxViewer.getSlides().length)).toBe(count);
});

test('reuses only the selected slide from a local presentation', async ({ page }) => {
	const count = await page.evaluate(() => window.__pptxViewer.getSlides().length);
	await page
		.getByRole('navigation', { name: 'Editor actions' })
		.getByRole('button', { name: 'Insert' })
		.tap();
	await page
		.locator('pptx-ui-mobile-insert')
		.getByRole('button', { name: 'Reuse Slides', exact: true })
		.tap();
	const dialog = picker(page, 'Reuse Slides');
	await dialog.locator('input[type=file]').setInputFiles(SAMPLE_DECK);
	await dialog.getByRole('checkbox', { name: /^Slide 2\b/ }).check();
	await dialog.getByRole('button', { name: 'Insert', exact: true }).tap();
	await expect
		.poll(() => page.evaluate(() => window.__pptxViewer.getSlides().length))
		.toBe(count + 1);
	await expect(slideStage(page).getByText('Agenda', { exact: true })).toBeVisible();
});

test('adds local models and Office file objects from the phone Insert sheet', async ({ page }) => {
	const before = await slideElements(page).count();
	const source = await readFile(SAMPLE_DECK);
	const add = async (heading: string, file: { name: string; mimeType: string; buffer: Buffer }) => {
		await page
			.getByRole('navigation', { name: 'Editor actions' })
			.getByRole('button', { name: 'Insert' })
			.tap();
		await page
			.locator('pptx-ui-mobile-insert')
			.getByRole('button', { name: heading, exact: true })
			.tap();
		const dialog = picker(page, heading);
		const browse = dialog.getByRole('button', { name: 'Browse this device' });
		await expect(browse).toBeFocused();
		const chooser = page.waitForEvent('filechooser');
		await browse.press('Enter');
		await (await chooser).setFiles(file);
		await expect(dialog.getByRole('status')).toHaveText(file.name);
		await dialog.getByRole('button', { name: 'Insert', exact: true }).tap();
		await expect(dialog).toHaveCount(0);
	};
	await add('3D Models', {
		name: 'Triangle.glb',
		mimeType: 'model/gltf-binary',
		buffer: triangleGlb(),
	});
	await add('Object', {
		name: 'Agenda.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: source,
	});
	await expect(slideElements(page)).toHaveCount(before + 2);
	const pending = page.waitForEvent('download');
	await page
		.getByRole('toolbar', { name: 'Toolbar', exact: true })
		.getByRole('button', { name: 'Save', exact: true })
		.tap();
	const archive = await JSZip.loadAsync(await readFile(await (await pending).path()));
	const slides = await archive.file('ppt/slides/slide1.xml')?.async('string');
	expect(slides).toContain('<am3d:model3d');
	expect(slides).toContain('showAsIcon="1"');
	const embedding = Object.keys(archive.files).find((path) =>
		/^ppt\/embeddings\/oleObject\d+\.pptx$/.test(path),
	);
	expect(await archive.file(embedding!)?.async('nodebuffer')).toEqual(source);
});

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
	bytes.writeUInt32LE(0x4e4f534a, 16);
	json.copy(bytes, 20);
	bytes.writeUInt32LE(36, 20 + json.length);
	bytes.writeUInt32LE(0x004e4942, 24 + json.length);
	[0, 0, 0, 1, 0, 0, 0, 1, 0].forEach((value, i) =>
		bytes.writeFloatLE(value, 28 + json.length + i * 4),
	);
	return bytes;
}
