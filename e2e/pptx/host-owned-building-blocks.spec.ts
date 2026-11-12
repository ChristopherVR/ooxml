/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * Host-owned picker, dialog, find/replace and layout workflows of `useViewerBuildingBlocks` (#285),
 * driven through the React demo's custom shell (`HostOwnedWorkflows`). The building blocks are a
 * React-only API, so the other bindings skip.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { thumbnail } from './support/deck';
import { hostOwnedSessionUrl } from './support/host-owned-session';

test.beforeEach(() => {
	test.skip(test.info().project.name !== 'react', 'React-only building-block API');
});

const PNG = Buffer.from(
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
	'base64',
);

const ids = (page: Page) =>
	page
		.locator('[data-pptx-viewport] [data-element-id]')
		.evaluateAll((nodes) => [
			...new Set(nodes.map((node) => node.getAttribute('data-element-id') ?? '')),
		]);
const out = (page: Page, label: string) => page.getByLabel(label, { exact: true });
const click = (page: Page, name: string) => page.getByRole('button', { name, exact: true }).click();

async function open(page: Page, room: string, extra: Record<string, string> = {}): Promise<void> {
	await page.goto(hostOwnedSessionUrl(room, { headless: '1', sample: '1', ...extra }));
	await expect(page.locator('[data-host-workflows]')).toBeVisible({ timeout: 30_000 });
	await expect(page.locator('[data-pptx-viewport] [data-element-id]').first()).toBeVisible({
		timeout: 30_000,
	});
	if (extra.editable !== '0') {
		await expect(
			page.getByRole('status', { name: 'Collaboration status', exact: true }),
		).toContainText('Editable', { timeout: 30_000 });
	}
}

async function selectTitle(page: Page): Promise<void> {
	await page
		.locator('[data-pptx-viewport] [data-pptx-element="true"]')
		.filter({ hasText: 'Product Overview' })
		.click();
}

/**
 * The shell is in a collaboration room, so Undo runs the room's Yjs undo manager, which merges
 * edits made within its capture window into one step. Let it pass between two edits.
 */
const separateUndoSteps = (page: Page) => page.waitForTimeout(800);

async function undo(page: Page): Promise<void> {
	await click(page, 'Host undo');
}

test.describe('host-owned building-block workflows', () => {
	test('image: request, host picker, insert, undo', async ({ page }) => {
		await open(page, `bb-image-${Date.now()}`);
		const before = await ids(page);
		await click(page, 'Host picture');
		await expect(out(page, 'Host pending request')).toHaveText('image');
		await page.getByLabel('Host image file').setInputFiles({
			name: 'dot.png',
			mimeType: 'image/png',
			buffer: PNG,
		});
		await expect(out(page, 'Host last result')).toHaveText('image: ok');
		await expect(out(page, 'Host pending request')).toHaveText('none');
		expect((await ids(page)).length).toBe(before.length + 1);
		await undo(page);
		await expect.poll(async () => (await ids(page)).length).toBe(before.length);
	});

	test('image: validation codes for a wrong type', async ({ page }) => {
		await open(page, `bb-image-bad-${Date.now()}`);
		await click(page, 'Host picture');
		await page.getByLabel('Host image file').setInputFiles({
			name: 'a.txt',
			mimeType: 'text/plain',
			buffer: Buffer.from('hello'),
		});
		await expect(out(page, 'Host last result')).toHaveText('image: unsupported-type');
	});

	test('media: request, host picker, insert', async ({ page }) => {
		await open(page, `bb-media-${Date.now()}`);
		const before = await ids(page);
		await click(page, 'Host media');
		await expect(out(page, 'Host pending request')).toHaveText('media');
		await page.getByLabel('Host media file').setInputFiles({
			name: 'tone.mp3',
			mimeType: 'audio/mpeg',
			buffer: Buffer.from([0xff, 0xfb, 0x90, 0x00, 0x00, 0x00, 0x00, 0x00]),
		});
		await expect(out(page, 'Host last result')).toHaveText('media: ok');
		expect((await ids(page)).length).toBe(before.length + 1);
		await undo(page);
		await expect.poll(async () => (await ids(page)).length).toBe(before.length);
	});

	test('hyperlink: request, apply, current value, remove, undo', async ({ page }) => {
		await open(page, `bb-link-${Date.now()}`);
		await expect(out(page, 'Host action availability')).toContainText('hyperlink=no-selection');
		await selectTitle(page);
		await expect(out(page, 'Host action availability')).toContainText('hyperlink=yes*');
		await click(page, 'Host link');
		await expect(out(page, 'Host pending request')).toHaveText('hyperlink');
		await page.getByLabel('Host link URL').fill('https://example.test/atlas');
		await click(page, 'Apply link');
		await expect(out(page, 'Host last result')).toHaveText('link: ok');
		await expect(out(page, 'Host current link')).toHaveText('https://example.test/atlas');
		await separateUndoSteps(page);
		await click(page, 'Remove link');
		await expect(out(page, 'Host last result')).toHaveText('unlink: ok');
		await expect(out(page, 'Host current link')).toHaveText('none');
		await undo(page);
		// Undo clears the selection, as in the bundled viewer; select again to read the link back.
		await selectTitle(page);
		await expect(out(page, 'Host current link')).toHaveText('https://example.test/atlas');
	});

	test('SmartArt: request, insert, undo', async ({ page }) => {
		await open(page, `bb-smartart-${Date.now()}`);
		const before = await ids(page);
		await click(page, 'Host SmartArt');
		await expect(out(page, 'Host pending request')).toHaveText('smartArt');
		await page.getByLabel('Host SmartArt items').fill('Plan,Build,Ship');
		await click(page, 'Insert SmartArt');
		await expect(out(page, 'Host last result')).toHaveText('smartart: ok');
		await expect.poll(async () => (await ids(page)).length).toBe(before.length + 1);
		await undo(page);
		await expect.poll(async () => (await ids(page)).length).toBe(before.length);
	});

	test('equation: insert, current value, update', async ({ page }) => {
		await open(page, `bb-equation-${Date.now()}`);
		const before = await ids(page);
		await click(page, 'Host equation');
		await expect(out(page, 'Host pending request')).toHaveText('equation');
		await click(page, 'Insert equation');
		await expect(out(page, 'Host last result')).toHaveText('equation: ok');
		await expect.poll(async () => (await ids(page)).length).toBe(before.length + 1);
		await expect(out(page, 'Host current equation')).toContainText('x+1');
		await separateUndoSteps(page);
		await click(page, 'Update equation');
		await expect(out(page, 'Host current equation')).toContainText('y+2');
		await undo(page);
		const added = (await ids(page)).find((id) => !before.includes(id));
		await page.locator(`[data-pptx-viewport] [data-element-id="${added}"]`).first().click();
		await expect(out(page, 'Host current equation')).toContainText('x+1');
	});

	test('chart: request, insert of the chosen type with data, undo', async ({ page }) => {
		await open(page, `bb-chart-${Date.now()}`);
		const before = await ids(page);
		await expect(out(page, 'Host action availability')).toContainText('insertChart=yes*');
		await click(page, 'Host chart');
		await expect(out(page, 'Host pending request')).toHaveText('chart');
		await page.getByLabel('Host chart type').selectOption('pie');
		await click(page, 'Insert chart');
		await expect(out(page, 'Host last result')).toHaveText('chart: ok');
		await expect.poll(async () => (await ids(page)).length).toBe(before.length + 1);
		await expect(page.locator('[data-pptx-viewport]')).toContainText('Host chart');
		await undo(page);
		await expect.poll(async () => (await ids(page)).length).toBe(before.length);
	});

	test('find and replace: matches, navigation, replace all, undo', async ({ page }) => {
		await open(page, `bb-find-${Date.now()}`);
		await page.getByLabel('Host find', { exact: true }).fill('Atlas');
		await expect(out(page, 'Host match count')).not.toHaveText(/^0 /u);
		await page.getByLabel('Host replace', { exact: true }).fill('Orion');
		await click(page, 'Replace all');
		await expect(out(page, 'Host last result')).toHaveText('replace: ok');
		await expect(page.locator('[data-pptx-viewport]')).toContainText('Orion');
		await expect(page.locator('[data-pptx-viewport]')).not.toContainText('Atlas');
		await undo(page);
		await expect(page.locator('[data-pptx-viewport]')).toContainText('Atlas');
		await page.getByLabel('Host find', { exact: true }).fill('zzz-no-such-text');
		await click(page, 'Replace all');
		await expect(out(page, 'Host last result')).toHaveText('replace: no-match');
	});

	test('layout: current, apply, reset, unknown layout refused', async ({ page }) => {
		await open(page, `bb-layout-${Date.now()}`);
		const choice = page.getByLabel('Host layout choice');
		const options = await choice
			.locator('option')
			.evaluateAll((nodes) =>
				nodes.map((node) => (node as HTMLOptionElement).value).filter(Boolean),
			);
		expect(options.length).toBeGreaterThan(1);
		const current = await out(page, 'Host current layout').innerText();
		expect(current).not.toBe('none');
		await click(page, 'Apply layout');
		await expect(out(page, 'Host last result')).toHaveText('layout: unknown-layout');
		await choice.selectOption(options[options.length - 1]!);
		await click(page, 'Apply layout');
		await expect(out(page, 'Host last result')).toHaveText('layout: ok');
		await click(page, 'Reset layout');
		await expect(out(page, 'Host last result')).toHaveText('layout: ok');
	});

	test('read-only: requests are not raised and operations return read-only', async ({ page }) => {
		await open(page, `bb-readonly-${Date.now()}`, { editable: '0' });
		await expect(out(page, 'Host action availability')).toContainText('insertImage=read-only*');
		await expect(out(page, 'Host action availability')).toContainText('insertChart=read-only*');
		await click(page, 'Host picture');
		await click(page, 'Host chart');
		await expect(out(page, 'Host pending request')).toHaveText('none');
		await page.getByLabel('Host find', { exact: true }).fill('Atlas');
		await page.getByLabel('Host replace', { exact: true }).fill('Orion');
		await click(page, 'Replace all');
		await expect(out(page, 'Host last result')).toHaveText('replace: read-only');
		await expect(page.locator('[data-pptx-viewport]')).toContainText('Atlas');
		await click(page, 'Apply layout');
		await expect(out(page, 'Host last result')).toHaveText('layout: read-only');
	});

	test('stale target: switching slides while a picture decodes inserts nothing', async ({
		page,
	}) => {
		await open(page, `bb-stale-${Date.now()}`);
		await page.evaluate(() => {
			const read = FileReader.prototype.readAsDataURL;
			FileReader.prototype.readAsDataURL = function delayed(this: FileReader, blob: Blob) {
				setTimeout(() => read.call(this, blob), 800);
			};
		});
		const before = await ids(page);
		await click(page, 'Host picture');
		await page.getByLabel('Host image file').setInputFiles({
			name: 'dot.png',
			mimeType: 'image/png',
			buffer: PNG,
		});
		await thumbnail(page, 2).click();
		await expect(out(page, 'Host last result')).toHaveText('image: stale-target', {
			timeout: 10_000,
		});
		await thumbnail(page, 1).click();
		await expect.poll(async () => (await ids(page)).length).toBe(before.length);
	});
});
