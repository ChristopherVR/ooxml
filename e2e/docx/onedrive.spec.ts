/**
 * File > Open > OneDrive in the demo, against a test-only Graph stand-in: sign in through the demo
 * dialog with a token, open a file, edit and File > Save uploads to the same item. The document
 * reads saved only once the upload landed; a failed upload leaves it unsaved; a file on another
 * drive opens read-only.
 */
import { expect, test, type Locator, type Page } from '@playwright/test';
import { clickInEditor, openSample, saveButton, saveStateLabel, fileNameLabel } from './helpers';
import { savedDocument } from './saved-document';
import { ONEDRIVE_TEST_TOKEN, routeOneDriveGraph, type DoubleFile } from './support/onedrive-graph';

const editor = (page: Page) => page.locator('docx-editor');
const cloud = (page: Page): Locator => editor(page).locator('office-ui-cloud-files');

/** The sample document's own bytes, so the stand-in serves a real .docx. */
async function sampleBytes(page: Page): Promise<Buffer> {
	return Buffer.from((await savedDocument(page, 'sample.docx')).bytes);
}

async function openOneDrivePlace(page: Page): Promise<void> {
	await editor(page).locator('.dve-ribbon .file').click();
	await editor(page)
		.locator('.dve-backstage .item', { hasText: /^Open$/ })
		.click();
	const places = editor(page).locator('office-ui-open-places');
	await expect(places.getByRole('tab')).toHaveText(['This device', 'OneDrive']);
	await places.getByRole('tab', { name: 'OneDrive' }).click();
}

async function signIn(page: Page): Promise<void> {
	await cloud(page).getByRole('button', { name: 'Sign in', exact: true }).click();
	const dialog = page.getByRole('dialog', { name: 'Connect OneDrive (demo)' });
	await dialog.getByLabel('Access token').fill(ONEDRIVE_TEST_TOKEN);
	await dialog.getByRole('button', { name: 'Continue', exact: true }).click();
	await expect(cloud(page).locator('[data-view="recent"]')).toBeVisible();
}

async function openFromRecent(page: Page, name: string): Promise<void> {
	await openOneDrivePlace(page);
	await signIn(page);
	await cloud(page).locator('[data-view="recent"]').click();
	await cloud(page)
		.getByRole('option', { name: new RegExp(name) })
		.click();
	await expect(fileNameLabel(page)).toHaveText(name, { timeout: 30_000 });
	await expect(editor(page).locator('.dve-backstage')).toBeHidden();
}

/** Opens the sample once to get real bytes, then starts over with the stand-in routed. */
async function prepare(
	page: Page,
	framework: string,
	files: (bytes: Buffer) => Record<string, DoubleFile>,
) {
	await openSample(page, framework);
	const bytes = await sampleBytes(page);
	const graph = await routeOneDriveGraph(page, files(bytes));
	return graph;
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: opens from OneDrive, edits and saves to the same item`, async ({ page }) => {
		let original = Buffer.alloc(0);
		const graph = await prepare(page, framework, (bytes) => {
			original = bytes;
			return { mine: { name: 'Plan.docx', bytes } };
		});
		await openFromRecent(page, 'Plan.docx');
		await expect(editor(page).locator('.ProseMirror')).toContainText('Document title');
		await expect(saveStateLabel(page)).toHaveText('Saved');
		await clickInEditor(editor(page).locator('.ProseMirror p').last());
		await page.keyboard.press('End');
		await page.keyboard.type(' Edited in the cloud.');
		await expect(saveStateLabel(page)).toHaveText('Unsaved changes');
		await saveButton(page).click();
		await expect(saveStateLabel(page)).toHaveText('Saved', { timeout: 30_000 });
		expect(graph.writes).toStrictEqual(['mine']);
		expect(graph.files.get('mine')!.bytes.equals(original)).toBe(false);
	});
}

test('a failed upload keeps the document unsaved', async ({ page }) => {
	const graph = await prepare(page, 'vanilla', (bytes) => ({ mine: { name: 'Plan.docx', bytes } }));
	await openFromRecent(page, 'Plan.docx');
	await clickInEditor(editor(page).locator('.ProseMirror p').last());
	await page.keyboard.press('End');
	await page.keyboard.type(' Not saved.');
	graph.failUploads(503);
	await saveButton(page).click();
	await expect(saveStateLabel(page)).toHaveText('Unsaved changes', { timeout: 30_000 });
	expect(graph.writes).toStrictEqual([]);
});

test('a file on another drive opens read-only', async ({ page }) => {
	await prepare(page, 'vanilla', (bytes) => ({
		shared: { name: 'Shared with me.docx', bytes, driveId: 'colleague-drive' },
	}));
	await openFromRecent(page, 'Shared with me.docx');
	await expect(editor(page)).toHaveJSProperty('readOnly', true);
});

test('browses into a folder and opens a file there', async ({ page }) => {
	await openSample(page, 'vanilla');
	const bytes = await sampleBytes(page);
	await routeOneDriveGraph(
		page,
		{ nested: { name: 'Nested.docx', bytes, folder: 'reports' } },
		{ reports: 'Reports' },
	);
	await openOneDrivePlace(page);
	await signIn(page);
	await cloud(page).locator('[data-view="files"]').click();
	await cloud(page)
		.getByRole('option', { name: /Reports/ })
		.click();
	await cloud(page)
		.getByRole('option', { name: /Nested\.docx/ })
		.click();
	await expect(fileNameLabel(page)).toHaveText('Nested.docx', { timeout: 30_000 });
});

test('This device keeps the Browse button', async ({ page }) => {
	await openSample(page, 'vanilla');
	await editor(page).locator('.dve-ribbon .file').click();
	await editor(page)
		.locator('.dve-backstage .item', { hasText: /^Open$/ })
		.click();
	await expect(editor(page).getByRole('button', { name: 'Browse…' })).toBeVisible();
});
