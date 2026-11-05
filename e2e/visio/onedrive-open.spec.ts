/**
 * File > Open > OneDrive in every demo binding, against a test-only Graph stand-in: sign in through
 * the demo dialog, open a drawing, edit it through the public API and File > Save to the same
 * item. A failed upload stays unsaved; a drawing on another drive opens read-only.
 */
import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { openDemo } from './demo-page';
import { createVsdxFixture } from './fixture.mjs';
import { ONEDRIVE_TEST_TOKEN, routeOneDriveGraph, type GraphSeed } from './support/onedrive-graph';
import { downloadCopy, fileBackstage } from './ribbon';

type VisioViewerElement = HTMLElementTagNameMap['visio-viewer'];

const EDITED = 'Edited in the cloud';
const paths: Record<string, string> = {
	vanilla: '/demo/?sample=1',
	react: '/demo-react/?sample=1',
	vue: '/demo-vue/?sample=1',
	angular: '/demo-angular/?sample=1',
	svelte: '/demo-svelte/?sample=1',
	solid: '/demo-solid/?sample=1',
};

async function seeds(extra: Record<string, GraphSeed> = {}): Promise<Record<string, GraphSeed>> {
	return {
		reports: { name: 'Reports', folder: true },
		mine: { name: 'Plan.vsdx', bytes: await createVsdxFixture('Cloud drawing'), parent: 'reports' },
		...extra,
	};
}

/** File > Open > OneDrive, signed in through the demo dialog with the test token. */
async function showCloud(page: Page): Promise<void> {
	const viewer = page.locator('visio-viewer');
	await fileBackstage(viewer, 'open');
	await viewer.locator('office-ui-open-places').getByRole('tab', { name: 'OneDrive' }).click();
	await viewer.getByRole('button', { name: 'Sign in', exact: true }).click();
	await page.getByLabel('Access token').fill(ONEDRIVE_TEST_TOKEN);
	await page.getByRole('button', { name: 'Continue', exact: true }).click();
	await expect(viewer.getByRole('button', { name: 'Recent', exact: true })).toBeVisible();
}

async function openFile(page: Page, name: string, via: 'folder' | 'recent'): Promise<void> {
	const viewer = page.locator('visio-viewer');
	await showCloud(page);
	if (via === 'folder') {
		await viewer.getByRole('button', { name: 'My files', exact: true }).click();
		await viewer.getByRole('option', { name: /Reports/u }).click();
	}
	await viewer.getByRole('option', { name: new RegExp(name.replace('.', '\\.'), 'u') }).click();
	await expect
		.poll(
			() =>
				page.evaluate(
					() => (document.querySelector('visio-viewer') as VisioViewerElement).fileName,
				),
			{
				timeout: 30_000,
			},
		)
		.toBe(name);
	await expect(viewer.locator('svg.paper text')).toContainText('Cloud drawing');
}

const edit = (page: Page) =>
	page.evaluate(async (text) => {
		const element = document.querySelector('visio-viewer') as VisioViewerElement;
		await element.controller.replacePlainText('1', '1', text);
	}, EDITED);
const dirty = (page: Page) =>
	page.evaluate(
		() =>
			(document.querySelector('visio-viewer') as VisioViewerElement).controller.state.edit.dirty,
	);
const save = async (page: Page) => (await downloadCopy(page.locator('visio-viewer'))).click();

for (const [framework, path] of Object.entries(paths)) {
	test(`${framework}: opens a OneDrive drawing and File > Save uploads to the same item`, async ({
		page,
	}) => {
		const cloud = await routeOneDriveGraph(page, await seeds());
		await openDemo(page, path);
		await openFile(page, 'Plan.vsdx', 'folder');
		expect(await dirty(page)).toBe(false);
		await edit(page);
		expect(await dirty(page)).toBe(true);
		await save(page);
		await expect.poll(() => cloud.writes).toStrictEqual(['mine']);
		await expect.poll(() => dirty(page)).toBe(false);
		const saved = await JSZip.loadAsync(cloud.items.get('mine')!.bytes);
		expect(await saved.file('visio/pages/page1.xml')!.async('string')).toContain(EDITED);
	});
}

test('a failed upload keeps the drawing unsaved', async ({ page }) => {
	const cloud = await routeOneDriveGraph(page, await seeds());
	await openDemo(page, paths.vanilla);
	await openFile(page, 'Plan.vsdx', 'folder');
	await edit(page);
	await page.evaluate(() => {
		const w = window as unknown as { saveErrors: number };
		w.saveErrors = 0;
		document
			.querySelector('visio-viewer')!
			.addEventListener('document-error', () => (w.saveErrors += 1));
	});
	cloud.failUploads(503);
	await save(page);
	// The failed save is reported as a document-error.
	await expect
		.poll(() => page.evaluate(() => (window as unknown as { saveErrors: number }).saveErrors))
		.toBe(1);
	await expect.poll(() => cloud.writes.length).toBe(0);
	expect(await dirty(page)).toBe(true);
	// Once the service answers again, the same Save succeeds and marks the drawing clean.
	cloud.failUploads(0);
	await save(page);
	await expect.poll(() => cloud.writes).toStrictEqual(['mine']);
	await expect.poll(() => dirty(page)).toBe(false);
});

test('a drawing on another drive opens read-only', async ({ page }) => {
	await routeOneDriveGraph(
		page,
		await seeds({
			shared: {
				name: 'Shared with me.vsdx',
				bytes: await createVsdxFixture('Cloud drawing'),
				driveId: 'colleague-drive',
			},
		}),
	);
	await openDemo(page, paths.vanilla);
	await openFile(page, 'Shared with me.vsdx', 'recent');
	expect(
		await page.evaluate(
			() => (document.querySelector('visio-viewer') as VisioViewerElement).canEdit,
		),
	).toBe(false);
	await expect(page.locator('visio-viewer')).toContainText('read-only');
});

test('This device keeps the plain file chooser next to the cloud place', async ({ page }) => {
	await routeOneDriveGraph(page, await seeds());
	await openDemo(page, paths.vanilla);
	const viewer = page.locator('visio-viewer');
	await fileBackstage(viewer, 'open');
	await expect(
		viewer.locator('office-ui-open-places').getByRole('tab', { name: 'This device' }),
	).toHaveAttribute('aria-selected', 'true');
	await expect(
		viewer.locator('[data-backstage-page="open"] [data-backstage-action="open"]'),
	).toBeVisible();
	await viewer.locator('office-ui-open-places').getByRole('tab', { name: 'OneDrive' }).click();
	await expect(
		viewer.locator('[data-backstage-page="open"] [data-backstage-action="open"]'),
	).toBeHidden();
});

for (const [framework, path] of Object.entries(paths))
	test(`${framework}: Open from OneDrive on the start screen opens a drawing and saves to the same item`, async ({
		page,
	}) => {
		const cloud = await routeOneDriveGraph(page, await seeds());
		await page.goto(path.replace('?sample=1', ''));
		await page.locator('#start-onedrive').click();
		const dialog = page.getByRole('dialog', { name: 'Open from OneDrive' });
		await dialog.getByRole('button', { name: 'Sign in', exact: true }).click();
		await page.getByLabel('Access token').fill(ONEDRIVE_TEST_TOKEN);
		await page.getByRole('button', { name: 'Continue', exact: true }).click();
		await dialog.getByRole('button', { name: 'My files', exact: true }).click();
		await dialog.getByRole('option', { name: /Reports/u }).click();
		await dialog.getByRole('option', { name: /Plan\.vsdx/u }).click();
		await expect(dialog).toBeHidden();
		await expect
			.poll(
				() =>
					page.evaluate(
						() => (document.querySelector('visio-viewer') as VisioViewerElement).fileName,
					),
				{ timeout: 30_000 },
			)
			.toBe('Plan.vsdx');
		await expect(page.locator('visio-viewer').locator('svg.paper text')).toContainText(
			'Cloud drawing',
		);
		await edit(page);
		await save(page);
		await expect.poll(() => cloud.writes).toStrictEqual(['mine']);
		await expect.poll(() => dirty(page)).toBe(false);
	});
