import { expect, test, type Page, type Route } from '@playwright/test';
import { createWorkbook, createEditSession, getCell, loadXlsx, saveXlsx } from 'ooxml-core/xlsx';
import {
	FRAMEWORKS,
	editor,
	editorProperty,
	goToCell,
	newWorkbook,
	openLanding,
	pageErrors,
	typeInActiveCell,
} from './helpers';

/**
 * File > Open > OneDrive against a test-only Microsoft Graph double (Playwright routing). It
 * serves a small drive, a second drive the signed-in user cannot write, downloads and upload
 * sessions, and can be told to fail uploads.
 */
const OWN = 'drive-own';
const OTHER = 'drive-other';

interface Item {
	drive: string;
	id: string;
	name: string;
	parent: string;
	bytes?: Uint8Array;
	eTag: string;
}

async function workbookBytes(sheet: string): Promise<Uint8Array> {
	const book = createWorkbook({ sheets: [sheet] });
	createEditSession(book).setCellInput(0, 0, 0, 'Hello');
	return saveXlsx(book);
}

async function graphDouble(page: Page) {
	const items: Item[] = [
		{
			drive: OWN,
			id: 'item-plan',
			name: 'Plan.xlsx',
			parent: 'root-own',
			bytes: await workbookBytes('Cloud plan'),
			eTag: '"e1"',
		},
		{ drive: OWN, id: 'item-reports', name: 'Reports', parent: 'root-own', eTag: '"f1"' },
		{
			drive: OWN,
			id: 'item-budget',
			name: 'Budget.xlsx',
			parent: 'item-reports',
			bytes: await workbookBytes('Budget'),
			eTag: '"b1"',
		},
		{
			drive: OTHER,
			id: 'item-shared',
			name: 'Shared.xlsx',
			parent: 'root-other',
			bytes: await workbookBytes('Shared sheet'),
			eTag: '"s1"',
		},
	];
	const state = {
		failUploads: false,
		uploads: [] as Array<{ drive: string; id: string; bytes: Uint8Array }>,
		requests: [] as string[],
	};
	const cors = {
		'access-control-allow-origin': '*',
		'access-control-allow-headers': '*',
		'access-control-allow-methods': '*',
	};
	const json = (route: Route, body: unknown, status = 200) =>
		route.fulfill({
			status,
			headers: { ...cors, 'content-type': 'application/json' },
			body: JSON.stringify(body),
		});
	const metadata = (item: Item, withDownload = false) => ({
		id: item.id,
		name: item.name,
		eTag: item.eTag,
		size: item.bytes?.length,
		webUrl: `https://onedrive.test/${item.id}`,
		parentReference: { driveId: item.drive, id: item.parent, path: '/drive/root:' },
		lastModifiedDateTime: '2026-10-01T10:00:00Z',
		...(item.bytes
			? { file: { mimeType: 'application/octet-stream' } }
			: { folder: { childCount: 1 } }),
		...(withDownload
			? { '@microsoft.graph.downloadUrl': `https://download.test/${item.drive}/${item.id}` }
			: {}),
	});
	const find = (drive: string, id: string) =>
		items.find((item) => item.drive === drive && item.id === id);

	await page.route('https://graph.microsoft.com/**', async (route) => {
		const request = route.request();
		if (request.method() === 'OPTIONS') return route.fulfill({ status: 204, headers: cors });
		const url = new URL(request.url());
		const path = decodeURIComponent(url.pathname.replace('/v1.0', ''));
		state.requests.push(`${request.method()} ${path}${url.search}`);
		if (path === '/me/drive') return json(route, { id: OWN });
		if (path === '/me/drive/root/children')
			return json(route, {
				value: items.filter((item) => item.parent === 'root-own').map((item) => metadata(item)),
			});
		if (path === '/me/drive/recent')
			return json(route, {
				value: items.filter((item) => item.drive === OTHER).map((item) => metadata(item)),
			});
		const children = /^\/drives\/([^/]+)\/items\/([^/]+)\/children$/u.exec(path);
		if (children)
			return json(route, {
				value: items.filter((item) => item.parent === children[2]).map((item) => metadata(item)),
			});
		const session = /^\/drives\/([^/]+)\/items\/([^/]+)\/createUploadSession$/u.exec(path);
		if (session) {
			const target = find(session[1]!, session[2]!);
			const match = request.headers()['if-match'];
			if (!target) return json(route, {}, 404);
			if (match && match !== target.eTag) return json(route, {}, 412);
			return json(route, { uploadUrl: `https://upload.test/${target.drive}/${target.id}` });
		}
		const item = /^\/drives\/([^/]+)\/items\/([^/]+)$/u.exec(path);
		const target = item && find(item[1]!, item[2]!);
		if (target) return json(route, metadata(target, true));
		return json(route, { error: { code: 'itemNotFound' } }, 404);
	});
	await page.route('https://download.test/**', async (route) => {
		if (route.request().method() === 'OPTIONS')
			return route.fulfill({ status: 204, headers: cors });
		const [, drive, id] = new URL(route.request().url()).pathname.split('/');
		const target = find(drive!, id!);
		return route.fulfill({
			status: 200,
			headers: { ...cors, 'content-type': 'application/octet-stream' },
			body: Buffer.from(target!.bytes!),
		});
	});
	await page.route('https://upload.test/**', async (route) => {
		if (route.request().method() === 'OPTIONS')
			return route.fulfill({ status: 204, headers: cors });
		const [, drive, id] = new URL(route.request().url()).pathname.split('/');
		if (state.failUploads) return json(route, { error: { code: 'generalException' } }, 500);
		const target = find(drive!, id!)!;
		const bytes = new Uint8Array(route.request().postDataBuffer()!);
		target.bytes = bytes;
		target.eTag = `"e${state.uploads.length + 2}"`;
		state.uploads.push({ drive: drive!, id: id!, bytes });
		return json(route, metadata(target));
	});
	return state;
}

async function openCloudPlace(page: Page) {
	await editor(page).getByRole('button', { name: 'File', exact: true }).click();
	await editor(page).locator('[data-backstage-item="open"]').click();
	await editor(page).locator('.xve-backstage-place-tab', { hasText: 'OneDrive' }).click();
}

async function signIn(page: Page) {
	await editor(page).getByRole('button', { name: 'Sign in', exact: true }).first().click();
	await page.getByLabel('Access token').fill('test-token');
	await page.getByRole('button', { name: 'Continue' }).click();
	await expect(editor(page).locator('office-ui-cloud-files [data-entry]').first()).toBeVisible();
}

test('File > Open > OneDrive opens, saves back to the item, keeps failed saves dirty, and opens other drives read-only', async ({
	page,
}) => {
	const errors = pageErrors(page);
	const graph = await graphDouble(page);
	await newWorkbook(page);
	await openCloudPlace(page);
	await signIn(page);

	const cloud = editor(page).locator('office-ui-cloud-files');
	await cloud.locator('[data-view="files"]').click();
	await expect(cloud.locator('[data-entry]')).toHaveCount(2);
	await cloud.locator('[data-entry]', { hasText: 'Plan.xlsx' }).click();
	await expect.poll(() => editorProperty<string>(page, 'fileName')).toBe('Plan.xlsx');
	await expect(editor(page).locator('[part="sheet-tabs"]')).toContainText('Cloud plan');
	expect(await editorProperty<boolean>(page, 'readOnly')).toBe(false);

	// An edit saved with Ctrl+S goes to the same OneDrive item and then marks the workbook clean.
	await goToCell(page, 'B2');
	await typeInActiveCell(page, '42');
	await expect.poll(() => editorProperty<boolean>(page, 'dirty')).toBe(true);
	await page.keyboard.press('Control+s');
	await expect.poll(() => graph.uploads.length).toBe(1);
	expect(graph.uploads[0]).toMatchObject({ drive: OWN, id: 'item-plan' });
	const saved = await loadXlsx(graph.uploads[0]!.bytes);
	expect(getCell(saved.sheets[0]!, 1, 1)?.value).toBe(42);
	await expect.poll(() => editorProperty<boolean>(page, 'dirty')).toBe(false);

	// A failed upload stays unsaved and reports the error.
	graph.failUploads = true;
	await goToCell(page, 'B3');
	await typeInActiveCell(page, '7');
	await page.keyboard.press('Control+s');
	await expect(editor(page)).toContainText('OneDrive request failed');
	expect(await editorProperty<boolean>(page, 'dirty')).toBe(true);
	expect(graph.uploads).toHaveLength(1);
	graph.failUploads = false;
	await page.keyboard.press('Control+s');
	await expect.poll(() => graph.uploads.length).toBe(2);
	await expect.poll(() => editorProperty<boolean>(page, 'dirty')).toBe(false);

	// A file on another drive opens read-only (the demo allows edits on the user's own drive only).
	await openCloudPlace(page);
	await editor(page).locator('office-ui-cloud-files [data-view="recent"]').click();
	await editor(page)
		.locator('office-ui-cloud-files [data-entry]', { hasText: 'Shared.xlsx' })
		.click();
	await expect.poll(() => editorProperty<string>(page, 'fileName')).toBe('Shared.xlsx');
	await expect.poll(() => editorProperty<boolean>(page, 'readOnly')).toBe(true);
	await expect(editor(page).locator('[part="sheet-tabs"]')).toContainText('Shared sheet');
	expect(errors).toEqual([]);
});

for (const framework of FRAMEWORKS)
	test(`File > Open offers the OneDrive place in ${framework}`, async ({ page }) => {
		await openLanding(page, framework);
		await page.locator('#blank').click();
		await editor(page).getByRole('button', { name: 'File', exact: true }).click();
		await editor(page).locator('[data-backstage-item="open"]').click();
		const tabs = editor(page).locator('.xve-backstage-place-tab');
		await expect(tabs).toHaveText(['This device', 'OneDrive']);
		await tabs.nth(1).click();
		await expect(editor(page).locator('office-ui-cloud-files')).toBeVisible();
	});
