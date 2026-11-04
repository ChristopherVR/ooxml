import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const require = createRequire(new URL('../../e2e/docx/package.json', import.meta.url));
const { expect } = require('@playwright/test');
const { inspectDocx, setDocxRunText } = await import('../../src/core/dist/automation/index.mjs');

/** Test-only backend exercises the actual provider extension through the built suite. */
export async function checkCustomStorage({ browser, base, bytes }) {
	const context = await browser.newContext({ serviceWorkers: 'block' });
	let data = Buffer.from(bytes),
		revision = 1,
		writes = 0;
	const grants = [];
	await context.route('**/host-files/report', async (route) => {
		const request = route.request();
		if (request.method() === 'PUT') {
			if (request.headers()['if-match'] !== String(revision)) {
				await route.fulfill({ status: 412 });
				return;
			}
			data = request.postDataBuffer();
			revision++;
			writes++;
		}
		await route.fulfill({ status: 200, headers: { 'x-revision': String(revision) }, body: data });
	});
	await context.route('**/host-files/share', async (route) => {
		grants.push(route.request().postDataJSON());
		await route.fulfill({ status: 200, body: 'https://company.example/view/report' });
	});
	await context.addInitScript(() => {
		globalThis.ooxmlSkipStart = true;
		const file = (revision) => ({
			key: 'report',
			revision,
			name: 'Company.docx',
			kind: 'docx',
			access: 'write',
			webUrl: 'https://company.example/report',
		});
		const read = async () => {
			const response = await fetch('/host-files/report');
			return {
				file: file(Number(response.headers.get('x-revision'))),
				bytes: new Uint8Array(await response.arrayBuffer()),
			};
		};
		globalThis.ooxmlCloudOptions = {
			id: 'company',
			label: 'Company Files',
			keyOf: (value) => value.key,
			serializeFile: ({ key, revision, name, kind, access }) => ({
				key,
				revision,
				name,
				kind,
				access,
			}),
			client: {
				open: read,
				read,
				async save(value, bytes, options) {
					const response = await fetch('/host-files/report', {
						method: 'PUT',
						headers: { 'If-Match': String(value.revision) },
						body: bytes,
						signal: options?.signal,
					});
					if (response.status === 412)
						throw Object.assign(new Error('stale'), { code: 'stale-revision' });
					if (!response.ok) throw new Error('save failed');
					return file(Number(response.headers.get('x-revision')));
				},
				isConflict: (error) => error?.code === 'stale-revision',
				async share(_file, options) {
					return (
						await fetch('/host-files/share', { method: 'POST', body: JSON.stringify(options) })
					).text();
				},
			},
			maxMergeAttempts: 2,
			shareChoices: [
				{ label: 'Project viewers', options: { audience: 'project', role: 'reader' } },
			],
		};
	});
	const page = await context.newPage();
	const errors = [];
	page.on('pageerror', (error) => errors.push(error.message));
	try {
		await page.goto(`${base}/#/`);
		await page.getByRole('button', { name: 'Open from Company Files', exact: true }).click();
		await page.locator('#onedrive-link').fill('company:report');
		await page.locator('#onedrive-open-form button').click();
		await expect(page.locator('#save-state')).toHaveText('Saved to Company Files');
		const editor = page.locator('.document-host:not([hidden]) docx-editor');
		await editor.locator('[contenteditable=true]').first().click();
		await page.keyboard.press('Control+End');
		await page.keyboard.type(' Custom provider local edit.');
		const paragraph = (await inspectDocx(data)).blocks.find((block) => block.type === 'paragraph');
		data = Buffer.from(
			(
				await setDocxRunText(
					data,
					paragraph.id,
					0,
					`Remote company edit: ${paragraph.runs[0].text}`,
				)
			).bytes,
		);
		revision++;
		await page.locator('#save').click();
		await expect.poll(() => writes).toBe(1);
		await expect(page.locator('#save-state')).toHaveText('Saved to Company Files');
		const saved = JSON.stringify(await inspectDocx(data));
		assert.match(saved, /Custom provider local edit/u);
		assert.match(saved, /Remote company edit/u);
		await expect(editor).toContainText('Remote company edit');
		await page.locator('#share').click();
		await expect(page.locator('#onedrive-share-form select')).toHaveCount(1);
		await page.locator('#onedrive-share-form button').click();
		await expect(page.locator('#onedrive-share-link')).toHaveValue(
			'https://company.example/view/report',
		);
		assert.deepEqual(grants, [{ audience: 'project', role: 'reader' }]);
		assert.deepEqual(errors, []);
		console.log(
			'Custom provider: host identity/revision, concurrent merge, editor adoption and explicit sharing passed',
		);
	} finally {
		await context.close();
	}
}
