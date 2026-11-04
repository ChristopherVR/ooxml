/**
 * Screenshots for docs/onedrive-quickstart.md. Microsoft Graph is answered by the test-only HTTP
 * double the suite's browser check uses, so no Microsoft account is involved.
 *
 *   bun run build:suite && node scripts/preview-site.mjs 8151
 *   node scripts/onedrive-doc-shots.mjs suite
 *
 * The React shots need the guide's example served by a Vite app whose OneDrive client returns
 * 'test-only-token' from getAccessToken: REACT_URL=http://localhost:4190 node scripts/onedrive-doc-shots.mjs react
 *
 * The PowerPoint demo's OneDrive mode (`?onedrive=1`), from a running React demo:
 *   DEMO_URL=http://localhost:4173 node scripts/onedrive-doc-shots.mjs demo
 */
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { interceptOneDrive } from './test-support/onedrive-http.mjs';

const require = createRequire(new URL('../e2e/docx/package.json', import.meta.url));
const { chromium, expect } = require('@playwright/test');
const { createDocx } = await import('../src/core/dist/automation/index.mjs');
const { createPresentationFromOutline } = await import('../src/core/dist/pptx/index.mjs');

const SUITE = process.env.SUITE_URL ?? 'http://127.0.0.1:8151';
const REACT = process.env.REACT_URL;
const DEMO = process.env.DEMO_URL;
const out = (name) => fileURLToPath(new URL(`../docs/images/onedrive/${name}`, import.meta.url));
const link = (kind) => `https://contoso-my.sharepoint.com/:p:/g/personal/${kind}`;
const only = process.argv[2];

const deck = () =>
	createPresentationFromOutline(
		[
			{
				title: 'Quarterly review',
				bullets: ['Revenue up 12% on last quarter', 'Two new regions launched', 'Hiring on plan'],
			},
			{ title: 'Next steps', bullets: ['Finalise the budget', 'Confirm the launch dates'] },
		],
		{ title: 'Quarterly review' },
	);
async function fixtures() {
	return { pptx: await deck(), docx: await createDocx(['Project plan.']) };
}
function rename(cloud) {
	cloud.files.get('pptx').name = 'Quarterly review.pptx';
	cloud.files.get('docx').name = 'Project plan.docx';
}

const browser = await chromium.launch();
const viewport = { width: 1280, height: 800 };
const errors = [];
async function newPage({ configured }) {
	const context = await browser.newContext({ viewport, serviceWorkers: 'block' });
	await context.addInitScript((configured) => {
		globalThis.ooxmlSkipStart = true;
		if (configured)
			globalThis.ooxmlOneDriveOptions = {
				getAccessToken: async () => 'test-only-token',
				canWrite: () => true,
			};
	}, configured);
	const cloud = await interceptOneDrive(context, await fixtures());
	rename(cloud);
	const page = await context.newPage();
	page.setDefaultTimeout(60000);
	page.on('pageerror', (error) => errors.push(error.message));
	return { context, page, cloud };
}

try {
	if (!only || only === 'suite') {
		const { context, page, cloud } = await newPage({ configured: true });
		await page.goto(`${SUITE}/#/`);
		await expect(page.locator('#open-onedrive')).toBeVisible();
		await page.locator('#open-onedrive').scrollIntoViewIfNeeded();
		await page.waitForTimeout(2500);
		await page.screenshot({ path: out('suite-home.png') });
		await page.locator('#open-onedrive').click();
		await page.locator('#onedrive-link').fill(link('pptx'));
		await page.locator('#dialog').screenshot({ path: out('suite-open-dialog.png') });
		await page.locator('#onedrive-open-submit').click();
		await expect(page.locator('#dialog')).toBeHidden();
		await expect(page.locator('#save-state')).toContainText('OneDrive');
		const host = page.locator('.document-host:not([hidden])');
		await host.getByRole('button', { name: 'New Slide', exact: true }).click();
		await host.getByRole('button', { name: 'Save', exact: true }).first().click();
		await expect(page.locator('#save-state')).toContainText('Saved to OneDrive');
		await host.getByText('Quarterly review').first().click();
		await page.waitForTimeout(1500);
		await page.screenshot({ path: out('suite-powerpoint-saved.png') });

		await page.locator('#rail [data-home]').click();
		await page.locator('#open-onedrive').click();
		await page.locator('#onedrive-link').fill(link('docx'));
		await page.locator('#onedrive-open-submit').click();
		await expect(page.locator('#dialog')).toBeHidden();
		await page
			.locator('.document-host:not([hidden]) docx-editor [contenteditable=true]')
			.first()
			.click();
		await page.keyboard.press('Control+A');
		await page.keyboard.type('Plan, local version.');
		const remote = cloud.files.get('docx');
		remote.bytes = Buffer.from(await createDocx(['Plan, version saved by a colleague.']));
		remote.revision++;
		await page.locator('#save').click();
		await expect(page.locator('#onedrive-conflict')).toBeVisible();
		await page.locator('#dialog').screenshot({ path: out('suite-conflict.png') });
		await context.close();

		const plain = await newPage({ configured: false });
		await plain.page.goto(`${SUITE}/#/`);
		await plain.page.locator('#open-onedrive').click();
		await plain.page.locator('#dialog').screenshot({ path: out('suite-not-configured.png') });
		await plain.context.close();
		console.log('suite screenshots written');
	}

	if (only === 'react' || (!only && REACT)) {
		if (!REACT) throw new Error('Set REACT_URL to the served React example.');
		const { context, page, cloud } = await newPage({ configured: false });
		await page.goto(REACT);
		await page.getByLabel('OneDrive sharing link').fill(link('pptx'));
		await page.screenshot({ path: out('react-open.png') });
		await page.getByRole('button', { name: 'Open', exact: true }).click();
		await expect(page.locator('header [role=status]')).toHaveText('Opened from OneDrive');
		await expect(page.getByText('Quarterly review').first()).toBeVisible();
		await page.waitForTimeout(1500);
		await page.screenshot({ path: out('react-editor.png') });

		await page.getByRole('button', { name: 'New Slide', exact: true }).click();
		await page.getByRole('button', { name: 'Save', exact: true }).first().click();
		await expect(page.locator('header [role=status]')).toHaveText('Saved to OneDrive');
		await expect.poll(() => cloud.writes.filter((id) => id === 'pptx').length).toBe(1);
		// The status bar still reads "Unsaved changes" here (a binding bug), so no screenshot.
		await context.close();
		console.log('react screenshots written; File > Save uploaded to the same item');
	}
	if (only === 'demo' || (!only && DEMO)) {
		if (!DEMO) throw new Error('Set DEMO_URL to the running React demo.');
		const { context, page, cloud } = await newPage({ configured: false });
		await page.goto(`${DEMO}/?onedrive=1`);
		const bar = page.locator('.od-bar');
		await bar.getByLabel('How to sign in').selectOption('token');
		await bar.getByLabel('Access token').fill('test-only-token');
		await bar.getByRole('button', { name: 'Use token', exact: true }).click();
		await bar.getByLabel('OneDrive sharing link').fill(link('pptx'));
		await bar.getByRole('button', { name: 'Open', exact: true }).click();
		await expect(bar.getByRole('status')).toContainText('Opened');
		await expect(
			page.locator('[data-pptx-viewport] [data-pptx-element="true"]').first(),
		).toBeVisible({ timeout: 30000 });
		await page
			.getByRole('button', { name: /new slide/iu })
			.first()
			.click();
		await page
			.getByRole('toolbar', { name: 'Presentation toolbar' })
			.getByRole('tab', { name: 'File', exact: true })
			.click();
		await page
			.getByRole('dialog', { name: 'File' })
			.getByRole('button', { name: 'Save', exact: true })
			.click();
		await expect(bar.getByRole('status')).toContainText('Saved to OneDrive');
		await expect.poll(() => cloud.writes.filter((id) => id === 'pptx').length).toBe(1);
		await page.getByText('Quarterly review').first().click();
		await page.waitForTimeout(1500);
		await page.screenshot({ path: out('demo-onedrive-saved.png') });
		await context.close();
		console.log('demo screenshot written; File > Save uploaded to the same item');
	}
} finally {
	if (errors.length) console.log('page errors:', errors);
	await browser.close();
}
