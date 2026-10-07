import { readFile } from 'node:fs/promises';
import { expect, test, type Page } from '@playwright/test';
import { fixture, shape, cell, rectangle } from '../../src/core/visio/test-fixtures';

async function open(page: Page, name: string, kind: 'docx' | 'xlsx' | 'vsdx' | 'other') {
	await page.locator('teams-app').evaluate(async (element) => {
		await (element as HTMLElement & { updateComplete: Promise<boolean> }).updateComplete;
	});
	await page.locator('teams-app').evaluate(
		(element, detail) => {
			(element as HTMLElement & { previewContent(detail: unknown): void }).previewContent(detail);
		},
		{ attachment: { name, kind }, url: `${new URL(page.url()).origin}/content/${name}` },
	);
}

test('Markdown and sites preview inside the workspace and close back to chat', async ({ page }) => {
	await page.route('**/content/notes.md', (route) =>
		route.fulfill({
			contentType: 'text/markdown',
			body: '# Project notes\n<script>window.pwned=true</script>\n```\n<b>literal</b>\n```',
		}),
	);
	await page.route('**/content/index.html', (route) =>
		route.fulfill({ contentType: 'text/html', body: '<h1>Static project site</h1>' }),
	);
	await page.goto('/?local=1&name=Ada&room=content-browser');
	await open(page, 'notes.md', 'other');
	await expect(page.getByRole('heading', { name: 'Project notes' })).toBeVisible();
	await expect(page.locator('teams-content-preview code')).toHaveText('<b>literal</b>');
	expect(await page.evaluate(() => 'pwned' in window)).toBe(false);
	await page.screenshot({ path: test.info().outputPath('markdown-preview.png') });
	await page.getByRole('button', { name: 'Close preview' }).click();
	await expect(page.locator('office-ui-chat-composer')).toBeVisible();
	await open(page, 'index.html', 'other');
	await expect(
		page
			.frameLocator('teams-content-preview iframe')
			.getByRole('heading', { name: 'Static project site' }),
	).toBeVisible();
	await expect(page.locator('teams-content-preview iframe')).toHaveAttribute(
		'sandbox',
		'allow-scripts allow-forms',
	);
});

test('native Office viewers load actual bytes in the workspace', async ({ page }) => {
	test.setTimeout(120_000);
	const word = await readFile(
		new URL('../../src/core/docx/__fixtures__/smartart.docx', import.meta.url),
	);
	const excel = await readFile(
		new URL('../../src/core/xlsx/__fixtures__/openpyxl-features.xlsx', import.meta.url),
	);
	const visio = await fixture({
		pages: [
			{
				id: '1',
				contents:
					'<Shapes>' +
					shape(
						'1',
						rectangle +
							cell('PinX', 2) +
							cell('PinY', 2) +
							cell('Width', 2) +
							cell('Height', 1) +
							'<Text>Teams drawing</Text>',
					) +
					'</Shapes>',
			},
		],
	});
	for (const [name, bytes] of [
		['plan.docx', word],
		['budget.xlsx', excel],
		['drawing.vsdx', visio],
	] as const)
		await page.route(`**/content/${name}`, (route) =>
			route.fulfill({ body: Buffer.from(bytes), contentType: 'application/octet-stream' }),
		);
	await page.goto('/?local=1&name=Ada&room=office-preview');
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible();
	for (const [name, kind, tag] of [
		['plan.docx', 'docx', 'docx-editor'],
		['budget.xlsx', 'xlsx', 'xlsx-editor'],
		['drawing.vsdx', 'vsdx', 'visio-viewer'],
	] as const) {
		await open(page, name, kind);
		await expect(page.locator(`teams-content-preview ${tag}`)).toBeVisible({ timeout: 60_000 });
		await expect
			.poll(
				() =>
					page
						.locator('teams-content-preview')
						.evaluate((el) => (el as HTMLElement & { status: string }).status),
				{ timeout: 60_000 },
			)
			.toBe('ready');
		if (kind !== 'vsdx')
			expect(
				await page
					.locator(tag)
					.evaluate((el) => (el as HTMLElement & { readOnly: boolean }).readOnly),
			).toBe(true);
		else
			await expect(
				page.locator('visio-viewer').getByRole('region', { name: 'Diagram canvas' }).locator('svg'),
			).toBeVisible();
	}
});
