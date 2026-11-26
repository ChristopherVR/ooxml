/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { expect, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';
import {
	addInName,
	jsonManifest,
	packageFiles,
	paneOrigin,
	publisherPage,
	publisherScript,
	xmlManifest,
} from '../add-in-fixtures/manifest';

export async function enforcePublisherFramePolicy(page: Page) {
	await page.route('**/*', async (route) => {
		if (!route.request().isNavigationRequest() || route.request().frame() !== page.mainFrame())
			return route.fallback();
		const response = await route.fetch();
		return route.fulfill({
			response,
			headers: {
				...response.headers(),
				'content-security-policy': `frame-src ${paneOrigin}`,
			},
		});
	});
}

export const addInReviewScenarios = [1280, 800].flatMap((width) =>
	(['light', 'dark'] as const).map((scheme) => ({
		width,
		height: width === 1280 ? 800 : 700,
		scheme,
	})),
);

export async function prepareAddInDemo(
	page: Page,
	{ width, height, scheme }: (typeof addInReviewScenarios)[number],
) {
	await enforcePublisherFramePolicy(page);
	await page.emulateMedia({ colorScheme: scheme });
	await page.setViewportSize({ width, height });
	await page.addInitScript((color) => {
		localStorage.setItem('vitepress-theme-appearance', color);
		localStorage.setItem('pptx-demo-theme', color === 'dark' ? 'vermilionDark' : 'vermilionLight');
	}, scheme);
}

interface Flow {
	page: Page;
	editor: Locator;
	openDialog: () => Promise<void>;
	reloadEditor: () => Promise<void>;
	host: 'Word' | 'Excel' | 'PowerPoint';
}

export async function checkAddInRuntime({ page, editor, openDialog, reloadEditor, host }: Flow) {
	const parentOrigin = new URL(page.url()).origin;
	await page.route(`${paneOrigin}/**`, (route) => {
		const path = new URL(route.request().url()).pathname;
		if (path === '/client.js')
			return route.fulfill({ contentType: 'application/javascript', body: publisherScript() });
		if (path === '/pane')
			return route.fulfill({
				contentType: 'text/html',
				headers: { 'content-security-policy': "default-src 'none'; script-src 'self'" },
				body: publisherPage(parentOrigin),
			});
		return route.abort();
	});
	const enableRuntime = () =>
		editor.evaluate(
			(node, { origin, selection }) => {
				Reflect.set(node, Symbol.for('office-ui.add-in-runtime'), {
					origins: [origin],
					frameSrcEnforced: true,
					...(selection && {
						selection: () => ({ readText: () => 'Slide selection', writeText: () => false }),
					}),
				});
				node.dispatchEvent(new Event('office-add-in-runtime-changed'));
			},
			{ origin: paneOrigin, selection: host === 'PowerPoint' },
		);
	const zip = new JSZip();
	for (const [name, contents] of packageFiles) zip.file(name, contents);
	const uploads = [
		{ name: 'manifest.xml', mimeType: 'application/xml', buffer: Buffer.from(xmlManifest) },
		{ name: 'manifest.json', mimeType: 'application/json', buffer: Buffer.from(jsonManifest) },
		{
			name: 'app.zip',
			mimeType: 'application/zip',
			buffer: await zip.generateAsync({ type: 'nodebuffer' }),
		},
	];
	const dialog = page.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	const pane = page.getByRole('complementary', { name: addInName, exact: true });
	const frame = pane.frameLocator('iframe');

	for (const file of uploads) {
		await openDialog();
		await expect(dialog.getByRole('tab', { name: 'Store', exact: true })).toHaveCount(0);
		await expect(dialog.getByLabel('Add from Marketplace link')).toHaveCount(0);
		await dialog.getByLabel('Choose a local add-in file').setInputFiles(file);
		const consent = dialog.getByRole('region', { name: 'Review add-in permissions' });
		await expect(consent).toContainText('Lab Notes');
		await expect(consent).toContainText('Word, Excel, PowerPoint');
		await expect(consent).toContainText('Read and change this document');
		if (file.name === 'manifest.xml') {
			await expect(consent).toContainText('Running add-ins is unavailable');
			if (host !== 'PowerPoint') await expect(consent).toContainText('Function commands');
		}
		if (file.name === 'app.zip')
			await expect(consent.getByRole('img', { name: 'Icon preview' }).first()).toBeVisible();
		else await expect(consent.locator('img')).toHaveCount(0);
		const install = dialog.getByRole('button', { name: 'Install Notes', exact: true });
		await expect(install).toHaveAttribute('slot', 'footer');
		await install.click();
		await expect(dialog).toBeHidden();
		await expect(editor.locator('[data-office-add-in-status]')).toHaveText('Notes installed.');
		if (file.name !== 'manifest.xml')
			await expect(editor.getByRole('tab', { name: addInName, exact: true })).toBeFocused();

		if (file.name === 'manifest.json') {
			const tab = editor.getByRole('tab', { name: addInName, exact: true });
			if (host !== 'PowerPoint' && (page.viewportSize()?.width ?? 0) === 800) {
				await editor.getByRole('tab', { name: 'Home', exact: true }).click();
				const more = editor.getByRole('button', { name: 'More tabs', exact: true });
				await expect(more).toBeVisible();
				await more.focus();
				await more.press('Enter');
				await expect(tab).toBeInViewport();
			}
			await page.reload();
			await reloadEditor();
			await expect(tab).toBeAttached();
			await tab.click();
			const command = editor.getByRole('button', { name: 'Open Notes', exact: true });
			await expect(command).toBeDisabled();
			await enableRuntime();
			await expect(command).toBeEnabled();
			if (host === 'PowerPoint')
				await expect(editor.locator('[data-pptx-chrome="ribbon-content"]:visible')).toHaveCount(0);
			await command.click();
		} else {
			await openDialog();
			if (file.name === 'app.zip')
				await expect(dialog.locator('.tile-icon[src^="data:image/png;base64,"]')).toBeVisible();
			await dialog.getByText('Add-in details', { exact: true }).click();
			await expect(
				dialog.getByRole('button', { name: 'Open Notes', exact: true }),
			).toBeInViewport();
			await expect(
				dialog.getByRole('button', { name: 'Remove Notes', exact: true }),
			).toBeInViewport();
			await expect(dialog.getByText('Lab Notes', { exact: true })).toBeVisible();
			const open = dialog.getByRole('button', { name: 'Open Notes', exact: true });
			if (file.name === 'manifest.xml') {
				await expect(open).toBeDisabled();
				await expect(dialog).toContainText('Running add-ins is unavailable');
				await enableRuntime();
			}
			await open.click();
		}
		await expect(pane.getByRole('heading', { name: addInName })).toBeFocused();
		const separator = editor.getByRole('separator', { name: 'Resize add-in pane' });
		await expect(separator).toHaveAttribute('aria-valuenow', '320');
		await separator.focus();
		await separator.press('ArrowLeft');
		await expect(separator).toHaveAttribute('aria-valuenow', '336');
		const bounds = await pane.boundingBox();
		const viewport = editor
			.locator(
				host === 'Word'
					? '.dve-canvas'
					: host === 'Excel'
						? '.xve-workspace'
						: '[data-pptx-viewport]',
			)
			.first();
		const document = await viewport.boundingBox();
		expect(bounds?.height).toBeGreaterThan(300);
		expect((document?.x ?? 0) + (document?.width ?? 0)).toBeLessThanOrEqual((bounds?.x ?? 0) + 1);
		expect(bounds?.x).toBeGreaterThan(320);
		expect((bounds?.x ?? 0) + (bounds?.width ?? 0)).toBeLessThanOrEqual(
			page.viewportSize()?.width ?? 0,
		);
		await expect(pane.locator('iframe')).toHaveAttribute('sandbox', 'allow-scripts');
		await expect(pane.locator('iframe')).toHaveAttribute('referrerpolicy', 'no-referrer');
		await expect(frame.locator('#ready')).toHaveText(
			`${host} / ${host} / true / true / false / false`,
		);
		await expect(frame.locator('#isolated')).toHaveText('Editor DOM is isolated');
		if (host !== 'PowerPoint') {
			await frame.getByRole('button', { name: 'Write selection' }).click();
			await expect(frame.locator('#result')).toHaveText('Wrote text');
			await frame.getByRole('button', { name: 'Read selection' }).click();
			await expect(frame.locator('#result')).toHaveText(
				host === 'Excel' ? 'Read: Updated by add-in' : 'Read: (empty)',
			);
		} else {
			await frame.getByRole('button', { name: 'Read selection' }).click();
			await expect(frame.locator('#result')).toHaveText('Read: Slide selection');
		}
		await pane.getByRole('button', { name: 'Close', exact: true }).click();
		await openDialog();
		await dialog.getByRole('button', { name: 'Remove Notes', exact: true }).click();
		await expect(dialog).toBeHidden();
		await expect(editor.getByRole('tab', { name: addInName, exact: true })).toHaveCount(0);
	}
}
