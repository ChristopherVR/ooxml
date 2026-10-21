import { expect, test, type Page } from '@playwright/test';
import type { NativeHarness } from '../../demos/docx/demo-vanilla/building-blocks-types';

const FRAMEWORKS = ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'] as const;
const TEXT = 'Native Word composition';

async function open(page: Page, framework: string) {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto(`/building-blocks.html?framework=${framework}`);
	await expect(page.locator('.ProseMirror')).toContainText(TEXT);
	await expect(page.getByRole('button', { name: 'Bold', exact: true })).toBeVisible();
	return errors;
}
const selection = (page: Page) =>
	page.evaluate(() => {
		const host: NativeHarness = window.nativeWord!;
		return host.toolbar()?.selection ?? null;
	});

for (const framework of FRAMEWORKS) {
	test(`${framework}: a read-only surface can be focused, selected, navigated and copied but not edited`, async ({
		page,
	}) => {
		const errors = await open(page, framework);
		await page.evaluate(() => window.nativeWord!.permission(false));
		await expect.poll(() => page.evaluate(() => window.nativeWord!.mode())).toBe('view');
		const text = page.locator('.ProseMirror p');
		// A double click selects a word, and the toolbar props follow it.
		await text.dblclick({ position: { x: 20, y: 8 } });
		await expect
			.poll(async () => {
				const range = await selection(page);
				return range ? range.to - range.from : 0;
			})
			.toBeGreaterThan(2);
		// Keyboard navigation: Home, then Shift+Right three times extends by three characters.
		await page.keyboard.press('Home');
		await page.keyboard.press('Shift+ArrowRight');
		await page.keyboard.press('Shift+ArrowRight');
		await page.keyboard.press('Shift+ArrowRight');
		await expect
			.poll(async () => {
				const range = await selection(page);
				return range ? range.to - range.from : 0;
			})
			.toBe(3);
		await page.keyboard.press('ControlOrMeta+a');
		await expect
			.poll(async () => {
				const range = await selection(page);
				return range ? range.to - range.from : 0;
			})
			.toBe(TEXT.length);
		// Nothing types, and no editing command runs.
		await page.keyboard.type('zzz');
		await page.keyboard.press('Delete');
		await page.keyboard.press('ControlOrMeta+b');
		await expect(page.locator('.ProseMirror')).toContainText(TEXT);
		await expect(page.locator('.ProseMirror')).not.toContainText('zzz');
		expect(await page.evaluate(() => window.nativeWord!.handle.dirty)).toBe(false);
		expect(await page.evaluate(() => window.nativeWord!.toolbar()!.onBold())).toBe(false);
		// The selection survives for copy.
		expect(
			await page.evaluate(() => {
				const root = document.querySelector('docx-editor')!.shadowRoot as ShadowRoot & {
					getSelection?: () => Selection | null;
				};
				return (root.getSelection?.() ?? document.getSelection())?.toString();
			}),
		).toContain('Native');
		expect(errors).toEqual([]);
	});

	test(`${framework}: navigation, status and the full toolbar are driven by their props`, async ({
		page,
	}) => {
		const errors = await open(page, framework);
		// Library status controls: word count and zoom.
		await expect(page.locator('[role="status"]', { hasText: '3 words' })).toBeVisible();
		await page.getByLabel('Zoom', { exact: true }).selectOption('150');
		await expect.poll(() => page.evaluate(() => window.nativeWord!.status()!.zoom)).toBe(150);
		await page.evaluate(() => window.nativeWord!.status()!.onZoomOut());
		await expect.poll(() => page.evaluate(() => window.nativeWord!.status()!.zoom)).toBe(140);
		// Pages exist in Print Layout only.
		expect(await page.evaluate(() => window.nativeWord!.navigation()!.pagesAvailable)).toBe(false);
		await expect(page.getByRole('button', { name: 'Previous page' })).toBeDisabled();
		await page.evaluate(() => {
			const model = window.nativeWord!.handle.element.documentModel!;
			window.nativeWord!.options({
				documentModel: {
					...model,
					blocks: Array.from({ length: 140 }, (_, index) => ({
						type: 'paragraph' as const,
						id: `p${index}`,
						runs: [{ text: `Paragraph number ${index} fills the page with words to wrap.` }],
					})),
				},
			});
		});
		await page.getByLabel('Layout view').selectOption('print');
		await expect
			.poll(() => page.evaluate(() => window.nativeWord!.navigation()!.pageCount))
			.toBeGreaterThan(1);
		await expect(page.getByRole('button', { name: 'Next page' })).toBeEnabled();
		await page.getByRole('button', { name: 'Next page' }).click();
		await expect.poll(() => page.evaluate(() => window.nativeWord!.navigation()!.page)).toBe(2);
		await expect(page.locator('[role="status"]', { hasText: /^Page 2 of \d+$/ })).toBeVisible();
		await page.getByRole('button', { name: 'Previous page' }).click();
		await expect.poll(() => page.evaluate(() => window.nativeWord!.navigation()!.page)).toBe(1);
		expect(await page.evaluate(() => window.nativeWord!.navigation()!.onGoToPage(2))).toBe(true);
		await expect.poll(() => page.evaluate(() => window.nativeWord!.navigation()!.page)).toBe(2);
		// Toolbar: alignment, font size, lists and formatting reflect state with pressed buttons.
		await page.getByLabel('Layout view').selectOption('draft');
		await page.evaluate(() => {
			const model = window.nativeWord!.handle.element.documentModel!;
			window.nativeWord!.options({
				documentModel: {
					...model,
					blocks: [{ type: 'paragraph', id: 'only', runs: [{ text: 'Native Word composition' }] }],
				},
			});
		});
		await expect(page.locator('.ProseMirror')).toContainText(TEXT);
		await page.locator('.ProseMirror p').dblclick({ position: { x: 20, y: 8 } });
		await page.getByRole('button', { name: 'Align center' }).click();
		await expect(page.getByRole('button', { name: 'Align center' })).toHaveAttribute(
			'aria-pressed',
			'true',
		);
		await expect(page.getByRole('button', { name: 'Align left' })).toHaveAttribute(
			'aria-pressed',
			'false',
		);
		await page.getByLabel('Font size', { exact: true }).selectOption('20');
		await expect.poll(() => page.evaluate(() => window.nativeWord!.toolbar()!.fontSize)).toBe(20);
		await page.getByRole('button', { name: 'Bulleted list' }).click();
		await expect(page.getByRole('button', { name: 'Bulleted list' })).toHaveAttribute(
			'aria-pressed',
			'true',
		);
		await page.getByRole('button', { name: 'Bulleted list' }).click();
		await expect(page.getByRole('button', { name: 'Bulleted list' })).toHaveAttribute(
			'aria-pressed',
			'false',
		);
		expect(errors).toEqual([]);
	});

	test(`${framework}: host workflows are enabled only by their callbacks and open no library dialog`, async ({
		page,
	}) => {
		const errors = await open(page, framework);
		await expect(page.getByRole('button', { name: 'Find', exact: true })).toHaveCount(0);
		expect(
			await page.evaluate(() => {
				const props = window.nativeWord!.toolbar()!;
				return [
					props.canInsertLink,
					props.canFindReplace,
					props.canInsertPicture,
					props.onOpenHyperlinkDialog(),
					props.onOpenFindReplace(),
					props.onOpenImagePicker(),
				];
			}),
		).toEqual([false, false, false, false, false, false]);
		await page.evaluate(() => {
			const host = window.nativeWord!;
			const store = window as unknown as { requests: Record<string, unknown> };
			store.requests = {};
			host.options({
				documentModel: host.handle.element.documentModel!,
				onOpenHyperlinkDialog: (request) => (store.requests.link = request),
				onOpenFindReplace: (request) => (store.requests.find = request),
				onOpenImagePicker: (request) => (store.requests.picture = request),
			});
		});
		await page.locator('.ProseMirror p').dblclick({ position: { x: 20, y: 8 } });
		await expect(page.getByRole('button', { name: 'Insert link' })).toBeEnabled();
		await page.getByRole('button', { name: 'Insert link' }).click();
		await page.getByRole('button', { name: 'Find', exact: true }).click();
		await page.getByRole('button', { name: 'Insert picture' }).click();
		const result = await page.evaluate(() => {
			const requests = (
				window as unknown as {
					requests: {
						link: { href: string | null; text: string; apply(href: string): boolean };
						find: { search(query: string): { count: number } };
						picture: { insert(file: File): Promise<boolean> };
					};
				}
			).requests;
			const applied = requests.link.apply('https://example.com/');
			return {
				text: requests.link.text,
				href: requests.link.href,
				applied,
				found: requests.find.search('Word').count,
				picture: typeof requests.picture.insert,
			};
		});
		expect(result).toMatchObject({ href: null, applied: true, found: 1, picture: 'function' });
		await expect(page.locator('.ProseMirror a[href="https://example.com/"]')).toHaveCount(1);
		// No library dialog, search panel or file input is ever mounted in surface mode.
		expect(
			await page.evaluate(
				() =>
					document
						.querySelector('docx-editor')!
						.shadowRoot!.querySelectorAll(
							'[role="dialog"], .dve-search-panel, input[type="file"], office-ui-dialog',
						).length,
			),
		).toBe(0);
		// Read-only closes the editing workflows but keeps find.
		await page.evaluate(() => window.nativeWord!.permission(false));
		await expect(page.getByRole('button', { name: 'Insert link' })).toBeDisabled();
		await expect(page.getByRole('button', { name: 'Find', exact: true })).toBeEnabled();
		expect(errors).toEqual([]);
	});
}
