import { expect, test, type Page } from '@playwright/test';
import { openSample } from './helpers';

type Paragraph = { type: string; id: string; runs: { text: string }[] };
type WordEditor = HTMLElement & { documentModel: { blocks: Paragraph[] } };
type Probe = Window & { removedEditor?: WeakRef<HTMLElement> };

const editor = (page: Page) => page.locator('docx-editor');

/** Replaces the sample with enough text for virtualized Print Layout, then shows that view. */
async function showLongPrintLayout(page: Page) {
	await editor(page).evaluate((element) => {
		const word = element as WordEditor;
		const text = 'Enough words to fill a line of the page and keep the sheets coming. '.repeat(3);
		word.documentModel = {
			...word.documentModel,
			blocks: Array.from({ length: 900 }, (_, index) => ({
				type: 'paragraph',
				id: `long-${index}`,
				runs: [{ text }],
			})),
		};
		const view = word.shadowRoot!.querySelector<HTMLSelectElement>('[aria-label="Layout view"]')!;
		view.value = 'print';
		view.dispatchEvent(new Event('change'));
	});
	await expect.poll(() => editor(page).locator('.dve-print-page').count()).toBeGreaterThan(10);
}

/** Forces collection until the probe's editor is gone. */
async function expectCollected(page: Page) {
	const session = await page.context().newCDPSession(page);
	await expect
		.poll(async () => {
			await session.send('HeapProfiler.collectGarbage');
			return page.evaluate(() => (window as Probe).removedEditor!.deref() === undefined);
		})
		.toBe(true);
}

const reconnect = (page: Page) =>
	editor(page).evaluate(async (element) => {
		const parent = element.parentElement!;
		element.remove();
		await new Promise((resolve) => setTimeout(resolve, 50));
		parent.append(element);
		await new Promise((resolve) => setTimeout(resolve, 200));
	});

test('a removed editor is collected', async ({ page }) => {
	await openSample(page);
	await page.evaluate(async () => {
		const element = document.createElement('docx-editor');
		document.body.append(element);
		await new Promise((resolve) => setTimeout(resolve, 200));
		(window as Probe).removedEditor = new WeakRef(element);
		element.remove();
	});
	await expectCollected(page);
});

test('an editor removed with its context menu open is collected', async ({ page }) => {
	await openSample(page);
	await page.evaluate(async () => {
		const element = document.createElement('docx-editor');
		document.body.append(element);
		await new Promise((resolve) => setTimeout(resolve, 200));
		element
			.shadowRoot!.querySelector('.ProseMirror')!
			.dispatchEvent(
				new MouseEvent('contextmenu', { bubbles: true, composed: true, cancelable: true }),
			);
		await new Promise((resolve) => setTimeout(resolve, 100));
		if (!element.shadowRoot!.querySelector('office-ui-context-menu'))
			throw new Error('The context menu did not open');
		(window as Probe).removedEditor = new WeakRef(element);
		element.remove();
	});
	await expectCollected(page);
});

test('a reconnected editor still prints every page', async ({ page }) => {
	await openSample(page);
	await showLongPrintLayout(page);
	await reconnect(page);
	const sheets = () =>
		editor(page)
			.locator('.dve-print-page')
			.evaluateAll((pages) => ({
				total: pages.length,
				filled: pages.filter((sheet) => sheet.childElementCount).length,
			}));
	await expect.poll(async () => (await sheets()).total).toBeGreaterThan(10);
	const before = await sheets();
	expect(before.filled).toBeLessThan(before.total);
	await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
	const printing = await sheets();
	expect(printing.filled).toBe(printing.total);
	await page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
	expect((await sheets()).filled).toBeLessThan(printing.total);
});

test('a reconnected editor lays out again when a font arrives', async ({ page }) => {
	await openSample(page);
	await showLongPrintLayout(page);
	await reconnect(page);
	await expect.poll(() => editor(page).locator('.dve-print-page').count()).toBeGreaterThan(10);
	const replaced = await editor(page).evaluate((element) => {
		const pages = element.shadowRoot!.querySelector('.dve-print-pages')!;
		const before = pages.firstElementChild;
		document.fonts.dispatchEvent(new Event('loadingdone'));
		return pages.firstElementChild !== before;
	});
	expect(replaced).toBe(true);
});
