/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * Rich clipboard interop with other applications, framework-neutral:
 *
 * - formatted HTML on the system clipboard pastes into a text box being edited as runs (bold
 *   kept, Keep Source Formatting), the editor reopens after it, and plain text then pastes as text;
 * - copying a shape writes a picture (`image/png`) and its text (`text/plain`) to the system
 *   clipboard while the in-app paste still clones the shape.
 *
 * Run: bunx playwright test rich-clipboard
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import type { PowerPointViewerAPI } from 'pptx-viewer-shared';

import { elementWithText, loadDeck, SAMPLE_DECK, selectElement } from './support/deck';
import { pressShortcut } from './support/keyboard';

declare global {
	interface Window {
		/** Public component handle exposed by each development demo. */
		__pptxViewer: PowerPointViewerAPI;
	}
}

const SHAPE_TEXT = 'Product Overview';

async function writeHtmlClipboard(page: Page, html: string, text: string): Promise<void> {
	await page.evaluate(
		async ([markup, plain]) => {
			await navigator.clipboard.write([
				new ClipboardItem({
					'text/html': new Blob([markup], { type: 'text/html' }),
					'text/plain': new Blob([plain], { type: 'text/plain' }),
				}),
			]);
		},
		[html, text] as const,
	);
}

/** The runs of the first element whose text contains `needle`. */
async function runsOf(page: Page, needle: string) {
	return page.evaluate((text) => {
		const element = window.__pptxViewer
			.getElements(window.__pptxViewer.getActiveSlideIndex())
			.find((candidate) => (candidate as { text?: string }).text?.includes(text)) as
			| { textSegments?: { text: string; style?: { bold?: boolean } }[] }
			| undefined;
		return (element?.textSegments ?? []).map((segment) => ({
			text: segment.text,
			bold: segment.style?.bold === true,
		}));
	}, needle);
}

test.beforeEach(async ({ context }) => {
	await context.grantPermissions(['clipboard-read', 'clipboard-write']);
});

test('formatted HTML pastes into a text box as runs and plain text after it stays text', async ({
	page,
}) => {
	await loadDeck(page, SAMPLE_DECK);
	await elementWithText(page, SHAPE_TEXT).dblclick();
	const editor = page.locator('[data-inline-editor]').first();
	await expect(editor).toBeVisible();
	await editor.press('ControlOrMeta+End');
	await writeHtmlClipboard(page, '<p><b>Rich</b> tail</p>', 'Rich tail');
	await editor.press('ControlOrMeta+v');

	// The paste is written to the model as runs; the editor reopens with the caret after it.
	await expect
		.poll(async () =>
			(await runsOf(page, 'Rich tail')).some((run) => run.bold && run.text === 'Rich'),
		)
		.toBe(true);
	const reopened = page.locator('[data-inline-editor]').first();
	await expect(reopened).toBeVisible();
	await expect(reopened).toBeFocused();

	// Plain text from another app lands at the caret after the pasted runs, as text, not markup.
	await page.evaluate(() => navigator.clipboard.writeText('!'));
	await reopened.press('ControlOrMeta+v');
	await reopened.press('Escape');
	await expect
		.poll(async () => {
			const runs = await runsOf(page, 'Rich');
			return {
				text: runs.map((run) => run.text).join(''),
				boldRich: runs.filter((run) => run.bold && run.text.includes('Rich')).length,
			};
		})
		.toEqual({ text: expect.stringContaining(`${SHAPE_TEXT}Rich tail!`), boldRich: 1 });
});

test('a rich paste mid-way into a bulleted list keeps the bullet and puts the caret after it', async ({
	page,
}) => {
	await loadDeck(page, SAMPLE_DECK);
	await page.evaluate((needle) => {
		const api = window.__pptxViewer;
		const target = api
			.getElements(api.getActiveSlideIndex())
			.find((candidate) => (candidate as { text?: string }).text?.includes(needle))!;
		const style = { fontSize: 24 };
		api.updateElement(target.id, {
			text: 'One\nTwo',
			textSegments: [
				{ text: 'One', style, bulletInfo: { char: '\u2022' } },
				{ text: '\n', style, isParagraphBreak: true },
				{ text: 'Two', style, bulletInfo: { char: '\u2022' } },
			],
		} as never);
	}, SHAPE_TEXT);
	await elementWithText(page, 'Two').dblclick();
	const editor = page.locator('[data-inline-editor]').first();
	await expect(editor).toBeVisible();
	await editor.press('ControlOrMeta+End');
	await editor.press('Home');
	await writeHtmlClipboard(page, '<b>New</b>', 'New');
	await editor.press('ControlOrMeta+v');
	await expect
		.poll(async () => (await runsOf(page, 'NewTwo')).some((run) => run.bold && run.text === 'New'))
		.toBe(true);
	const reopened = page.locator('[data-inline-editor]').first();
	await expect(reopened).toBeFocused();
	await page.keyboard.type('Z');
	await reopened.press('Escape');
	await expect
		.poll(() =>
			page.evaluate(() => {
				const api = window.__pptxViewer;
				const element = api
					.getElements(api.getActiveSlideIndex())
					.find((candidate) => (candidate as { text?: string }).text?.includes('Two')) as
					| { text?: string; textSegments?: { text: string; bulletInfo?: unknown }[] }
					| undefined;
				return {
					text: element?.text,
					bullets: (element?.textSegments ?? []).filter((segment) => segment.bulletInfo).length,
				};
			}),
		)
		.toEqual({ text: 'One\nNewZTwo', bullets: 2 });
});

test('copying a shape offers a picture and its text to other applications', async ({ page }) => {
	await loadDeck(page, SAMPLE_DECK);
	await selectElement(page, elementWithText(page, SHAPE_TEXT));
	await pressShortcut(page, 'ControlOrMeta+c');
	await expect
		.poll(
			() =>
				page.evaluate(async () => {
					const items = await navigator.clipboard.read();
					return items.flatMap((item) => [...item.types]).sort();
				}),
			{ timeout: 15000 },
		)
		.toEqual(['image/png', 'text/plain']);
	expect(await page.evaluate(() => navigator.clipboard.readText())).toContain(SHAPE_TEXT);
	const before = await page.evaluate(
		() => window.__pptxViewer.getElements(window.__pptxViewer.getActiveSlideIndex()).length,
	);
	// The in-app paste still clones the shape rather than pasting the picture.
	await pressShortcut(page, 'ControlOrMeta+v');
	await expect
		.poll(() =>
			page.evaluate(
				() => window.__pptxViewer.getElements(window.__pptxViewer.getActiveSlideIndex()).length,
			),
		)
		.toBe(before + 1);
});
