/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * The shared-contract input `collaboration: { roomId, user: { name } }` (no server, transport,
 * provider or session) opens a BroadcastChannel room between windows of one browser. Two pages
 * of ONE context join the same room through `?sameBrowser=<room>` and must converge.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

const elements = (page: Page) => page.locator('[data-pptx-viewport] [data-pptx-element="true"]');
const state = (page: Page) => page.getByLabel('Same-browser collaboration', { exact: true });

async function open(page: Page, room: string, name: string, join: boolean): Promise<void> {
	const params = new URLSearchParams({ sameBrowser: room, name });
	if (join) {
		params.set('join', '1');
	}
	await page.goto(`/?${params}`);
	await expect(elements(page).first()).toBeVisible({ timeout: 30_000 });
	await expect(state(page)).toContainText('synced: true', { timeout: 30_000 });
}

async function replaceText(page: Page, original: string, replacement: string): Promise<void> {
	await elements(page).filter({ hasText: original }).dblclick();
	const editor = page.locator('[data-inline-editor]').first();
	await expect(editor).toBeVisible();
	await editor.press('ControlOrMeta+A');
	await page.keyboard.type(replacement);
	await state(page).click();
}

test.use({ viewport: { width: 1440, height: 1000 } });

test.describe('same-browser collaboration', () => {
	test.setTimeout(120_000);

	test('undo and redo preserve another participant’s later edit', async ({
		browser,
		baseURL,
	}, info) => {
		const context = await browser.newContext({ baseURL, viewport: { width: 1440, height: 1000 } });
		const a = await context.newPage();
		const b = await context.newPage();
		const room = `undo-${info.project.name}-${Date.now()}`;
		try {
			await open(a, room, 'Alice', false);
			await open(b, room, 'Bob', true);
			await replaceText(a, 'Product Overview', 'Alice edit');
			await expect(elements(b).filter({ hasText: 'Alice edit' })).toBeVisible();
			await replaceText(b, 'Q2 2026', 'Bob edit');
			await expect(elements(a).filter({ hasText: 'Bob edit' })).toBeVisible();
			const undo = a.getByRole('button', { name: /^undo/iu }).first();
			await expect(undo).toBeEnabled();
			await undo.click();
			await expect(elements(a).filter({ hasText: 'Product Overview' })).toBeVisible();
			await expect(elements(a).filter({ hasText: 'Bob edit' })).toBeVisible();
			await expect(elements(b).filter({ hasText: 'Bob edit' })).toBeVisible();
			const redo = a.getByRole('button', { name: /^redo/iu }).first();
			await expect(redo).toBeEnabled();
			await redo.click();
			await expect(elements(a).filter({ hasText: 'Alice edit' })).toBeVisible();
			await expect(elements(b).filter({ hasText: 'Alice edit' })).toBeVisible();
			await expect(elements(b).filter({ hasText: 'Bob edit' })).toBeVisible();
		} finally {
			await context.close();
		}
	});

	test('Escape keeps live text in both peers through a subsequent edit', async ({
		browser,
		baseURL,
	}, info) => {
		const context = await browser.newContext({ baseURL, viewport: { width: 1440, height: 1000 } });
		const a = await context.newPage();
		const b = await context.newPage();
		const room = `escape-${info.project.name}-${Date.now()}`;
		try {
			await open(a, room, 'Alice', false);
			await open(b, room, 'Bob', true);
			await elements(a).filter({ hasText: 'Product Overview' }).dblclick();
			const editor = a.locator('[data-inline-editor]').first();
			await expect(editor).toBeVisible();
			await editor.press('ControlOrMeta+A');
			await a.keyboard.type('Escape first');
			await expect(elements(b).filter({ hasText: 'Escape first' })).toBeVisible();
			await a.keyboard.press('Escape');
			await expect(editor).toBeHidden();
			await expect(elements(a).filter({ hasText: 'Escape first' })).toBeVisible();
			await elements(a).filter({ hasText: 'Escape first' }).dblclick();
			await expect(editor).toBeVisible();
			await editor.press('End');
			await a.keyboard.type(' second');
			await a.keyboard.press('Escape');
			await expect(elements(a).filter({ hasText: 'Escape first second' })).toBeVisible();
			await expect(elements(b).filter({ hasText: 'Escape first second' })).toBeVisible();
		} finally {
			await context.close();
		}
	});

	test('two pages of one context share a room, an edit and their collaborators', async ({
		browser,
		baseURL,
	}, info) => {
		const context = await browser.newContext({ baseURL, viewport: { width: 1440, height: 1000 } });
		const a = await context.newPage();
		const b = await context.newPage();
		const room = `same-browser-${info.project.name}-${Date.now()}`;
		try {
			await open(a, room, 'Alice', false);
			// A edits before B joins: B can only show the edit by receiving the room's state.
			await replaceText(a, 'Product Overview', 'Edited before B joined');
			await open(b, room, 'Bob', true);
			await expect(elements(b).filter({ hasText: 'Edited before B joined' })).toBeVisible({
				timeout: 30_000,
			});
			await expect(elements(b)).toHaveCount(await elements(a).count());

			// A live edit in A reaches B.
			await replaceText(a, 'Edited before B joined', 'Live edit from A');
			await expect(elements(b).filter({ hasText: 'Live edit from A' })).toBeVisible({
				timeout: 30_000,
			});

			// And the other way round.
			await replaceText(b, 'Live edit from A', 'Reply from B');
			await expect(elements(a).filter({ hasText: 'Reply from B' })).toBeVisible({
				timeout: 30_000,
			});

			// Both report two people, each named from the other's side.
			for (const [page, other] of [
				[a, 'Bob'],
				[b, 'Alice'],
			] as const) {
				await expect(state(page)).toContainText('people: 2', { timeout: 30_000 });
				await expect(state(page)).toContainText(other);
			}
			await a.screenshot({ path: info.outputPath('same-browser-a.png') });
			await b.screenshot({ path: info.outputPath('same-browser-b.png') });
		} finally {
			await context.close();
		}
	});
});
