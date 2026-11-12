import { expect, test, type Page } from '@playwright/test';
import type { NativeHarness } from '../../demos/docx/demo-vanilla/building-blocks-types';

const FRAMEWORKS = ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'] as const;
const TEXT = 'Native Word composition';

declare global {
	interface Window {
		__speech?: {
			spoken: string[];
			cancels: number;
			fire: (type: 'end' | 'error', error?: string) => void;
		};
	}
}

/** Replaces the browser speech synthesis with a recorder the test can drive, or removes it. */
async function stubSpeech(page: Page, supported: boolean) {
	await page.addInitScript((enabled) => {
		if (!enabled) {
			Object.defineProperty(window, 'speechSynthesis', { value: undefined, configurable: true });
			Object.defineProperty(window, 'SpeechSynthesisUtterance', {
				value: undefined,
				configurable: true,
			});
			return;
		}
		type Listener = (event?: unknown) => void;
		const queue: Array<{ listeners: Record<string, Listener[]> }> = [];
		const record = { spoken: [] as string[], cancels: 0 };
		class FakeUtterance {
			listeners: Record<string, Listener[]> = {};
			constructor(public text: string) {}
			addEventListener(name: string, fn: Listener) {
				(this.listeners[name] ??= []).push(fn);
			}
		}
		Object.defineProperty(window, 'SpeechSynthesisUtterance', {
			value: FakeUtterance,
			configurable: true,
		});
		Object.defineProperty(window, 'speechSynthesis', {
			configurable: true,
			value: {
				speak(utterance: FakeUtterance) {
					record.spoken.push(utterance.text);
					queue.push(utterance);
				},
				cancel() {
					record.cancels++;
					queue.length = 0;
				},
			},
		});
		(window as unknown as { __speech: unknown }).__speech = {
			get spoken() {
				return record.spoken;
			},
			get cancels() {
				return record.cancels;
			},
			fire(type: string, error?: string) {
				const utterance = queue.shift();
				for (const fn of utterance?.listeners[type] ?? []) fn(error ? { error } : {});
			},
		};
	}, supported);
}

async function open(page: Page, framework: string, supported = true) {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await stubSpeech(page, supported);
	await page.goto(`/building-blocks.html?framework=${framework}`);
	await expect(page.locator('.ProseMirror')).toContainText(TEXT);
	await expect(page.getByRole('button', { name: 'Bold', exact: true })).toBeVisible();
	return errors;
}
type Toolbar = NonNullable<ReturnType<NativeHarness['toolbar']>>;
const props = <T>(page: Page, read: (toolbar: Toolbar) => T) =>
	page.evaluate(`(${read.toString()})(window.nativeWord.toolbar())`) as Promise<T>;
const spoken = (page: Page) => page.evaluate(() => window.__speech!.spoken);
const hostButton = (page: Page) => page.locator('#host-read-aloud');
const hostState = (page: Page) => page.locator('#host-read-aloud-state');
const range = (page: Page) =>
	props(page, (toolbar) =>
		toolbar.selection ? { from: toolbar.selection.from, to: toolbar.selection.to } : null,
	);
const selectedLength = async (page: Page) => {
	const selection = await range(page);
	return selection ? selection.to - selection.from : 0;
};

for (const framework of FRAMEWORKS) {
	test.describe(`${framework}: host Read Aloud control`, () => {
		test('reads the selection and shows the playing state', async ({ page }) => {
			const errors = await open(page, framework);
			await expect(hostButton(page)).toBeEnabled();
			await expect(hostState(page)).toHaveText('idle');
			await page.locator('.ProseMirror p').dblclick({ position: { x: 20, y: 8 } });
			await expect.poll(() => selectedLength(page)).toBeGreaterThan(2);
			await hostButton(page).click();
			await expect(hostState(page)).toHaveText('reading');
			await expect(hostButton(page)).toHaveAttribute('aria-pressed', 'true');
			expect(await spoken(page)).toEqual(['Native']);
			// The stock Toolbar toggle follows the same state.
			await expect(page.getByRole('button', { name: 'Read aloud' })).toHaveAttribute(
				'aria-pressed',
				'true',
			);
			// The browser finishing the utterance returns the control to idle.
			await page.evaluate(() => window.__speech!.fire('end'));
			await expect(hostState(page)).toHaveText('idle');
			await expect(hostButton(page)).toHaveAttribute('aria-pressed', 'false');
			expect(errors).toEqual([]);
		});

		test('reads from the caret to the end without a selection', async ({ page }) => {
			await open(page, framework);
			await page.locator('.ProseMirror p').click({ position: { x: 5, y: 8 } });
			await page.keyboard.press('Home');
			for (let i = 0; i < 7; i++) await page.keyboard.press('ArrowRight');
			await expect.poll(async () => (await range(page))?.from).toBe(8);
			await hostButton(page).click();
			await expect(hostState(page)).toHaveText('reading');
			expect(await spoken(page)).toEqual(['Word composition']);
		});

		test('stops on request and reports a failing voice', async ({ page }) => {
			await open(page, framework);
			await page.locator('.ProseMirror p').click({ position: { x: 5, y: 8 } });
			await page.keyboard.press('Home');
			await hostButton(page).click();
			await expect(hostState(page)).toHaveText('reading');
			await hostButton(page).click();
			await expect(hostState(page)).toHaveText('idle');
			expect(await page.evaluate(() => window.__speech!.cancels)).toBeGreaterThan(0);
			await expect(page.getByRole('button', { name: 'Read aloud' })).toHaveAttribute(
				'aria-pressed',
				'false',
			);
			await hostButton(page).click();
			await expect(hostState(page)).toHaveText('reading');
			await page.evaluate(() => window.__speech!.fire('error', 'synthesis-failed'));
			await expect(hostState(page)).toHaveText('error');
			expect(await props(page, (toolbar) => toolbar.readAloudIssue)).toBe('error');
			// A deliberate cancel is not a failure.
			await hostButton(page).click();
			await expect(hostState(page)).toHaveText('reading');
			await page.evaluate(() => window.__speech!.fire('error', 'canceled'));
			await expect(hostState(page)).toHaveText('idle');
		});

		test('reads a read-only document without dirtying it or moving the selection', async ({
			page,
		}) => {
			await open(page, framework);
			await page.evaluate(() => window.nativeWord!.permission(false));
			await expect.poll(() => page.evaluate(() => window.nativeWord!.mode())).toBe('view');
			await page.locator('.ProseMirror p').dblclick({ position: { x: 20, y: 8 } });
			await expect.poll(() => selectedLength(page)).toBeGreaterThan(2);
			const before = await range(page);
			await expect(hostButton(page)).toBeEnabled();
			await hostButton(page).click();
			await expect(hostState(page)).toHaveText('reading');
			expect(await spoken(page)).toEqual(['Native']);
			expect(await range(page)).toEqual(before);
			expect(await page.evaluate(() => window.nativeWord!.handle.dirty)).toBe(false);
		});

		test('keeps the selection and focus when the host control is clicked', async ({ page }) => {
			await open(page, framework);
			await page.locator('.ProseMirror p').dblclick({ position: { x: 20, y: 8 } });
			await expect.poll(() => selectedLength(page)).toBeGreaterThan(2);
			const before = await range(page);
			await hostButton(page).click();
			await expect(hostState(page)).toHaveText('reading');
			expect(await range(page)).toEqual(before);
			expect(
				await page.evaluate(() => document.activeElement?.closest('docx-editor') !== null),
			).toBe(true);
		});

		test('an edit stops playback', async ({ page }) => {
			await open(page, framework);
			await page.locator('.ProseMirror p').click({ position: { x: 5, y: 8 } });
			await page.keyboard.press('Home');
			await hostButton(page).click();
			await expect(hostState(page)).toHaveText('reading');
			await page.keyboard.type('x');
			await expect(page.locator('.ProseMirror')).toContainText('xNative');
			await expect(hostState(page)).toHaveText('idle');
			expect(await props(page, (toolbar) => toolbar.readingAloud)).toBe(false);
			expect(await page.evaluate(() => window.__speech!.cancels)).toBeGreaterThan(0);
		});

		test('reports unsupported and disables the control without speech synthesis', async ({
			page,
		}) => {
			const errors = await open(page, framework, false);
			await expect(hostState(page)).toHaveText('unsupported');
			await expect(hostButton(page)).toBeDisabled();
			const state = await props(page, (toolbar) => ({
				can: toolbar.canReadAloud,
				issue: toolbar.readAloudIssue,
				reading: toolbar.readingAloud,
				start: toolbar.onReadAloud(),
			}));
			expect(state).toEqual({
				can: false,
				issue: 'unsupported',
				reading: false,
				start: 'unsupported',
			});
			// The stock Toolbar hides its toggle where the browser cannot speak.
			await expect(page.getByRole('button', { name: 'Read aloud' })).toHaveCount(0);
			expect(errors).toEqual([]);
		});
	});
}
