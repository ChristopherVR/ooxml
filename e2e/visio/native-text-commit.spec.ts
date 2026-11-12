import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { createVsdxFixture } from './fixture.mjs';
import type { TextEditOutcome } from 'ooxml-ui/visio';

// The harness' host Save, next to the `nativeBlocks.controller` other specs declare on `Window`.
interface HostSaveWindow {
	nativeBlocks: {
		save(): Promise<{ outcome: TextEditOutcome; bytes: number }>;
		readonly lastExport: number[] | undefined;
	};
}

// A host-owned Save (the harness' Save button, Ctrl+S and `nativeBlocks.save()`) awaits
// `commitTextEdit()` before `exportVsdx()`, with no reliance on blur.
const frameworks = ['react', 'vue', 'angular', 'svelte', 'solid', 'vanilla'];
const SHAPE = 'svg.paper [data-shape-id="1"]';

async function openFile(page: Page, text: string, name: string) {
	await page.locator('#content').setInputFiles({
		name,
		mimeType: 'application/octet-stream',
		buffer: await createVsdxFixture(text),
	});
	await expect(page.locator('svg.paper text')).toContainText(text);
}
async function load(page: Page, framework: string, text: string) {
	await page.goto(`/native-blocks.html?framework=${framework}`);
	await openFile(page, text, 'draft.vsdx');
	await expect(page.locator('output')).toHaveText('edit:1');
}
async function startDraft(page: Page, draft: string) {
	await page.locator(SHAPE).dblclick();
	const editor = page.locator('#edit-text');
	await expect(editor).toBeFocused();
	await editor.fill(draft);
	return editor;
}
/** The page XML of the last exported package. */
async function exportedText(page: Page): Promise<string> {
	const bytes = await page.evaluate(
		() => (window as unknown as HostSaveWindow).nativeBlocks.lastExport,
	);
	const zip = await JSZip.loadAsync(Buffer.from(bytes!));
	const parts = Object.keys(zip.files).filter((name) => /visio\/pages\/page\d+\.xml$/.test(name));
	return (await Promise.all(parts.map((name) => zip.file(name)!.async('string')))).join('\n');
}
const status = (page: Page) =>
	page.locator('#save-status').evaluate((el) => JSON.parse(el.textContent || '{}'));

for (const framework of frameworks) {
	test.describe(framework, () => {
		test.describe.configure({ mode: 'parallel' });
		test('Save called without a blur commits, exports and undo reverts', async ({ page }) => {
			await load(page, framework, 'Original');
			const editor = await startDraft(page, 'Typed then saved');
			// Called directly: the textarea keeps focus, so nothing blurred first.
			const result = await page.evaluate(() =>
				(window as unknown as HostSaveWindow).nativeBlocks.save(),
			);
			await expect(editor).toHaveCount(0);
			expect(result.outcome).toMatchObject({
				status: 'committed',
				text: 'Typed then saved',
				shapeId: '1',
			});
			const outcome = result.outcome as Extract<TextEditOutcome, { status: 'committed' }>;
			expect(outcome.pageId).toBeTruthy();
			expect(outcome.sourceGeneration).toBeGreaterThan(0);
			expect(outcome.documentGeneration).toBeGreaterThan(0);
			expect(result.bytes).toBeGreaterThan(0);
			expect(await exportedText(page)).toContain('Typed then saved');
			await expect(page.locator('svg.paper text')).toContainText('Typed then saved');
			await page.getByRole('button', { name: 'Undo', exact: true }).click();
			await expect(page.locator('svg.paper text')).toContainText('Original');
			expect(
				await page.evaluate(() => (window as unknown as HostSaveWindow).nativeBlocks.save()),
			).toMatchObject({
				outcome: { status: 'unchanged', reason: 'no-draft' },
			});
			expect(await exportedText(page)).not.toContain('Typed then saved');
		});

		test('the Save button and Ctrl+S commit the visible draft', async ({ page }) => {
			await load(page, framework, 'Original');
			await startDraft(page, 'Via button');
			await page.locator('#save').click();
			await expect(page.locator('#save-status')).toHaveAttribute('data-status', 'committed');
			expect(await status(page)).toMatchObject({ text: 'Via button', shapeId: '1' });
			expect(await exportedText(page)).toContain('Via button');
			// The keyboard shortcut, with the textarea focused.
			const editor = await startDraft(page, 'Via shortcut');
			await editor.press('Control+s');
			await expect(editor).toHaveCount(0);
			await expect.poll(async () => (await status(page)).text).toBe('Via shortcut');
			expect((await status(page)).status).toBe('committed');
			expect(await exportedText(page)).toContain('Via shortcut');
		});

		test('Save during a commit that Escape already started awaits it', async ({ page }) => {
			await load(page, framework, 'Original');
			await startDraft(page, 'In flight');
			// Escape starts the commit; Save is called before it has settled.
			const result = await page.locator('#edit-text').evaluate((input) => {
				input.dispatchEvent(
					new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
				);
				return (window as unknown as HostSaveWindow).nativeBlocks.save();
			});
			expect(result.outcome).toMatchObject({ status: 'committed', text: 'In flight' });
			expect(await exportedText(page)).toContain('In flight');
			// One undo step: the draft was not applied twice.
			await page.getByRole('button', { name: 'Undo', exact: true }).click();
			await expect(page.locator('svg.paper text')).toContainText('Original');
		});

		test('Save with no draft is unchanged', async ({ page }) => {
			await load(page, framework, 'Original');
			const result = await page.evaluate(() =>
				(window as unknown as HostSaveWindow).nativeBlocks.save(),
			);
			expect(result.outcome).toMatchObject({ status: 'unchanged', reason: 'no-draft' });
			expect(result.bytes).toBeGreaterThan(0);
		});

		test('a read-only drawing refuses the draft and keeps its text', async ({ page }) => {
			await load(page, framework, 'Original');
			await startDraft(page, 'Kept for recovery');
			await page.locator('#readonly').click();
			await expect(page.locator('output')).toHaveText('view:1');
			const result = await page.evaluate(() =>
				(window as unknown as HostSaveWindow).nativeBlocks.save(),
			);
			expect(result.outcome).toMatchObject({
				status: 'refused',
				reason: 'read-only',
				text: 'Kept for recovery',
			});
			await expect(page.locator('svg.paper text')).toContainText('Original');
		});

		test('a document switched mid-draft is never modified by the old draft', async ({ page }) => {
			await load(page, framework, 'First');
			await startDraft(page, 'Belongs to first');
			await openFile(page, 'Second', 'second.vsdx');
			const result = await page.evaluate(() =>
				(window as unknown as HostSaveWindow).nativeBlocks.save(),
			);
			expect(result.outcome).toMatchObject({
				status: 'cancelled',
				reason: 'stale-document',
				text: 'Belongs to first',
				shapeId: '1',
			});
			await expect(page.locator('svg.paper text')).toContainText('Second');
			expect(await exportedText(page)).not.toContain('Belongs to first');
			// Reported once.
			expect(
				await page.evaluate(() => (window as unknown as HostSaveWindow).nativeBlocks.save()),
			).toMatchObject({
				outcome: { status: 'unchanged', reason: 'no-draft' },
			});
		});
	});
}
