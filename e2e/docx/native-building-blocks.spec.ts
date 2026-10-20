import { expect, test } from '@playwright/test';
import type { NativeHarness } from '../../demos/docx/demo-vanilla/building-blocks-types';
for (const framework of ['react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: native Word composition reacts, persists and disposes`, async ({ page }) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/building-blocks.html?framework=${framework}`);
		await expect(page.locator('.ProseMirror')).toContainText('Native Word composition');
		await expect(page.locator('docx-editor')).toHaveCount(1);
		await expect(
			page.locator('office-ui-ribbon, office-ui-title-bar, office-ui-backstage'),
		).toHaveCount(0);
		const editor = page.locator('.ProseMirror');
		await editor.click();
		await page.keyboard.press('End');
		await page.getByRole('button', { name: 'Bold', exact: true }).click();
		await expect(page.getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute(
			'aria-pressed',
			'true',
		);
		await page.keyboard.type(' native bold');
		await expect(editor.locator('strong')).toContainText('native bold');
		const saved = await page.evaluate(async () => {
			const host: NativeHarness = window.nativeWord!;
			const saved = await host.handle.save();
			const dirty = host.handle.dirty;
			await host.handle.load(await saved.arrayBuffer());
			host.handle.markClean();
			return { dirty, clean: !host.handle.dirty };
		});
		expect(saved).toEqual({ dirty: true, clean: true });
		await expect(editor.locator('strong')).toContainText('native bold');
		const engine = await page.locator('docx-editor').elementHandle();
		const boldButton = page.getByRole('button', { name: 'Bold', exact: true });
		await boldButton.focus();
		const button = await boldButton.elementHandle();
		await page.evaluate(() =>
			window.nativeWord!.handle.element.dispatchEvent(new Event('editor-state-change')),
		);
		await expect(boldButton).toBeFocused();
		// Editing commands intentionally return focus to the document; the controls must keep their nodes.
		await page.evaluate(() => window.nativeWord!.toolbar()!.onUnderline());
		expect(
			await page.evaluate(
				(node) =>
					node ===
					[...document.querySelectorAll('button')].find(
						(item) => item.textContent?.trim() === 'Bold',
					),
				button,
			),
		).toBe(true);
		expect(
			await page.evaluate((node) => node === document.querySelector('docx-editor'), engine),
		).toBe(true);
		const bytes = await page.evaluate(async () =>
			Array.from(new Uint8Array(await (await window.nativeWord!.handle.save()).arrayBuffer())),
		);
		await page.evaluate(() =>
			window.nativeWord!.options({
				documentModel: {
					...window.nativeWord!.handle.element.documentModel!,
					blocks: [
						{ type: 'paragraph', id: 'replacement', runs: [{ text: 'Replacement Word document' }] },
					],
				},
			}),
		);
		await expect(editor).toContainText('Replacement Word document');
		await page.evaluate(
			(data) => window.nativeWord!.options({ bytes: new Uint8Array(data), canEdit: true }),
			bytes,
		);
		await expect(editor.locator('strong')).toContainText('native bold');
		await page.evaluate(() => window.nativeWord!.permission(false));
		expect(
			await page.evaluate((node) => node === document.querySelector('docx-editor'), engine),
		).toBe(true);
		await expect(page.getByRole('button', { name: 'Bold', exact: true })).toBeDisabled();
		expect(await page.evaluate(() => window.nativeWord!.mode())).toBe('view');
		expect(await page.evaluate(() => window.nativeWord!.toolbar()!.onItalic())).toBe(false);
		await page.evaluate(async () => {
			try {
				await window.nativeWord!.handle.load(new Uint8Array([0x50, 0x4b, 3, 4]));
			} catch {
				/* Expected invalid ZIP. */
			}
		});
		await expect.poll(() => page.evaluate(() => Boolean(window.nativeWord!.error()))).toBe(true);
		await page.evaluate(() =>
			window.nativeWord!.options({
				documentModel: { ...window.nativeWord!.handle.element.documentModel! },
				canEdit: true,
			}),
		);
		await expect(page.getByRole('button', { name: 'Bold', exact: true })).toBeEnabled();
		await expect.poll(() => page.evaluate(() => window.nativeWord!.error())).toBeNull();
		if (framework === 'angular') {
			const events = await page.evaluate(() => window.nativeWord!.events!);
			expect(events.document).toBeGreaterThan(0);
			expect(events.dirty).toBeGreaterThan(0);
			expect(events.error).toBeGreaterThan(0);
		}
		await page.evaluate(() => window.remountNativeWord!());
		await expect(page.locator('docx-editor')).toHaveCount(1);
		await expect(editor).toContainText('Native Word composition');
		await page.evaluate(() => window.nativeWord!.destroy());
		await expect(page.locator('docx-editor')).toHaveCount(0);
		expect(errors).toEqual([]);
	});
}
