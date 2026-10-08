import { expect, test } from '@playwright/test';
import { openSample } from './helpers';

// Every framework adapter exposes the same imperative handle (`EDITOR_HANDLE_KEYS` in the
// bindings). The demo drives the editor through that handle and leaves it on its host as
// `docxEditorHandle`, so this checks the real React ref, Vue template ref, Angular component,
// Svelte exports and Solid `editorRef` in a browser, not a copy built by the demo.
const HANDLE_KEYS = ['element', 'load', 'save', 'download', 'markClean', 'dirty'];

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: the binding handle exposes a live element and dirty state`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await openSample(page, framework);
		const host = page.locator('#editor');
		const shape = () =>
			host.evaluate((element, keys) => {
				const handle = (element as HTMLElement & { docxEditorHandle?: Record<string, unknown> })
					.docxEditorHandle;
				if (!handle) return null;
				const editor = element.querySelector('docx-editor');
				return {
					missing: keys.filter((key) => !(key in handle)),
					sameElement: handle.element === editor,
					dirty: handle.dirty,
					methods: ['load', 'save', 'download', 'markClean'].every(
						(key) => typeof handle[key] === 'function',
					),
				};
			}, HANDLE_KEYS);
		await expect
			.poll(shape)
			.toEqual({ missing: [], sameElement: true, dirty: false, methods: true });

		// `dirty` is live: an edit sets it, `markClean()` through the handle clears it.
		await page.locator('docx-editor .ProseMirror').click();
		await page.keyboard.press('Control+End');
		await page.keyboard.type(' edited');
		await expect.poll(async () => (await shape())?.dirty).toBe(true);
		await host.evaluate((element) =>
			(
				element as HTMLElement & { docxEditorHandle: { markClean(): void } }
			).docxEditorHandle.markClean(),
		);
		await expect.poll(async () => (await shape())?.dirty).toBe(false);

		// `save()` through the handle returns the document bytes (a DOCX zip).
		const signature = await host.evaluate(async (element) => {
			const handle = (element as HTMLElement & { docxEditorHandle: { save(): Promise<unknown> } })
				.docxEditorHandle;
			const saved = await handle.save();
			const bytes =
				saved instanceof Blob ? new Uint8Array(await saved.arrayBuffer()) : (saved as Uint8Array);
			return String.fromCharCode(bytes[0]!, bytes[1]!);
		});
		expect(signature).toBe('PK');
		expect(errors).toEqual([]);
	});
}
