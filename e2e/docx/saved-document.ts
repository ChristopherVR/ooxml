import type { Page } from '@playwright/test';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import JSZip from 'jszip';

/**
 * The document the editor would save, as bytes and as `word/document.xml`. When `DOCX_E2E_OUT` names
 * a folder the bytes are also written there as `name`, so a Windows run can open them in Word.
 */
export async function savedDocument(
	page: Page,
	name: string,
): Promise<{ bytes: Uint8Array; xml: string }> {
	const list = await page
		.locator('docx-editor')
		.evaluate(async (element) =>
			Array.from(await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
	const bytes = new Uint8Array(list);
	const out = process.env.DOCX_E2E_OUT;
	if (out) {
		await mkdir(out, { recursive: true });
		await writeFile(join(out, name), bytes);
	}
	const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
	return { bytes, xml };
}
