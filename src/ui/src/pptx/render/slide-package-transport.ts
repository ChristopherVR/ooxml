/**
 * System clipboard transport for slide packages, with the in-page fallback Safari needs.
 *
 * `createSlidePackageClipboard` owns the package operations; these helpers only move bytes and text
 * between it, the system clipboard and the page's store.
 */
import { deserializeSlidePackageClipboard, SLIDE_PACKAGE_CLIPBOARD_MIME } from 'ooxml-core/pptx';
import { readInPageSlidePackage } from './slide-package-page-store';

/**
 * Offer the copied slides' text (and picture) to other applications when the package itself cannot
 * go to the system clipboard. Best effort: it never rejects, since the package is kept in the page.
 * The first write starts synchronously so it stays inside the user's gesture.
 */
export async function writeSystemPreview(
	clipboard: Clipboard | undefined,
	canWriteItems: boolean,
	preview: string,
	picture: Promise<Blob> | undefined,
): Promise<void> {
	if (canWriteItems && clipboard) {
		const formats: Record<string, Blob | Promise<Blob>> = {
			'text/plain': new Blob([preview], { type: 'text/plain' }),
		};
		if (picture) formats['image/png'] = picture;
		try {
			await clipboard.write([new ClipboardItem(formats)]);
			return;
		} catch {
			// Fall through to plain text below.
		}
	}
	try {
		await clipboard?.writeText?.(preview);
	} catch {
		// The package is still pasteable from the page.
	}
}

/** The slide package on the system clipboard, or `null` when it holds something else. */
async function readSystemSlidePackage(clipboard: Clipboard): Promise<Uint8Array | null> {
	const items = clipboard.read ? await clipboard.read() : [];
	const item = items.find((value) => value.types.includes(SLIDE_PACKAGE_CLIPBOARD_MIME));
	const text = item
		? await (await item.getType(SLIDE_PACKAGE_CLIPBOARD_MIME)).text()
		: clipboard.readText
			? await clipboard.readText()
			: '';
	return deserializeSlidePackageClipboard(text);
}

/**
 * The slide package to paste: the system clipboard's first (another tab or window may have copied
 * it), then the one copied in this page. A system read that fails (no permission, Safari's paste
 * prompt dismissed) is reported only when the page holds no package either.
 */
export async function readSlidePackage(clipboard: Clipboard | undefined): Promise<Uint8Array> {
	let systemError: unknown;
	if (clipboard) {
		try {
			const source = await readSystemSlidePackage(clipboard);
			if (source) return source;
		} catch (error) {
			systemError = error;
		}
	}
	const stored = readInPageSlidePackage();
	const source = stored ? deserializeSlidePackageClipboard(stored.text) : null;
	if (source) return source;
	if (systemError) throw systemError;
	if (!clipboard) throw new Error('System clipboard is unavailable in this browser');
	throw new Error('The clipboard does not contain copied PowerPoint slides');
}
