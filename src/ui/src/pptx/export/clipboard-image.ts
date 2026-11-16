/**
 * Copy a rendered PNG to the system clipboard (File > Copy Slide as Image).
 *
 * Safari only lets `navigator.clipboard.write` run while the user's click is
 * still being handled. Rendering a slide takes longer than that, so awaiting the
 * PNG first and then writing it was rejected with `NotAllowedError` on Safari and
 * iPadOS. The clipboard item is therefore created and written at once, with the
 * PNG as a promise the browser waits for. A browser whose `ClipboardItem`
 * refuses a promise gets the PNG once it is ready.
 */

/** Whether this browser can put an image on the clipboard. */
export function canCopyImageToClipboard(): boolean {
	return (
		typeof ClipboardItem !== 'undefined' &&
		typeof navigator !== 'undefined' &&
		typeof navigator.clipboard?.write === 'function'
	);
}

/**
 * Write the PNG that `png` resolves to onto the clipboard. Call it directly
 * from the click handler, before any `await`, and pass the render as a promise.
 */
export async function copyPngToClipboard(png: Promise<Blob>): Promise<void> {
	if (!canCopyImageToClipboard()) {
		throw new Error('Image clipboard is unavailable');
	}
	let item: ClipboardItem | undefined;
	try {
		item = new ClipboardItem({ 'image/png': png });
	} catch {
		// Older engines accept only a ready Blob.
	}
	await navigator.clipboard.write([item ?? new ClipboardItem({ 'image/png': await png })]);
}
