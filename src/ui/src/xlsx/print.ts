import type { Workbook } from 'ooxml-core/xlsx';
import { buildPrintHtml } from './print-pages';
export { buildPrintHtml } from './print-pages';
export { cellCss } from './print-cell';
export { printRange } from 'ooxml-core/xlsx';

/** How long a print frame may stay without `afterprint` before it is removed anyway. */
export const PRINT_FRAME_TIMEOUT_MS = 60_000;
/** The frame of the previous print, per document, removed when the next print starts. */
const pendingFrames = new WeakMap<Document, () => void>();

/**
 * Prints through a hidden frame so the editor page itself is never reflowed. The frame is removed
 * on `afterprint` (some browsers return from `print()` before the dialog closes, so removing it
 * then would cancel the job), on the next print, or after a long safety timeout.
 */
export function printHtml(doc: Document, html: string): void {
	pendingFrames.get(doc)?.();
	const frame = doc.createElement('iframe');
	frame.setAttribute('aria-hidden', 'true');
	frame.tabIndex = -1;
	// The print document needs no script. The parent still reaches the frame's
	// window and document (focus, print, afterprint), so it stays same-origin;
	// allow-modals lets the frame's print dialog open.
	frame.setAttribute('sandbox', 'allow-same-origin allow-modals');
	frame.style.cssText =
		'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden';
	doc.body.append(frame);
	const target = frame.contentWindow;
	const frameDoc = frame.contentDocument;
	if (!target || !frameDoc) {
		frame.remove();
		return;
	}
	frameDoc.open();
	frameDoc.write(html);
	frameDoc.close();
	let safety: ReturnType<typeof setTimeout> | undefined;
	const cleanup = () => {
		if (safety !== undefined) clearTimeout(safety);
		target.removeEventListener('afterprint', cleanup);
		frame.remove();
		if (pendingFrames.get(doc) === cleanup) pendingFrames.delete(doc);
	};
	pendingFrames.set(doc, cleanup);
	target.addEventListener('afterprint', cleanup);
	safety = setTimeout(cleanup, PRINT_FRAME_TIMEOUT_MS);
	setTimeout(() => {
		if (!frame.isConnected) return;
		try {
			target.focus();
			target.print();
		} catch (error) {
			// Nothing awaits this timer, so report instead of throwing an uncaught error.
			cleanup();
			console.error('xlsx-editor: printing failed', error);
		}
	}, 50);
}

/** Prints the active sheet of `workbook`. */
export function printWorkbook(
	doc: Document,
	workbook: Workbook,
	sheetIndex: number,
	title: string,
): void {
	printHtml(doc, buildPrintHtml(workbook, sheetIndex, title));
}
