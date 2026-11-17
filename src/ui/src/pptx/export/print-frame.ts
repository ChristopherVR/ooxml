/**
 * Print one complete HTML document (the notes page, or a binding's print job) without leaving
 * the page, shared by every binding.
 *
 * Desktop browsers print a hidden same-origin iframe. iOS and iPadOS do not: `print()` on an
 * iframe's window prints the parent page there, so WebKit on those devices gets a new window
 * instead, opened synchronously so it stays inside the user's click (see `print-window.ts`).
 */
import { currentWebKitEnvironment, isIosWebKit } from '../render/webkit-engine';
import type { WebKitEnvironment } from '../render/webkit-engine';
import { openPrintWindow } from './print-window';

/** How long the hidden print iframe survives if `afterprint` never fires. */
export const PRINT_FRAME_FALLBACK_TTL_MS = 60_000;

/** Options for {@link printHtmlDocument}. */
export interface PrintHtmlOptions {
	/** The document the hidden iframe is mounted in (defaults to `document`). */
	doc?: Document;
	/** The browser to decide for (defaults to `navigator`); a test seam. */
	env?: WebKitEnvironment;
}

/** Whether this browser must print through a window rather than an iframe (iOS and iPadOS). */
export function printNeedsWindow(env: WebKitEnvironment | undefined = currentWebKitEnvironment()) {
	return isIosWebKit(env);
}

/**
 * Print `htmlDocument`. Returns `false` when the print surface could not be opened (on iOS, a
 * blocked popup), so the caller can tell the user to allow popups.
 */
export function printHtmlDocument(htmlDocument: string, options: PrintHtmlOptions = {}): boolean {
	if (printNeedsWindow(options.env ?? currentWebKitEnvironment())) {
		return openPrintWindow(htmlDocument);
	}
	return printInHiddenFrame(options.doc ?? document, htmlDocument);
}

/**
 * Mount the document in a hidden same-origin iframe (through `srcdoc`, no `document.write`) and
 * print it once loaded. The frame may show modal dialogs (print) but runs no scripts, and it is
 * removed after `afterprint`, or after {@link PRINT_FRAME_FALLBACK_TTL_MS} where that never fires.
 */
function printInHiddenFrame(doc: Document, htmlDocument: string): boolean {
	const body = doc.body;
	if (!body) return false;
	const frame = doc.createElement('iframe');
	frame.setAttribute('aria-hidden', 'true');
	frame.setAttribute('sandbox', 'allow-same-origin allow-modals');
	frame.style.cssText =
		'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden';
	const remove = () => frame.remove();
	frame.addEventListener(
		'load',
		() => {
			const win = frame.contentWindow;
			if (!win) {
				remove();
				return;
			}
			win.addEventListener('afterprint', () => setTimeout(remove, 500), { once: true });
			win.focus();
			win.print();
		},
		{ once: true },
	);
	frame.srcdoc = htmlDocument;
	body.appendChild(frame);
	setTimeout(remove, PRINT_FRAME_FALLBACK_TTL_MS);
	return true;
}
