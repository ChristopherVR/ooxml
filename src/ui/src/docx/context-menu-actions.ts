/** Clipboard operations for the context menu. Nothing here pretends a denied permission succeeded. */
import { deleteFieldSelection } from 'ooxml-core/docx/ui';
import type { EditorView } from 'prosemirror-view';
import { pastePayload } from './paste-options';
import { deleteCollapsedOutlineBranches } from './view-outline-selection';

/**
 * Cut or copy through the browser's own command so ProseMirror's clipboard serializer keeps rich
 * content; the async fallback keeps its HTML slice when that API is available.
 */
export async function copyOrCut(view: EditorView, op: 'cut' | 'copy'): Promise<boolean> {
	if (view.state.selection.empty || (op === 'cut' && !view.editable)) return false;
	view.focus();
	if (typeof document.execCommand === 'function' && document.execCommand(op)) return true;
	if (typeof navigator === 'undefined' || !navigator.clipboard) return false;
	const { doc, selection } = view.state;
	const { dom, text } = view.serializeForClipboard(selection.content());
	try {
		if (navigator.clipboard.write && typeof ClipboardItem !== 'undefined') {
			await navigator.clipboard.write([
				new ClipboardItem({
					'text/html': new Blob([dom.innerHTML], { type: 'text/html' }),
					'text/plain': new Blob([text], { type: 'text/plain' }),
				}),
			]);
		} else if (navigator.clipboard.writeText) await navigator.clipboard.writeText(text);
		else return false;
	} catch {
		return false;
	}
	// A permission prompt can outlive the selection: never delete newly selected content.
	if (
		op === 'cut' &&
		!view.isDestroyed &&
		view.editable &&
		view.state.doc.eq(doc) &&
		view.state.selection.eq(selection)
	)
		view.dispatch(
			(
				deleteCollapsedOutlineBranches(view.state) ??
				deleteFieldSelection(view.state) ??
				view.state.tr.deleteSelection()
			)
				.scrollIntoView()
				.setMeta('uiEvent', 'cut'),
		);
	return true;
}

/** Whether the browser will let us read the clipboard; `true` when it cannot say. */
export async function clipboardReadAllowed(): Promise<boolean> {
	if (
		typeof navigator === 'undefined' ||
		(!navigator.clipboard?.read && !navigator.clipboard?.readText)
	)
		return false;
	try {
		const status = await navigator.permissions?.query({ name: 'clipboard-read' as PermissionName });
		return status?.state !== 'denied';
	} catch {
		return true;
	}
}

/** Reads the async clipboard, asking for Word's HTML as written (stylesheet and markers kept). */
async function readClipboardItems(clipboard: Clipboard): Promise<ClipboardItems> {
	const read = clipboard.read.bind(clipboard) as (options?: object) => Promise<ClipboardItems>;
	try {
		return await read({ unsanitized: ['text/html'] });
	} catch (error) {
		// A browser that rejects the option (rather than the permission) reads sanitised HTML.
		if (error instanceof TypeError) return read();
		throw error;
	}
}

/** The clipboard's HTML and text (`null` for a type it does not hold). */
export interface ClipboardContents {
	html: string | null;
	text: string | null;
}

/**
 * What the clipboard holds as HTML and as text, or `null` when the browser has no clipboard
 * API. Throws when it refuses to read (a denied permission).
 */
export async function readClipboardContents(): Promise<ClipboardContents | null> {
	const clipboard = typeof navigator === 'undefined' ? undefined : navigator.clipboard;
	if (!clipboard) return null;
	if (!clipboard.read) return { html: null, text: await clipboard.readText() };
	const contents: ClipboardContents = { html: null, text: null };
	for (const item of await readClipboardItems(clipboard)) {
		if (contents.html === null && item.types.includes('text/html'))
			contents.html = await (await item.getType('text/html')).text();
		if (contents.text === null && item.types.includes('text/plain'))
			contents.text = await (await item.getType('text/plain')).text();
	}
	return contents;
}

/**
 * Uses ProseMirror's HTML parser for formatted content, with a plain-text API fallback. What was
 * read is recorded, so Paste Options can paste it again another way.
 */
export async function pasteClipboard(view: EditorView, plainText = false): Promise<boolean> {
	if (!view.editable) return false;
	try {
		const contents = await readClipboardContents();
		if (!contents || !view.editable || view.isDestroyed) return false;
		const payload = { html: contents.html, text: contents.text ?? '' };
		if (!plainText && contents.html !== null)
			return contents.html ? pastePayload(view, payload, 'source') : true;
		if (contents.text === null) return false;
		return contents.text ? pastePayload(view, payload, 'text') : true;
	} catch {
		return false;
	}
}
