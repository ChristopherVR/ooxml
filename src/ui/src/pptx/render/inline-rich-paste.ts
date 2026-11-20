/**
 * Paste into the inline text editor (`[data-inline-editor]` contenteditable), shared by every
 * binding. One listener on the viewer root:
 *
 * - clipboard HTML with formatting, bullets or alignment becomes runs (`html-text-paste.ts`),
 *   spliced into the body at the selection; the editor leaves (its blur commits the draft) and
 *   the body is written through the binding's element update (one Undo step); the editor then
 *   reopens with the caret after the pasted text;
 * - anything else, and every paste after Ctrl/Cmd+Shift+V (Keep Text Only), is inserted as plain
 *   text, never as raw markup.
 *
 * Table cell editors (`<input>`) and the notes and outline editors are not inline text editors
 * and are left alone.
 *
 * @module render/inline-rich-paste
 */
import type { PptxElement, TextSegment, TextStyle } from 'ooxml-core/pptx';
import { hasTextProperties } from 'ooxml-core/pptx';

import { setPendingCaretOffset } from './inline-caret-offset';
import { readListActivationSelection } from './inline-list-activation';
import { getActiveInlineListSelection } from './inline-list-controller';
import { readEditableText } from './inline-text-extract';
import {
	htmlToPasteSegments,
	segmentsBodyText,
	spliceTextSegments,
	styleAtOffset,
	type TextPasteFormat,
} from './html-text-paste';
import { remapTextToSegments } from './remap-text';

/** The body after a rich paste, and where the caret goes. */
export interface InlineRichPasteResult {
	elementId: string;
	text: string;
	textSegments: TextSegment[];
	caret: number;
	format: TextPasteFormat;
}

function currentBody(
	root: HTMLElement,
	element: PptxElement & { textSegments?: TextSegment[]; textStyle?: TextStyle },
): { segments: TextSegment[]; range: { start: number; end: number } } | undefined {
	const list = getActiveInlineListSelection(root.ownerDocument.defaultView?.getSelection() ?? null);
	if (list) {
		if (list.kind !== 'supported' || !list.bodyRange || list.snapshot.elementId !== element.id) {
			return undefined;
		}
		return { segments: list.snapshot.textSegments ?? [], range: list.bodyRange };
	}
	const text = readEditableText(root);
	const range = readListActivationSelection(root, text);
	if (!range) {
		return undefined;
	}
	const model = element.textSegments?.length ? segmentsBodyText(element.textSegments) : undefined;
	const segments =
		model === text && element.textSegments
			? element.textSegments
			: remapTextToSegments(text, element.textSegments, element.textStyle);
	return { segments, range };
}

/**
 * The rich paste of `html` into the editor `root` showing `element`, or `undefined` when the
 * clipboard has nothing a plain-text paste would lose (or the selection cannot be read).
 */
export function resolveInlineRichPaste(
	root: HTMLElement,
	element: PptxElement | null | undefined,
	html: string,
	format: TextPasteFormat = 'keep-source-formatting',
): InlineRichPasteResult | undefined {
	if (!html || format === 'keep-text-only' || !element || !hasTextProperties(element)) {
		return undefined;
	}
	const body = currentBody(root, element);
	if (!body) {
		return undefined;
	}
	const start = Math.min(body.range.start, body.range.end);
	const end = Math.max(body.range.start, body.range.end);
	const destination = styleAtOffset(body.segments, start, element.textStyle ?? {});
	const pasted = htmlToPasteSegments(html, destination, format);
	if (!pasted?.rich) {
		return undefined;
	}
	const textSegments = spliceTextSegments(body.segments, start, end, pasted.segments);
	return {
		elementId: element.id,
		text: segmentsBodyText(textSegments),
		textSegments,
		caret: start + pasted.text.length,
		format,
	};
}

/** What the viewer supplies. */
export interface InlineRichPasteOptions {
	/** The element whose inline text editor is open, or null. */
	getEditingElement(): PptxElement | null | undefined;
	/**
	 * Make sure the editor is closed. The surface was just blurred, so the binding's blur
	 * commit has already run; this must not commit again.
	 */
	cancelEdit(): void;
	/** Write the patch through the editor's normal, history-tracked element update. */
	updateElement(elementId: string, patch: { text: string; textSegments: TextSegment[] }): void;
	/** Reopen the inline editor on the element. */
	enterEdit(elementId: string): void;
	/** Override the paste format (defaults to Keep Source Formatting). */
	getFormat?(): TextPasteFormat;
}

/**
 * Type `text` at the caret of the focused inline editor (a paste, a Tab). `insertText` keeps
 * native Undo and the editors' input pipeline; a Range is the fallback.
 */
export function insertInlineEditorText(doc: Document, text: string): void {
	if (typeof doc.execCommand === 'function' && doc.execCommand('insertText', false, text)) {
		return;
	}
	const selection = doc.defaultView?.getSelection();
	const range = selection?.rangeCount ? selection.getRangeAt(0) : null;
	if (!range) {
		return;
	}
	range.deleteContents();
	const node = doc.createTextNode(text);
	range.insertNode(node);
	range.setStartAfter(node);
	range.collapse(true);
	selection?.removeAllRanges();
	selection?.addRange(range);
	node.parentElement?.dispatchEvent(new Event('input', { bubbles: true }));
}

function editorOf(target: EventTarget | null): HTMLElement | null {
	const element = target as Element | null;
	const editor = element?.nodeType === 1 ? element.closest('[data-inline-editor]') : null;
	if (!(editor instanceof HTMLElement)) {
		return null;
	}
	const editable = editor.getAttribute('contenteditable');
	return editor.isContentEditable ||
		editable === '' ||
		editable === 'true' ||
		editable === 'plaintext-only'
		? editor
		: null;
}

function editorFromEvent(event: Event): HTMLElement | null {
	const target = event.target as Node | null;
	return editorOf(target?.nodeType === 3 ? target.parentElement : target);
}

/** Attach once per viewer root; returns the detach function. */
export function attachInlineRichPaste(
	root: HTMLElement,
	options: InlineRichPasteOptions,
): () => void {
	let plainNext = false;
	const onKeyDown = (event: KeyboardEvent): void => {
		plainNext =
			(event.ctrlKey || event.metaKey) &&
			event.shiftKey &&
			event.key.toLowerCase() === 'v' &&
			editorFromEvent(event) !== null;
	};
	const onPaste = (event: ClipboardEvent): void => {
		const editor = editorFromEvent(event);
		if (!editor || event.defaultPrevented || !root.contains(editor)) {
			return;
		}
		event.preventDefault();
		const plain = plainNext;
		plainNext = false;
		const data = event.clipboardData;
		const element = options.getEditingElement();
		const format = plain ? 'keep-text-only' : (options.getFormat?.() ?? 'keep-source-formatting');
		const result = resolveInlineRichPaste(
			editor,
			element,
			data?.getData('text/html') ?? '',
			format,
		);
		if (result) {
			// Leave the editor the ordinary way first: its blur commits the draft (which the
			// result already contains) now, so no blur fired later by unmounting the focused
			// surface can commit the pre-paste draft over the pasted body.
			editor.blur();
			options.cancelEdit();
			options.updateElement(result.elementId, {
				text: result.text,
				textSegments: result.textSegments,
			});
			const reopen = () => {
				setPendingCaretOffset(result.elementId, result.caret);
				options.enterEdit(result.elementId);
			};
			// The binding's own state settles (editor unmounted, model written) before it reopens.
			setTimeout(reopen, 0);
			return;
		}
		const text = data?.getData('text/plain') ?? '';
		if (text) {
			insertInlineEditorText(editor.ownerDocument, text.replace(/\r\n?/g, '\n'));
		}
	};
	root.addEventListener('keydown', onKeyDown, true);
	root.addEventListener('paste', onPaste);
	return () => {
		root.removeEventListener('keydown', onKeyDown, true);
		root.removeEventListener('paste', onPaste);
	};
}
