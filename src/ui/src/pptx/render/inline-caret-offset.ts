/**
 * A one-shot caret position by body text offset for the inline text editor that opens next.
 * A rich paste ends the edit, writes the pasted runs to the model and reopens the editor; the
 * caret then belongs after the pasted text, as in PowerPoint. `placeCaretAt` consumes it.
 */
import { readEditableText } from './inline-text-extract';

const PENDING_OFFSET_TTL_MS = 1000;
let pending: { elementId: string; offset: number; at: number } | null = null;

/** Remember where the caret goes when the editor for `elementId` next mounts. */
export function setPendingCaretOffset(elementId: string, offset: number): void {
	pending = { elementId, offset: Math.max(0, Math.trunc(offset)), at: Date.now() };
}

/** The recorded offset, once; `null` when none, expired, or for another element's editor. */
export function takePendingCaretOffset(elementId?: string): number | null {
	const value = pending;
	pending = null;
	if (!value || Date.now() - value.at > PENDING_OFFSET_TTL_MS) {
		return null;
	}
	return elementId === undefined || elementId === value.elementId ? value.offset : null;
}

/** Whether an offset is waiting (without consuming it). */
export function hasPendingCaretOffset(): boolean {
	return pending !== null && Date.now() - pending.at <= PENDING_OFFSET_TTL_MS;
}

const MAX_NODES = 4000;

/**
 * Collapse the selection at body offset `offset` of `el`, counted the way `readEditableText`
 * reads the surface (paragraph and line breaks are one character, list markers none).
 */
export function placeCaretAtTextOffset(el: HTMLElement, offset: number): boolean {
	const doc = el.ownerDocument;
	const selection = doc.defaultView?.getSelection();
	if (!selection) {
		return false;
	}
	const walker = doc.createTreeWalker(el, 4);
	let node = walker.nextNode() as Text | null;
	let count = 0;
	let target: { node: Node; offset: number } | undefined;
	while (node && count++ < MAX_NODES) {
		if (!node.parentElement?.closest('[data-pptx-bullet-marker]')) {
			const prefix = doc.createRange();
			prefix.selectNodeContents(el);
			prefix.setEnd(node, 0);
			const before = readEditableText(prefix.cloneContents()).length;
			if (offset >= before && offset <= before + node.length) {
				target = { node, offset: offset - before };
				break;
			}
		}
		node = walker.nextNode() as Text | null;
	}
	const range = doc.createRange();
	if (target) {
		range.setStart(target.node, target.offset);
	} else {
		range.selectNodeContents(el);
		range.collapse(false);
	}
	range.collapse(true);
	selection.removeAllRanges();
	selection.addRange(range);
	return Boolean(target);
}
