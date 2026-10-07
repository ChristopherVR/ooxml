import type { EditorView } from 'prosemirror-view';
import {
	applyWordLink,
	removeWordLink,
	wordLinkRangeAt,
	isNavigableHref,
	type LinkTarget,
} from 'ooxml-core/docx/ui';

export type { LinkTarget } from 'ooxml-core/docx/ui';
export const linkRangeAt = (view: EditorView, pos: number) => wordLinkRangeAt(view.state, pos);

export function linkAtSelection(view: EditorView): LinkTarget | undefined {
	const range = linkRangeAt(view, view.state.selection.from);
	if (!range) return undefined;
	const { href, anchor, tooltip } = range.mark.attrs;
	return {
		...(href ? { href } : {}),
		...(anchor ? { anchor } : {}),
		...(tooltip ? { tooltip } : {}),
	};
}

export function applyLink(view: EditorView, target: LinkTarget, displayText?: string): boolean {
	return applyWordLink(target, displayText)(view.state, view.dispatch, view);
}

export function removeLink(view: EditorView): boolean {
	return removeWordLink(view.state, view.dispatch, view);
}

/** Scrolls to the paragraph whose bookmarks include `name`; returns whether one was found. */
export function goToBookmark(view: EditorView, name: string): boolean {
	let target = -1;
	view.state.doc.descendants((node, pos) => {
		if (target >= 0) return false;
		if (
			node.type.name === 'paragraph' &&
			(node.attrs.bookmarks as string[] | undefined)?.includes(name)
		)
			target = pos;
		return true;
	});
	if (target < 0) return false;
	const dom = view.nodeDOM(target);
	if (dom instanceof HTMLElement) dom.scrollIntoView({ block: 'center' });
	return true;
}

/** Bookmark names in document order, for the "Place in this document" list. */
export function bookmarkNames(view: EditorView): string[] {
	const names: string[] = [];
	view.state.doc.descendants((node) => {
		if (node.type.name === 'paragraph') names.push(...((node.attrs.bookmarks as string[]) ?? []));
		return node.type.name !== 'paragraph';
	});
	return names.filter((name) => !name.startsWith('_'));
}

/**
 * Follows the link at `pos` like Word's Ctrl+Click: web and mail links open in a new tab,
 * internal links scroll to their bookmark. Returns whether a link was followed.
 */
export function followLinkAt(view: EditorView, pos: number): boolean {
	const range = linkRangeAt(view, pos);
	if (!range) return false;
	const { href, anchor } = range.mark.attrs;
	if (isNavigableHref(href)) {
		window.open(href, '_blank', 'noopener,noreferrer');
		return true;
	}
	return anchor ? goToBookmark(view, String(anchor)) : false;
}
