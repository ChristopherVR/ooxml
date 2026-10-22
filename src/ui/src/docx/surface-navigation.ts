import { TextSelection } from 'prosemirror-state';
import type { EditorCore } from './editor-core';

type Granularity = 'character' | 'word' | 'line' | 'lineboundary' | 'documentboundary';
interface Move {
	direction: 'forward' | 'backward';
	granularity: Granularity;
}

/** Caret keys for a document with no editable surface; the browser's selection does the moving. */
function moveFor(event: KeyboardEvent): Move | undefined {
	const word = event.ctrlKey || event.altKey;
	switch (event.key) {
		case 'ArrowLeft':
			return { direction: 'backward', granularity: word ? 'word' : 'character' };
		case 'ArrowRight':
			return { direction: 'forward', granularity: word ? 'word' : 'character' };
		case 'ArrowUp':
			return { direction: 'backward', granularity: 'line' };
		case 'ArrowDown':
			return { direction: 'forward', granularity: 'line' };
		case 'Home':
			return {
				direction: 'backward',
				granularity: event.ctrlKey ? 'documentboundary' : 'lineboundary',
			};
		case 'End':
			return {
				direction: 'forward',
				granularity: event.ctrlKey ? 'documentboundary' : 'lineboundary',
			};
		default:
			return undefined;
	}
}

/**
 * A read-only surface still lets the reader focus the document, select text, move a caret with the
 * keyboard and copy. ProseMirror renders a non-editable document without a caret, so the browser's
 * own selection is followed and mirrored into the editor state: custom controls then see it through
 * `selection` and the formatting snapshot. Nothing here writes to the document. Returns a disposer.
 */
export function attachSurfaceNavigation(core: EditorCore): () => void {
	const canvas = core.shell.canvas;
	if (!canvas) return () => {};
	const doc = canvas.ownerDocument;
	// A read-only ProseMirror surface cannot take focus, so the canvas stands in for it.
	canvas.tabIndex = -1;
	const domSelection = (): globalThis.Selection | null => {
		const root = core.view?.root as
			| { getSelection?: () => globalThis.Selection | null }
			| undefined;
		return root?.getSelection?.() ?? doc.getSelection();
	};
	/** The selection's ends as seen from inside the shadow tree, whichever API the browser has. */
	const ends = () => {
		const selection = domSelection();
		if (!selection) return undefined;
		if (selection.anchorNode && selection.anchorNode === selection.focusNode) {
			const root = core.view?.root;
			if (!(root instanceof ShadowRoot) || root.host !== selection.anchorNode)
				return {
					anchor: selection.anchorNode,
					anchorOffset: selection.anchorOffset,
					focus: selection.focusNode,
					focusOffset: selection.focusOffset,
				};
		}
		const root = core.view?.root;
		const composed = (
			selection as { getComposedRanges?: (options: { shadowRoots: ShadowRoot[] }) => StaticRange[] }
		).getComposedRanges;
		if (root instanceof ShadowRoot && typeof composed === 'function') {
			const range = composed.call(selection, { shadowRoots: [root] })[0];
			if (range)
				return {
					anchor: range.startContainer,
					anchorOffset: range.startOffset,
					focus: range.endContainer,
					focusOffset: range.endOffset,
				};
		}
		return {
			anchor: selection.anchorNode,
			anchorOffset: selection.anchorOffset,
			focus: selection.focusNode,
			focusOffset: selection.focusOffset,
		};
	};
	const mirror = () => {
		const view = core.view;
		if (!view || view.isDestroyed || view.editable) return;
		const found = ends();
		if (!found?.anchor || !found.focus) return;
		if (!view.dom.contains(found.anchor) || !view.dom.contains(found.focus)) return;
		try {
			const $anchor = view.state.doc.resolve(view.posAtDOM(found.anchor, found.anchorOffset));
			const $head = view.state.doc.resolve(view.posAtDOM(found.focus, found.focusOffset));
			const next = TextSelection.between($anchor, $head);
			if (next.eq(view.state.selection)) return;
			view.dispatch(view.state.tr.setSelection(next));
		} catch {
			/* A position inside an atom or a node view the editor cannot address keeps the old selection. */
		}
	};
	doc.addEventListener('selectionchange', mirror);
	const onKeyDown = (event: KeyboardEvent) => {
		const view = core.view;
		if (!view || view.isDestroyed || view.editable || event.defaultPrevented) return;
		if (event.target !== canvas && !view.dom.contains(event.target as Node | null)) return;
		const selection = domSelection();
		if (!selection) return;
		const mod = event.ctrlKey || event.metaKey;
		if (mod && event.key.toLowerCase() === 'a') {
			selection.selectAllChildren(view.dom);
			event.preventDefault();
			return;
		}
		const move = moveFor(event);
		if (!move || event.metaKey || typeof selection.modify !== 'function') return;
		if (!view.dom.contains(ends()?.anchor ?? null)) {
			selection.selectAllChildren(view.dom);
			selection.collapseToStart();
		}
		selection.modify(event.shiftKey ? 'extend' : 'move', move.direction, move.granularity);
		event.preventDefault();
		mirror();
	};
	canvas.addEventListener('keydown', onKeyDown);
	return () => {
		doc.removeEventListener('selectionchange', mirror);
		canvas.removeEventListener('keydown', onKeyDown);
	};
}
