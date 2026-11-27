// Ink on the editing surface: a node view whose painting (`ink-node-view.ts`) loads with the first
// ink a document shows, so an editor that never shows ink never downloads it. Once loaded it paints
// at once, so a story preview serialized right away (a header, a note) holds its ink; only the
// first ink of a session is painted a moment after its node.
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { NodeView } from 'prosemirror-view';

type Painter = typeof import('./ink-node-view');
let painter: Painter | undefined;
let loading: Promise<Painter> | undefined;

export function inkNodeView(node: ProseMirrorNode): NodeView {
	const dom = document.createElement('span');
	dom.contentEditable = 'false';
	dom.dataset.docxInk = '';
	const paint = (next: ProseMirrorNode) => {
		node = next;
		dom.style.width = `${Number(next.attrs.widthPx) || 1}px`;
		dom.style.height = `${Number(next.attrs.heightPx) || 1}px`;
		if (painter) painter.paintInk(dom, node);
		else
			void (loading ??= import('./ink-node-view').then((module) => (painter = module))).then(
				(module) => module.paintInk(dom, node),
				() => (loading = undefined),
			);
	};
	paint(node);
	return {
		dom,
		update: (next) =>
			next.type === node.type &&
			Boolean(next.attrs.ink || next.attrs.legacyInk) &&
			(paint(next), true),
		selectNode: () => dom.classList.add('dve-ink-selected'),
		deselectNode: () => dom.classList.remove('dve-ink-selected'),
		ignoreMutation: () => true,
	};
}
