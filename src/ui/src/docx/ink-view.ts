// Ink on the editing surface: a node view whose painting (`ink-node-view.ts`) loads with the first
// ink a document shows, so an editor that never shows ink never downloads it.
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { NodeView } from 'prosemirror-view';

let painter: Promise<typeof import('./ink-node-view')> | undefined;

export function inkNodeView(node: ProseMirrorNode): NodeView {
	const dom = document.createElement('span');
	dom.contentEditable = 'false';
	dom.dataset.docxInk = '';
	const paint = (next: ProseMirrorNode) => {
		node = next;
		dom.style.width = `${Number(next.attrs.widthPx) || 1}px`;
		dom.style.height = `${Number(next.attrs.heightPx) || 1}px`;
		void (painter ??= import('./ink-node-view')).then((module) => module.paintInk(dom, node));
	};
	paint(node);
	return {
		dom,
		update: (next) => next.type === node.type && Boolean(next.attrs.ink) && (paint(next), true),
		selectNode: () => dom.classList.add('dve-ink-selected'),
		deselectNode: () => dom.classList.remove('dve-ink-selected'),
		ignoreMutation: () => true,
	};
}
