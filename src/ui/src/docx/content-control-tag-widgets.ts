import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { controlOccurrences, everyControlSpan } from 'ooxml-core/docx/ui';

type Edge = 'start' | 'end';

/** Descend from a block, row or cell boundary to the text position inside its paragraph. */
function paragraphEdge(doc: ProseMirrorNode, position: number, edge: Edge): number | undefined {
	let node = edge === 'start' ? doc.nodeAt(position) : doc.resolve(position).nodeBefore;
	if (!node) return undefined;
	let start = edge === 'start' ? position : position - node.nodeSize;
	while (!node.isTextblock) {
		const child: ProseMirrorNode | null = edge === 'start' ? node.firstChild : node.lastChild;
		if (!child) return undefined;
		start += 1 + (edge === 'start' ? 0 : node.content.size - child.nodeSize);
		node = child;
	}
	return edge === 'start' ? start + 1 : start + node.nodeSize - 1;
}

function tagWidget(position: number, edge: Edge, title: string, side: number): Decoration {
	return Decoration.widget(
		position,
		(view) => {
			const tag = view.dom.ownerDocument.createElement('span');
			tag.className = 'dve-sdt-tag';
			tag.dataset.tagEdge = edge;
			tag.dataset.tagTitle = title;
			tag.setAttribute('aria-hidden', 'true');
			return tag;
		},
		{ side, key: JSON.stringify([edge, title, side]), ignoreSelection: true },
	);
}

export function contentControlTagWidgets(doc: ProseMirrorNode): DecorationSet {
	const spans = everyControlSpan(doc);
	const tagged = spans.filter((span) => span.control.appearance === 'tags');
	if (!tagged.length) return DecorationSet.empty;
	const inline = spans.filter((span) => span.level === 'inline');
	const widgets: Decoration[] = [];
	for (const { spans: parts, from, to } of controlOccurrences(doc, tagged, inline)) {
		const first = parts[0];
		if (!first) continue;
		const isInline = first.level === 'inline';
		const start = isInline ? from : paragraphEdge(doc, from, 'start');
		const end = isInline ? to : paragraphEdge(doc, to, 'end');
		if (start === undefined || end === undefined) continue;
		const title = (first.control.alias ?? '').replace(/\s+/g, ' ').trim();
		// Adjacent controls close before the next opens. At a shared paragraph edge, the
		// outer control opens first and closes last. An empty paragraph keeps both in order.
		const openSide = 1 / (start - from + 1);
		const closeSide = (start === end ? 2 : 0) - 1 / (to - end + 1);
		widgets.push(tagWidget(start, 'start', title, openSide), tagWidget(end, 'end', '', closeSide));
	}
	return DecorationSet.create(doc, widgets);
}
