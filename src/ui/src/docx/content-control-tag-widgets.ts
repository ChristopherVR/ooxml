import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { controlOccurrences, everyControlSpan } from 'ooxml-core/docx/ui';
import { contentControlTagMetrics, cssFontStack } from 'ooxml-core/docx/layout';
import { tagBoxStyle, tagEnds } from './content-control-tag-shape';

export type TagEdge = 'start' | 'end';

/** Descend from a block, row or cell boundary to the text position inside its paragraph. */
function paragraphEdge(doc: ProseMirrorNode, position: number, edge: TagEdge): number | undefined {
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

/**
 * The paragraph settings a tag's box depends on: exact line spacing fixes the line's height, so
 * a tag there must not stretch it, and a right-to-left paragraph mirrors the tag's outline.
 */
function tagParagraph(doc: ProseMirrorNode, position: number): { exact: boolean; rtl: boolean } {
	const { attrs } = doc.resolve(position).parent;
	return {
		exact: attrs.lineSpacingRule === 'exact' && Number.isSafeInteger(attrs.lineSpacingTwips),
		rtl: attrs.direction === 'rtl',
	};
}

function tagWidget(
	doc: ProseMirrorNode,
	position: number,
	edge: TagEdge,
	title: string,
	side: number,
	inline: boolean,
): Decoration {
	const { exact, rtl } = tagParagraph(doc, position);
	return Decoration.widget(
		position,
		(view) => {
			const tag = view.dom.ownerDocument.createElement('span');
			tag.className = 'dve-sdt-tag';
			tag.dataset.tagEdge = edge;
			tag.dataset.tagTitle = title;
			if (inline) tag.dataset.inlineTag = '';
			if (rtl) tag.dataset.tagDirection = 'rtl';
			tag.setAttribute('aria-hidden', 'true');
			const metrics = contentControlTagMetrics;
			Object.assign(tag.style, {
				fontFamily: cssFontStack(metrics.fontFamily),
				fontSize: `${metrics.fontSizePt}pt`,
				...tagBoxStyle(edge, inline, exact),
			});
			tag.append(...tagEnds(tag.ownerDocument, edge, inline));
			return tag;
		},
		{
			side,
			edge,
			title,
			inline,
			marks: [],
			key: JSON.stringify([edge, title, side, inline, exact, rtl]),
			ignoreSelection: true,
		},
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
		// Structural tags bracket the paragraph's caret. At a shared edge, the outer
		// control opens first and closes last. Inline tags default to the outside caret.
		const openSide = isInline ? 0.5 : -2 + 1 / (start - from + 1);
		const closeSide = isInline ? -0.5 : 2 - 1 / (to - end + 1);
		widgets.push(
			tagWidget(doc, start, 'start', title, openSide, isInline),
			tagWidget(doc, end, 'end', isInline ? '' : title, closeSide, isInline),
		);
	}
	return DecorationSet.create(doc, widgets);
}

/** Inline edges can share a document position while keeping distinct DOM sides. */
export function inlineTagsAt(widgets: DecorationSet, pos: number): Decoration[] {
	return widgets.find(pos, pos, (spec) => spec.inline === true);
}

export function widgetsWithCaretInside(
	doc: ProseMirrorNode,
	widgets: DecorationSet,
	pos: number,
	edge: TagEdge,
): DecorationSet {
	const widget = inlineTagsAt(widgets, pos).find(({ spec }) => spec.edge === edge);
	if (!widget) return widgets;
	// Keep an adjacent control's closing tag before the next one's opening tag.
	const side = edge === 'start' ? -0.25 : 0.25;
	return widgets
		.remove([widget])
		.add(doc, [tagWidget(doc, pos, edge, widget.spec.title, side, true)]);
}
