// A read-only view over one DrawingML element, so a reader is written once and serves both the
// shared DOM (`xml` area) and an object tree (the pptx area's fast-xml-parser nodes). Each format
// supplies a view; the readers (`gradient-read.ts`, `effect-read.ts`, `effect-dag.ts`) never see the node type.
import { parseDrawingColorIn, parseDrawingColorList } from './drawing-color';
import { NS, attributeReader, children, elements, first, type XmlElement } from './dom';
import type { DrawingColor } from './types';

export interface DrawingNodeView<N> {
	/** The attribute by local name; `undefined` when absent. */
	attribute(node: N, name: string): string | undefined;
	/** The first `a:` child with this local name. */
	child(node: N, localName: string): N | undefined;
	/** Every `a:` child with this local name, in document order. */
	children(node: N, localName: string): N[];
	/** The element children of the node in document order, with their local names. */
	elements(node: N): { name: string; node: N }[];
	/** The colour element `node` holds directly (`a:gs`, `a:outerShdw`, `a:solidFill`...). */
	color(node: N): DrawingColor | undefined;
	/** Every colour element `node` holds directly, in document order (`a:duotone`). */
	colors(node: N): DrawingColor[];
}

/** The view over the shared DOM. */
export const DOM_DRAWING_VIEW: DrawingNodeView<XmlElement> = {
	attribute: (node, name) => attributeReader(node)(name),
	child: (node, name) => first(node, name, NS.a),
	children: (node, name) => children(node, name, NS.a),
	elements: (node) => elements(node).map((element) => ({ name: element.localName, node: element })),
	color: (node) => parseDrawingColorIn(node),
	colors: (node) => parseDrawingColorList(node),
};
