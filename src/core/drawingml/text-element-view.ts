// A read-only view of one DrawingML text element (`a:pPr`, `a:lvl1pPr`, `a:tab`...), so the
// paragraph property codec (`text-paragraph-codec.ts`) reads the same rules over the shared DOM and
// over a product's own parsed tree (PowerPoint's object tree adapts to it in pptx). Child lookups
// are by local name in the DrawingML (`a:`) namespace.
import { NS, attributeReader, children, first, type XmlElement } from './dom';

/** One element: its attributes by local name and its `a:` children by local name. */
export interface DrawingTextElementView {
	/** The raw attribute value, `undefined` when absent. */
	attribute(name: string): string | undefined;
	/** The first `a:<local>` child. */
	child(local: string): DrawingTextElementView | undefined;
	/** Every `a:<local>` child, in document order. */
	children(local: string): DrawingTextElementView[];
}

/** The view of a DOM element. */
export function domTextElementView(element: XmlElement): DrawingTextElementView {
	const read = attributeReader(element);
	return {
		attribute: read,
		child: (local) => {
			const found = first(element, local, NS.a);
			return found ? domTextElementView(found) : undefined;
		},
		children: (local) => children(element, local, NS.a).map(domTextElementView),
	};
}
