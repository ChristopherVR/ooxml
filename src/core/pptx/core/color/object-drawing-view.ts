/**
 * The pptx object-tree view of a DrawingML element (`DrawingNodeView` over fast-xml-parser
 * nodes), so the shared `drawingml` readers (gradients, effect lists, effect DAGs) read the
 * runtime's object tree exactly as they read the DOM. See docs/agnostic-core-plan.md, step 3.
 *
 * Children match by local name whatever the prefix, as `drawingChild` does; same-name siblings
 * arrive as one array, so `elements` lists them together at the first one's place (the object
 * tree does not keep the interleaving of differently named siblings).
 *
 * @module object-drawing-view
 */
import type { DrawingNodeView } from '../../../drawingml/node-view';
import type { DrawingColor } from '../../../drawingml/types';
import type { XmlObject } from '../types';
import { CODEC_COLOR_CHOICE_ORDER, objectDrawingColor } from './color-transforms';

const isNode = (value: unknown): value is XmlObject =>
	Boolean(value) && typeof value === 'object' && !Array.isArray(value);

/** An empty element (`<a:grayscl/>`) is `''` in the object tree; it is read as an empty node. */
const asNode = (value: unknown): XmlObject | undefined =>
	isNode(value) ? value : value === '' ? {} : undefined;

const localName = (key: string) => key.slice(key.indexOf(':') + 1);
const isChildKey = (key: string) => !key.startsWith('@_') && key !== '#text';

/**
 * The named children that hold an object, as `drawingChild` and `drawingChildren` read them: an
 * empty element (`<a:tileRect/>`) is not returned here, so the codecs keep reading it as absent
 * (and keep writing it back verbatim); `elements` lists it, for the effect containers' empty
 * leaves (`<a:grayscl/>`).
 */
function childEntries(node: XmlObject, name: string): XmlObject[] {
	for (const [key, value] of Object.entries(node)) {
		if (!isChildKey(key) || localName(key) !== name) continue;
		return (Array.isArray(value) ? value : [value]).filter(isNode);
	}
	return [];
}

/** The colour element a node holds directly, in the codec's colour-choice order. */
export function objectNodeColor(node: XmlObject): DrawingColor | undefined {
	for (const [element, kind] of CODEC_COLOR_CHOICE_ORDER) {
		const value = node[element];
		const colorNode = Array.isArray(value) ? asNode(value[0]) : asNode(value);
		if (colorNode) return objectDrawingColor(kind, colorNode);
	}
	return undefined;
}

export const OBJECT_DRAWING_VIEW: DrawingNodeView<XmlObject> = {
	attribute(node, name) {
		const value = node[`@_${name}`];
		return value === undefined || value === null ? undefined : String(value);
	},
	child: (node, name) => childEntries(node, name)[0],
	children: childEntries,
	elements(node) {
		const found: { name: string; node: XmlObject }[] = [];
		for (const [key, value] of Object.entries(node)) {
			if (!isChildKey(key)) continue;
			for (const entry of Array.isArray(value) ? value : [value]) {
				const child = asNode(entry);
				if (child) found.push({ name: localName(key), node: child });
			}
		}
		return found;
	},
	color: objectNodeColor,
};
