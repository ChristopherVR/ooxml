/**
 * The pptx object-tree adapter of the shared DrawingML paragraph property codec
 * (`drawingml/text-paragraph-codec.ts`, docs/agnostic-core-plan.md step 3).
 *
 * The rules for the paragraph attributes, spacing, the bullet group (except its colour and
 * picture) and tab stops are the codec's, the same ones the DOM reader and writer run. This module
 * only presents a fast-xml-parser node as a {@link DrawingTextElementView} and turns the codec's
 * attribute entries back into object-tree nodes; the pptx readers and writers then convert units
 * (EMU and points to px) on the neutral {@link DrawingParagraphProperties}.
 *
 * @module paragraph-properties-neutral
 */
import type { DrawingTextElementView } from '../../../drawingml/text-element-view';
import type { DrawingParagraphProperties } from '../../../drawingml/text-model';
import {
	drawingTextSpacingChild,
	readDrawingParagraphAttributes,
	readDrawingParagraphCodecProperties,
	type DrawingAttributeEntry,
	type DrawingChildEntry,
} from '../../../drawingml/text-paragraph-codec';
import type { DrawingTextSpacing } from '../../../drawingml/types';
import type { XmlObject } from '../types';

/** An object-tree element: a bare element parses to `''`, so anything not an object is empty. */
function asNode(value: unknown): XmlObject {
	return value && typeof value === 'object' ? (value as XmlObject) : {};
}

/** The object-tree node of a DrawingML element as a codec view; `undefined` when absent. */
export function objectTextElementView(node: unknown): DrawingTextElementView | undefined {
	if (node === undefined || node === null) return undefined;
	const object = asNode(Array.isArray(node) ? node[0] : node);
	return {
		attribute: (name) => {
			const value = object[`@_${name}`];
			return value === undefined || value === null ? undefined : String(value);
		},
		child: (local) => objectTextElementView(object[`a:${local}`]),
		children: (local) => {
			const value = object[`a:${local}`];
			if (value === undefined || value === null) return [];
			return (Array.isArray(value) ? value : [value]).map(
				(entry) => objectTextElementView(entry) as DrawingTextElementView,
			);
		},
	};
}

/**
 * The codec families of an `a:pPr` / `a:lvlNpPr` object node on the neutral model: attributes,
 * spacing, bullet size, font and type, `a:buClrTx`, and tab stops. Empty for a missing node.
 */
export function objectParagraphProperties(pPr: unknown): DrawingParagraphProperties {
	const view = objectTextElementView(pPr);
	return view ? readDrawingParagraphCodecProperties(view) : {};
}

/** Only the attributes of an `a:pPr` / `a:lvlNpPr` object node, on the neutral model. */
export function objectParagraphAttributes(pPr: unknown): DrawingParagraphProperties {
	const view = objectTextElementView(pPr);
	const out: DrawingParagraphProperties = {};
	if (view) readDrawingParagraphAttributes(view, out);
	return out;
}

/** Attribute entries as object-tree attributes (`@_name`). */
export function objectAttributes(entries: readonly DrawingAttributeEntry[]): XmlObject {
	const node: XmlObject = {};
	for (const [name, value] of entries) node[`@_${name}`] = value;
	return node;
}

/** Assigns a codec child entry onto `parent` as `a:<local>`. */
export function assignObjectChild(parent: XmlObject, entry: DrawingChildEntry): void {
	parent[`a:${entry.local}`] = objectAttributes(entry.attributes);
}

/** The inner node of `a:lnSpc` / `a:spcBef` / `a:spcAft` (`{ 'a:spcPts': { '@_val' } }`). */
export function objectTextSpacing(spacing: DrawingTextSpacing): XmlObject {
	const node: XmlObject = {};
	assignObjectChild(node, drawingTextSpacingChild(spacing));
	return node;
}
