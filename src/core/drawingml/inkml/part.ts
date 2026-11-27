import { buildInkMlContent, parseInkMlContent } from './inkml-content-part';
import type { InkTargetBox } from './inkml-ink-space';
import { MAX_INK_PART_CHARS } from './inkml-limits';
import { inkStrokesEqual } from './equal';
import type { InkStroke, InkXmlObject } from './types';
import { inkObjectFromXml, inkObjectToXml } from './xml-tree';

/**
 * An `application/inkml+xml` part as text, the way Word (`w14:contentPart`) and Excel
 * (`xdr:contentPart`) relate it: read into strokes in the box of the drawing that shows it, and
 * written back in the dialect PowerPoint, Word and Excel 16 all open (see `buildInkMlContent`).
 *
 * @module drawingml/inkml/part
 */

export { INKML_CONTENT_TYPE, INK_RELATIONSHIP_TYPE } from './constants';
export { inkStrokesEqual } from './equal';

export interface InkPart {
	strokes: InkStroke[];
	/** Part of the ink could not be read (see `ParsedInkMlContent.truncated`): keep the part as it is. */
	truncated: boolean;
	/** The parsed part, kept so a rewrite preserves the nodes the model does not hold. */
	raw?: InkXmlObject;
}

/**
 * Reads an ink part into strokes placed in `box` (CSS pixels). A part that is too large or not
 * well-formed reads as no strokes; it is never thrown, since the drawing still keeps its markup.
 */
export function readInkPart(xml: string | undefined, box?: InkTargetBox): InkPart {
	if (!xml) return { strokes: [], truncated: false };
	// Too large or unreadable: nothing is read, and the part must be kept as it is.
	if (xml.length > MAX_INK_PART_CHARS) return { strokes: [], truncated: true };
	let tree: InkXmlObject;
	try {
		tree = inkObjectFromXml(xml);
	} catch {
		return { strokes: [], truncated: true };
	}
	const parsed = parseInkMlContent(tree, box);
	return { strokes: parsed.strokes, truncated: parsed.truncated, raw: parsed.rawXml };
}

/**
 * Writes strokes (drawn in a `space` box, CSS pixels) as an ink part, keeping the unknown nodes of
 * `raw` (a part read before).
 */
export function writeInkPart(
	strokes: readonly InkStroke[],
	raw?: InkXmlObject,
	space?: InkTargetBox,
): string {
	return inkObjectToXml(
		buildInkMlContent(strokes, raw, space && { width: space.width, height: space.height }),
	);
}

/**
 * The text to save for an ink part whose strokes are drawn in `space`: `undefined` when the part
 * already reads as `strokes` (an untouched part keeps its bytes, the applications' own trace
 * encoding included) or was read only in part, otherwise the rewritten part.
 */
export function updatedInkPart(
	xml: string | undefined,
	strokes: readonly InkStroke[],
	space: InkTargetBox,
): string | undefined {
	const current = readInkPart(xml, space);
	// A part read only in part is never rewritten: the strokes not read would be lost.
	if (current.truncated) return undefined;
	if (xml && inkStrokesEqual(current.strokes, strokes)) return undefined;
	return writeInkPart(strokes, current.raw, space);
}
