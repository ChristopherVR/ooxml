// DrawingML paragraph properties reader (`a:pPr`, `a:defPPr`, `a:lvlNpPr`,
// CT_TextParagraphProperties) over the shared DOM. The attribute, spacing, bullet and tab-stop
// rules are the shared codec's (`text-paragraph-codec.ts`, which PowerPoint's object-tree reader
// also runs): `lvl` clamped to 0-8, both lexical forms of percentages, `a:buNone` as an explicit
// opinion. This reader adds the bullet colour and picture and keeps unknown children verbatim.
import { parseDrawingColorIn } from './drawing-color';
import { elements, type XmlElement } from './dom';
import { isDrawingElement, outerXml, parseDrawingRunProperties } from './read-run-properties';
import { domTextElementView } from './text-element-view';
import type { DrawingParagraphProperties } from './text-model';
import {
	readDrawingBulletChild,
	readDrawingParagraphAttributes,
	readDrawingTabStop,
	readDrawingTextSpacing,
} from './text-paragraph-codec';
import type { DrawingTextSpacing } from './types';

/** Reads `a:lnSpc`, `a:spcBef` or `a:spcAft`: points (`a:spcPts`) or a fraction (`a:spcPct`). */
export function parseDrawingTextSpacing(
	element: XmlElement | undefined,
): DrawingTextSpacing | undefined {
	return readDrawingTextSpacing(element ? domTextElementView(element) : undefined);
}

/** Reads the bullet children (`a:buClr*`, `a:buSz*`, `a:buFont*`, `a:bu*`); false when not one. */
function readBulletChild(child: XmlElement, out: DrawingParagraphProperties): boolean {
	if (child.localName === 'buClr') {
		const color = parseDrawingColorIn(child);
		if (color) out.bulletColor = color;
		return color !== undefined;
	}
	if (child.localName === 'buBlip') {
		out.bullet = { kind: 'picture', xml: outerXml(child) };
		return true;
	}
	return readDrawingBulletChild(child.localName, domTextElementView(child), out);
}

const SPACING_KEYS = {
	lnSpc: 'lineSpacing',
	spcBef: 'spaceBefore',
	spcAft: 'spaceAfter',
} as const;

/** Reads paragraph properties; `undefined` for a missing element. */
export function parseDrawingParagraphProperties(
	element: XmlElement | undefined,
): DrawingParagraphProperties | undefined {
	if (!element) return undefined;
	const out: DrawingParagraphProperties = {};
	readDrawingParagraphAttributes(domTextElementView(element), out);
	const extra: string[] = [];
	for (const child of elements(element)) {
		if (!isDrawingElement(child)) {
			extra.push(outerXml(child));
			continue;
		}
		const name = child.localName;
		if (name === 'lnSpc' || name === 'spcBef' || name === 'spcAft') {
			const spacing = parseDrawingTextSpacing(child);
			if (spacing) out[SPACING_KEYS[name]] = spacing;
			else extra.push(outerXml(child));
		} else if (name === 'tabLst')
			out.tabStops = domTextElementView(child).children('tab').map(readDrawingTabStop);
		else if (name === 'defRPr') {
			const run = parseDrawingRunProperties(child);
			if (run) out.defaultRunProperties = run;
		} else if (!readBulletChild(child, out)) extra.push(outerXml(child));
	}
	if (extra.length) out.extraXml = extra;
	return out;
}
