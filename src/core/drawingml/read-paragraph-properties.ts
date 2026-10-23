// DrawingML paragraph properties reader (`a:pPr`, `a:defPPr`, `a:lvlNpPr`,
// CT_TextParagraphProperties) over the shared DOM. Semantics follow the PowerPoint reader this
// replaces (`pptx/core/utils/paragraph-properties-parser.ts`): `lvl` clamped to 0-8, both lexical
// forms of percentages, `a:buNone` as an explicit opinion.
import { parseOoxmlPercent } from '../color/color-primitives';
import { parseDrawingColorIn } from './drawing-color';
import {
	NS,
	booleanAttribute,
	children,
	elements,
	first,
	integerAttribute,
	stringAttribute,
} from './dom';
import type { XmlElement } from './dom';
import {
	isDrawingElement,
	outerXml,
	parseDrawingFont,
	parseDrawingRunProperties,
} from './read-run-properties';
import type { DrawingParagraphProperties, DrawingTabStop } from './text-model';
import type { DrawingTextSpacing } from './types';

/** Reads `a:lnSpc`, `a:spcBef` or `a:spcAft`: points (`a:spcPts`) or a fraction (`a:spcPct`). */
export function parseDrawingTextSpacing(
	element: XmlElement | undefined,
): DrawingTextSpacing | undefined {
	const points = integerAttribute(first(element, 'spcPts', NS.a), 'val');
	if (points !== undefined) return { unit: 'points', value: points / 100 };
	const percent = parseOoxmlPercent(first(element, 'spcPct', NS.a)?.getAttribute('val'));
	if (percent !== undefined) return { unit: 'percent', value: percent };
	return undefined;
}

const EMU_ATTRIBUTES = [
	['marginLeftEmu', 'marL'],
	['marginRightEmu', 'marR'],
	['indentEmu', 'indent'],
	['defaultTabSizeEmu', 'defTabSz'],
] as const;

const BOOLEAN_ATTRIBUTES = [
	['rightToLeft', 'rtl'],
	['eastAsianLineBreak', 'eaLnBrk'],
	['latinLineBreak', 'latinLnBrk'],
	['hangingPunctuation', 'hangingPunct'],
] as const;

function readAttributes(element: XmlElement, out: DrawingParagraphProperties): void {
	const level = integerAttribute(element, 'lvl');
	if (level !== undefined) out.level = Math.min(Math.max(level, 0), 8);
	const align = stringAttribute(element, 'algn');
	if (align) out.align = align;
	for (const [key, name] of EMU_ATTRIBUTES) {
		const value = integerAttribute(element, name);
		if (value !== undefined) out[key] = value;
	}
	for (const [key, name] of BOOLEAN_ATTRIBUTES) {
		const value = booleanAttribute(element, name);
		if (value !== undefined) out[key] = value;
	}
	const fontAlign = stringAttribute(element, 'fontAlgn');
	if (fontAlign) out.fontAlign = fontAlign;
}

function readTabStops(element: XmlElement): DrawingTabStop[] {
	return children(element, 'tab', NS.a).map((tab) => {
		const stop: DrawingTabStop = {};
		const position = integerAttribute(tab, 'pos');
		if (position !== undefined) stop.positionEmu = position;
		const align = stringAttribute(tab, 'algn');
		if (align) stop.align = align;
		return stop;
	});
}

/** Reads the bullet children (`a:buClr*`, `a:buSz*`, `a:buFont*`, `a:bu*`); false when not one. */
function readBulletChild(child: XmlElement, out: DrawingParagraphProperties): boolean {
	switch (child.localName) {
		case 'buClrTx':
			out.bulletColor = 'text';
			return true;
		case 'buClr': {
			const color = parseDrawingColorIn(child);
			if (color) out.bulletColor = color;
			return color !== undefined;
		}
		case 'buSzTx':
			out.bulletSize = 'text';
			return true;
		case 'buSzPct': {
			const value = parseOoxmlPercent(child.getAttribute('val'));
			if (value !== undefined) out.bulletSize = { unit: 'percent', value };
			return value !== undefined;
		}
		case 'buSzPts': {
			const value = integerAttribute(child, 'val');
			if (value !== undefined) out.bulletSize = { unit: 'points', value: value / 100 };
			return value !== undefined;
		}
		case 'buFontTx':
			out.bulletFont = 'text';
			return true;
		case 'buFont': {
			const font = parseDrawingFont(child);
			if (font) out.bulletFont = font;
			return font !== undefined;
		}
		case 'buNone':
			out.bullet = { kind: 'none' };
			return true;
		case 'buChar': {
			const char = child.getAttribute('char');
			if (char) out.bullet = { kind: 'char', char };
			return Boolean(char);
		}
		case 'buAutoNum': {
			const scheme = stringAttribute(child, 'type');
			if (!scheme) return false;
			const startAt = integerAttribute(child, 'startAt');
			out.bullet = { kind: 'autoNumber', scheme, ...(startAt !== undefined ? { startAt } : {}) };
			return true;
		}
		case 'buBlip':
			out.bullet = { kind: 'picture', xml: outerXml(child) };
			return true;
		default:
			return false;
	}
}

/** Reads paragraph properties; `undefined` for a missing element. */
export function parseDrawingParagraphProperties(
	element: XmlElement | undefined,
): DrawingParagraphProperties | undefined {
	if (!element) return undefined;
	const out: DrawingParagraphProperties = {};
	readAttributes(element, out);
	const extra: string[] = [];
	for (const child of elements(element)) {
		if (!isDrawingElement(child)) {
			extra.push(outerXml(child));
			continue;
		}
		const name = child.localName;
		if (name === 'lnSpc' || name === 'spcBef' || name === 'spcAft') {
			const spacing = parseDrawingTextSpacing(child);
			const key =
				name === 'lnSpc' ? 'lineSpacing' : name === 'spcBef' ? 'spaceBefore' : 'spaceAfter';
			if (spacing) out[key] = spacing;
			else extra.push(outerXml(child));
		} else if (name === 'tabLst') out.tabStops = readTabStops(child);
		else if (name === 'defRPr') {
			const run = parseDrawingRunProperties(child);
			if (run) out.defaultRunProperties = run;
		} else if (!readBulletChild(child, out)) extra.push(outerXml(child));
	}
	if (extra.length) out.extraXml = extra;
	return out;
}
