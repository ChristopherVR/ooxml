import type { Paragraph } from './model';
import { first, getW, type XmlElement } from './xml';
import { parseTabStops } from './tab-stops';
import { parseJustification } from './paragraph-alignment';
import { onOffElement, parseInteger, parseSignedTwips, parseTwips } from './simple-types';
import { PAGINATION_KEYS, parseOutlineLevel } from './paragraph-styles';
import { parseParagraphBorders, parseShadingFill } from './table-borders';

export type DirectParagraphProperties = Omit<
	Paragraph,
	| 'type'
	| 'id'
	| 'runs'
	| 'markRevision'
	| 'formatRevision'
	| 'bookmarks'
	| 'restoredParagraphPropertiesXml'
	| 'sourceParagraphPropertiesXml'
>;
const signedTwipValue = parseSignedTwips;
const twipValue = parseTwips;

/** Canonical direct properties shared by all stories and formatting restoration. */
export function parseDirectParagraphProperties(
	props: XmlElement | undefined,
): DirectParagraphProperties {
	const paragraph: DirectParagraphProperties = {};
	const bidi = onOffElement(first(props, 'bidi'));
	if (bidi !== undefined) paragraph.direction = bidi ? 'rtl' : 'ltr';
	const { align, justification } = parseJustification(props);
	if (align) paragraph.align = align;
	if (justification) paragraph.justification = justification;
	const style = getW(first(props, 'pStyle'), 'val');
	if (style) paragraph.style = style;
	const pageBreakBefore = onOffElement(first(props, 'pageBreakBefore'));
	if (pageBreakBefore !== undefined) paragraph.pageBreakBefore = pageBreakBefore;
	for (const key of PAGINATION_KEYS) {
		const value = onOffElement(first(props, key));
		if (value !== undefined) paragraph[key] = value;
	}
	const outline = parseOutlineLevel(first(props, 'outlineLvl'));
	if (outline !== undefined) paragraph.outlineLevel = outline;
	const frame = first(props, 'framePr');
	const dropCap = getW(frame, 'dropCap');
	if (dropCap === 'drop' || dropCap === 'margin') {
		const lines = Number(getW(frame, 'lines') ?? 3);
		paragraph.dropCap = {
			style: dropCap,
			lines: Number.isInteger(lines) && lines >= 1 && lines <= 10 ? lines : 3,
			...(twipValue(getW(frame, 'hSpace')) !== undefined
				? { distanceTwips: twipValue(getW(frame, 'hSpace'))! }
				: {}),
		};
	}
	const borders = parseParagraphBorders(first(props, 'pBdr'));
	if (borders) paragraph.borders = borders;
	const shading = parseShadingFill(first(props, 'shd'));
	if (shading) paragraph.shadingFill = shading;
	const tabStops = parseTabStops(first(props, 'tabs'));
	if (tabStops.length) paragraph.tabStops = tabStops;
	const spacing = first(props, 'spacing');
	const before = twipValue(getW(spacing, 'before'));
	const after = twipValue(getW(spacing, 'after'));
	const line = signedTwipValue(getW(spacing, 'line'));
	if (before !== undefined) paragraph.spacingBeforeTwips = before;
	if (after !== undefined) paragraph.spacingAfterTwips = after;
	if (line !== undefined) {
		paragraph.lineSpacingTwips = line;
	}
	const rule = getW(spacing, 'lineRule');
	if (rule === 'auto' || rule === 'exact' || rule === 'atLeast') paragraph.lineSpacingRule = rule;
	else if (line !== undefined) paragraph.lineSpacingRule = 'auto';
	const indent = first(props, 'ind');
	const indentLeft = signedTwipValue(getW(indent, 'left'));
	const indentRight = signedTwipValue(getW(indent, 'right'));
	const indentStart = signedTwipValue(getW(indent, 'start'));
	const indentEnd = signedTwipValue(getW(indent, 'end'));
	const firstLine = twipValue(getW(indent, 'firstLine'));
	const hanging = twipValue(getW(indent, 'hanging'));
	if (indentLeft !== undefined) paragraph.indentLeftTwips = indentLeft;
	if (indentRight !== undefined) paragraph.indentRightTwips = indentRight;
	if (indentStart !== undefined) paragraph.indentStartTwips = indentStart;
	if (indentEnd !== undefined) paragraph.indentEndTwips = indentEnd;
	if (firstLine !== undefined) paragraph.firstLineTwips = firstLine;
	if (hanging !== undefined) paragraph.hangingTwips = hanging;
	const numPr = first(props, 'numPr');
	const numId = parseInteger(getW(first(numPr, 'numId'), 'val'));
	if (numPr && numId !== undefined && numId >= 0)
		paragraph.numbering = {
			numId,
			level: parseInteger(getW(first(numPr, 'ilvl'), 'val')) ?? 0,
		};
	return paragraph;
}
