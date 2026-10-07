import type { Paragraph } from './model.js';
import { parseDirectParagraphProperties } from './paragraph-properties.js';
import { parsePropertiesSnapshot } from './revision-properties.js';

const FORMAT_KEYS = [
	'align',
	'justification',
	'direction',
	'style',
	'spacingBeforeTwips',
	'spacingAfterTwips',
	'lineSpacingTwips',
	'lineSpacingRule',
	'indentLeftTwips',
	'indentRightTwips',
	'indentStartTwips',
	'indentEndTwips',
	'firstLineTwips',
	'hangingTwips',
	'numbering',
	'pageBreakBefore',
	'tabStops',
	'keepNext',
	'keepLines',
	'widowControl',
	'contextualSpacing',
	'suppressLineNumbers',
	'outlineLevel',
	'dropCap',
	'borders',
	'shadingFill',
] as const satisfies readonly (keyof Paragraph)[];

/** Restore formatting without changing text, stable identity, bookmarks or paragraph mark revisions. */
export function restoreParagraphFormatting(paragraph: Paragraph): void {
	const xml = paragraph.formatRevision?.previousParagraphPropertiesXml;
	if (!xml)
		throw new Error(
			'Cannot reject a paragraph formatting revision without its prior properties snapshot.',
		);
	const previous = parseDirectParagraphProperties(parsePropertiesSnapshot(xml, 'pPr'));
	for (const key of FORMAT_KEYS) delete paragraph[key];
	Object.assign(paragraph, previous);
	paragraph.restoredParagraphPropertiesXml = xml;
	delete paragraph.formatRevision;
}
