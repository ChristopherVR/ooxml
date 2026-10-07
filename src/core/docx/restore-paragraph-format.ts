import type { Paragraph } from './model';
import { parseDirectParagraphProperties } from './paragraph-properties';
import { parsePropertiesSnapshot } from './revision-properties';
import { buildXml, first } from './xml';
import { orderParagraphProperties } from './tab-stops';

export const PARAGRAPH_FORMAT_KEYS = [
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
	const properties = parsePropertiesSnapshot(xml, 'pPr');
	// Paragraph-mark and section properties have independent revision histories. Native Word
	// retains their current values when a paragraph-format snapshot omits them.
	if (paragraph.sourceParagraphPropertiesXml) {
		const current = parsePropertiesSnapshot(paragraph.sourceParagraphPropertiesXml, 'pPr');
		for (const name of ['rPr', 'sectPr']) {
			const element = first(current, name);
			if (element && !first(properties, name))
				properties.appendChild(properties.ownerDocument!.importNode(element, true));
		}
		orderParagraphProperties(properties);
	}
	const previous = parseDirectParagraphProperties(properties);
	for (const key of PARAGRAPH_FORMAT_KEYS) delete paragraph[key];
	Object.assign(paragraph, previous);
	paragraph.restoredParagraphPropertiesXml = buildXml(properties);
	if (paragraph.sourceParagraphPropertiesXml)
		paragraph.sourceParagraphPropertiesXml = paragraph.restoredParagraphPropertiesXml;
	delete paragraph.formatRevision;
}
