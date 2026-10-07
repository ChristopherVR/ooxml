import type { Attrs } from 'prosemirror-model';
import type { Paragraph, TextRun } from '../model';
import { paragraphTwipsFromAttrs } from '../attr-units';

/** Paragraph pagination toggles carried through the editor as node attributes. */
export const PARAGRAPH_KEEP_KEYS = [
	'keepNext',
	'keepLines',
	'widowControl',
	'contextualSpacing',
	'suppressLineNumbers',
] as const;

/** Shared conversion for caller-provided paragraph schemas and inline run adapters. */
export function paragraphFromAttrs(attrs: Attrs, id: string, runs: TextRun[]): Paragraph {
	return {
		type: 'paragraph',
		id,
		runs,
		...(attrs.restoredParagraphPropertiesXml
			? { restoredParagraphPropertiesXml: String(attrs.restoredParagraphPropertiesXml) }
			: {}),
		...(attrs.sourceParagraphPropertiesXml
			? { sourceParagraphPropertiesXml: String(attrs.sourceParagraphPropertiesXml) }
			: {}),
		...(attrs.markRevision ? { markRevision: structuredClone(attrs.markRevision) } : {}),
		...(attrs.formatRevision ? { formatRevision: structuredClone(attrs.formatRevision) } : {}),
		...(attrs.align != null ? { align: attrs.align } : {}),
		...(attrs.justification != null ? { justification: attrs.justification } : {}),
		...(attrs.outlineLevel != null ? { outlineLevel: Number(attrs.outlineLevel) } : {}),
		...(attrs.direction != null ? { direction: attrs.direction } : {}),
		...(attrs.style ? { style: attrs.style } : {}),
		...paragraphTwipsFromAttrs(attrs),
		...(attrs.lineSpacingRule != null ? { lineSpacingRule: attrs.lineSpacingRule } : {}),
		...(attrs.numId != null ? { numbering: { numId: attrs.numId, level: attrs.ilvl ?? 0 } } : {}),
		...(attrs.pageBreakBefore != null ? { pageBreakBefore: Boolean(attrs.pageBreakBefore) } : {}),
		...(attrs.tabStops?.length ? { tabStops: structuredClone(attrs.tabStops) } : {}),
		...Object.fromEntries(
			PARAGRAPH_KEEP_KEYS.filter((key) => attrs[key] != null).map((key) => [key, attrs[key]]),
		),
		...(attrs.dropCap ? { dropCap: { ...attrs.dropCap } } : {}),
		...(attrs.borders ? { borders: structuredClone(attrs.borders) } : {}),
		...(attrs.shadingFill ? { shadingFill: attrs.shadingFill } : {}),
		...(attrs.bookmarks?.length ? { bookmarks: [...attrs.bookmarks] } : {}),
	};
}

/** Model attributes shared by editor construction and formatting snapshots. */
export function paragraphAttrs(paragraph: Paragraph): Attrs {
	return {
		id: paragraph.id,
		markRevision: paragraph.markRevision ?? null,
		formatRevision: paragraph.formatRevision ?? null,
		restoredParagraphPropertiesXml: paragraph.restoredParagraphPropertiesXml ?? null,
		sourceParagraphPropertiesXml: paragraph.sourceParagraphPropertiesXml ?? null,
		align: paragraph.align ?? null,
		justification: paragraph.justification ?? null,
		outlineLevel: paragraph.outlineLevel ?? null,
		direction: paragraph.direction ?? null,
		style: paragraph.style || '',
		spacingBeforeTwips: paragraph.spacingBeforeTwips ?? null,
		spacingAfterTwips: paragraph.spacingAfterTwips ?? null,
		lineSpacingTwips: paragraph.lineSpacingTwips ?? null,
		lineSpacingRule: paragraph.lineSpacingRule ?? null,
		indentLeftTwips: paragraph.indentLeftTwips ?? null,
		indentRightTwips: paragraph.indentRightTwips ?? null,
		indentStartTwips: paragraph.indentStartTwips ?? null,
		indentEndTwips: paragraph.indentEndTwips ?? null,
		firstLineTwips: paragraph.firstLineTwips ?? null,
		hangingTwips: paragraph.hangingTwips ?? null,
		numId: paragraph.numbering?.numId ?? null,
		ilvl: paragraph.numbering ? paragraph.numbering.level : null,
		pageBreakBefore: paragraph.pageBreakBefore ?? null,
		tabStops: paragraph.tabStops?.length ? paragraph.tabStops : null,
		keepNext: paragraph.keepNext ?? null,
		keepLines: paragraph.keepLines ?? null,
		widowControl: paragraph.widowControl ?? null,
		contextualSpacing: paragraph.contextualSpacing ?? null,
		suppressLineNumbers: paragraph.suppressLineNumbers ?? null,
		dropCap: paragraph.dropCap ?? null,
		borders: paragraph.borders ?? null,
		shadingFill: paragraph.shadingFill ?? null,
		bookmarks: paragraph.bookmarks ?? [],
	};
}
