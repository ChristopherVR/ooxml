/**
 * The inspector's paragraph and character spacing fields (PowerPoint's
 * Paragraph dialog: indentation, spacing, line spacing and direction; the Font
 * dialog's character spacing), as field descriptors every binding draws with
 * its one generic renderer, beside the Text Box section.
 *
 * Like the Text Box descriptors these read the element's `TextStyle` and write
 * a `Partial<TextStyle>` the binding merges with `textStylePatch` (or hands to
 * its text-style update path). Lengths are shown in points, as PowerPoint shows
 * them: the model keeps paragraph spacing, margins and indents in px and
 * character spacing in hundredths of a point (`a:rPr/@spc`).
 *
 * @module render/inspector/paragraph-fields
 */
import type { TextStyle } from 'ooxml-core/pptx';

import type { InspectorSectionDescriptor } from './inspector-field-model';

export type ParagraphFieldsTarget = TextStyle | undefined;
export type ParagraphFieldsSection = InspectorSectionDescriptor<
	ParagraphFieldsTarget,
	Partial<TextStyle>
>;

/** Points per CSS pixel (72 pt and 96 px to the inch). */
const PT_PER_PX = 0.75;
/** PowerPoint's largest spacing and indent entry, in points. */
const MAX_SPACING_PT = 1584;

const pt = (px: number | undefined): number => (px ?? 0) * PT_PER_PX;
const px = (points: number): number => points / PT_PER_PX;

type PxKey = 'paragraphSpacingBefore' | 'paragraphSpacingAfter' | 'paragraphMarginLeft';

function pointsField(
	id: string,
	key: PxKey | 'paragraphIndent',
	labelKey: string,
	min: number,
): ParagraphFieldsSection['fields'][number] {
	return {
		id,
		kind: 'length',
		labelKey,
		min,
		max: MAX_SPACING_PT,
		step: 1,
		unit: 'pt',
		decimals: 1,
		testId: `paragraph-${id}`,
		read: (style) => pt(style?.[key]),
		write: (_style, value) => ({ [key]: px(Number(value)) }),
	};
}

/** Indentation, spacing, line spacing and direction (untitled: the Paragraph pane heads it). */
export const PARAGRAPH_SPACING_SECTION: ParagraphFieldsSection = {
	id: 'paragraphSpacing',
	fields: [
		pointsField('marginLeft', 'paragraphMarginLeft', 'pptx.textAdvanced.marginLeft', 0),
		pointsField('indent', 'paragraphIndent', 'pptx.textAdvanced.indent', -MAX_SPACING_PT),
		pointsField('spaceBefore', 'paragraphSpacingBefore', 'pptx.textAdvanced.spaceBefore', 0),
		pointsField('spaceAfter', 'paragraphSpacingAfter', 'pptx.textAdvanced.spaceAfter', 0),
		{
			id: 'lineSpacing',
			kind: 'number',
			labelKey: 'pptx.textAdvanced.lineSpacing',
			min: 0.1,
			max: 9.99,
			step: 0.1,
			testId: 'paragraph-lineSpacing',
			read: (style) => style?.lineSpacing ?? 1,
			// A multiple replaces an exact line spacing (`a:spcPct` or `a:spcPts`).
			write: (_style, value) => ({ lineSpacing: Number(value), lineSpacingExactPt: undefined }),
		},
		{
			id: 'lineSpacingExact',
			kind: 'length',
			labelKey: 'pptx.textAdvanced.lineSpacingExact',
			min: 0,
			max: MAX_SPACING_PT,
			step: 1,
			unit: 'pt',
			decimals: 1,
			testId: 'paragraph-lineSpacingExact',
			read: (style) => style?.lineSpacingExactPt ?? 0,
			// Zero returns the paragraph to its multiple.
			write: (_style, value) =>
				Number(value) > 0
					? { lineSpacingExactPt: Number(value), lineSpacing: undefined }
					: { lineSpacingExactPt: undefined },
		},
		{
			id: 'rtl',
			kind: 'toggle',
			labelKey: 'pptx.textAdvanced.rtl',
			testId: 'paragraph-rtl',
			read: (style) => style?.rtl === true,
			write: (_style, value) => ({ rtl: value === true }),
		},
	],
};

/** Character spacing (untitled: the Font pane heads it). */
export const CHARACTER_SPACING_SECTION: ParagraphFieldsSection = {
	id: 'characterSpacing',
	fields: [
		{
			id: 'characterSpacing',
			kind: 'length',
			labelKey: 'pptx.textAdvanced.characterSpacing',
			min: -MAX_SPACING_PT,
			max: MAX_SPACING_PT,
			step: 0.1,
			unit: 'pt',
			decimals: 1,
			wide: true,
			testId: 'font-characterSpacing',
			read: (style) => (style?.characterSpacing ?? 0) / 100,
			write: (_style, value) => ({ characterSpacing: Math.round(Number(value) * 100) }),
		},
	],
};

/** The paragraph fields as a section list, what a generic renderer takes. */
export const PARAGRAPH_SPACING_SECTIONS: readonly ParagraphFieldsSection[] = [
	PARAGRAPH_SPACING_SECTION,
];
/** The character spacing field as a section list. */
export const CHARACTER_SPACING_SECTIONS: readonly ParagraphFieldsSection[] = [
	CHARACTER_SPACING_SECTION,
];
