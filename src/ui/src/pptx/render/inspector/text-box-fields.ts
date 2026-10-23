/**
 * The Text Box options of the inspector (Format Shape > Text Options > Text
 * Box): vertical alignment, text direction, autofit, wrap and columns, as one
 * untitled section of field descriptors every binding draws with its one
 * generic renderer.
 *
 * Every field is an `a:bodyPr` property (`BODY_TEXT_STYLE_KEYS` in the core),
 * so the descriptors read the element's `TextStyle` and produce a body-level
 * `Partial<TextStyle>`: merge it with `textStylePatch` (bindings that patch
 * elements) or hand it to a text-style update path as is. The value lists are
 * the shared `text-advanced.ts` catalogues.
 *
 * @module render/inspector/text-box-fields
 */
import type { TextStyle } from 'ooxml-core/pptx';

import { AUTOFIT_MODE_OPTIONS, TEXT_DIRECTION_OPTIONS, VALIGN_OPTIONS } from '../text-advanced';
import type { InspectorFieldOption, InspectorSectionDescriptor } from './inspector-field-model';

/** What the Text Box descriptors edit: the element's text style (absent for none). */
export type TextBoxTarget = TextStyle | undefined;
export type TextBoxSection = InspectorSectionDescriptor<TextBoxTarget, Partial<TextStyle>>;

/** Most columns a text box offers (`a:bodyPr@numCol` allows 16). */
const MAX_COLUMNS = 16;

/** The anchors the pane offers (`VALIGN_OPTIONS`); `justified`/`distributed` are not offered. */
const VALIGN_LABEL_KEYS: Partial<Record<NonNullable<TextStyle['vAlign']>, string>> = {
	top: 'pptx.textPanel.valignTop',
	middle: 'pptx.textPanel.valignMiddle',
	bottom: 'pptx.textPanel.valignBottom',
};

const AUTOFIT_LABEL_KEYS: Record<NonNullable<TextStyle['autoFitMode']>, string> = {
	none: 'pptx.textAdvanced.autoFitNone',
	normal: 'pptx.textAdvanced.autoFitShrink',
	shrink: 'pptx.textAdvanced.autoFitResize',
};

const VALIGN_CHOICES: readonly InspectorFieldOption[] = VALIGN_OPTIONS.flatMap(([value]) => {
	const labelKey = VALIGN_LABEL_KEYS[value];
	return labelKey ? [{ value, labelKey }] : [];
});
const DIRECTION_CHOICES: readonly InspectorFieldOption[] = TEXT_DIRECTION_OPTIONS.map(
	([value]) => ({ value, labelKey: `pptx.textAdvanced.direction.${value}` }),
);
const AUTOFIT_CHOICES: readonly InspectorFieldOption[] = AUTOFIT_MODE_OPTIONS.map(([value]) => ({
	value,
	labelKey: AUTOFIT_LABEL_KEYS[value],
}));

/** The Text Box options section (no title: each pane heads its text options). */
export const TEXT_BOX_SECTION: TextBoxSection = {
	id: 'textBox',
	fields: [
		{
			id: 'vAlign',
			kind: 'select',
			labelKey: 'pptx.textPanel.verticalAlign',
			options: VALIGN_CHOICES,
			testId: 'text-box-valign',
			read: (ts) => ts?.vAlign ?? 'top',
			write: (_ts, v) => ({ vAlign: String(v) as NonNullable<TextStyle['vAlign']> }),
		},
		{
			id: 'textDirection',
			kind: 'select',
			labelKey: 'pptx.textAdvanced.direction',
			options: DIRECTION_CHOICES,
			testId: 'text-box-direction',
			read: (ts) => ts?.textDirection ?? 'horizontal',
			write: (_ts, v) => ({
				textDirection: String(v) as NonNullable<TextStyle['textDirection']>,
			}),
		},
		{
			id: 'autoFitMode',
			kind: 'select',
			labelKey: 'pptx.textAdvanced.autoFit',
			options: AUTOFIT_CHOICES,
			testId: 'text-box-autofit',
			read: (ts) => ts?.autoFitMode ?? 'none',
			// `autoFit` mirrors the mode, as `autoFitModePatch` writes it.
			write: (_ts, v) => {
				const mode = String(v) as NonNullable<TextStyle['autoFitMode']>;
				return { autoFitMode: mode, autoFit: mode !== 'none' };
			},
		},
		{
			id: 'textWrap',
			kind: 'toggle',
			labelKey: 'pptx.textAdvanced.wrapText',
			wide: true,
			testId: 'text-box-wrap',
			read: (ts) => (ts?.textWrap ?? 'square') === 'square',
			write: (_ts, v) => ({ textWrap: v === true ? 'square' : 'none' }),
		},
		{
			id: 'columnCount',
			kind: 'number',
			labelKey: 'pptx.textFormatting.columns',
			min: 1,
			max: MAX_COLUMNS,
			decimals: 0,
			testId: 'text-box-columns',
			read: (ts) => ts?.columnCount ?? 1,
			write: (_ts, v) => ({ columnCount: Math.round(Number(v)) }),
		},
	],
};

/** The Text Box options as a section list, what a generic renderer takes. */
export const TEXT_BOX_SECTIONS: readonly TextBoxSection[] = [TEXT_BOX_SECTION];
