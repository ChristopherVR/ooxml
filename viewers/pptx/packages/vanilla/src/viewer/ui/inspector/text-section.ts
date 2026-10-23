import { TEXT_BOX_SECTIONS } from 'ooxml-ui/pptx';
import type { TextAdvancedChanges } from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import { makeColorControl, makeNumberField } from '../controls';
import { createRecentColorsRow } from '../recent-colors-row';
import { createThemeColorSwatchGrid } from '../theme-color-swatch-grid';
import { makeCheckboxField } from './controls-extra';
import { createInspectorFields } from './inspector-fields';
import { createTextEffectsControls } from './text-effects-controls';
import type { InspectorHandlers, InspectorState } from './types';

export interface TextSection {
	el: HTMLElement;
	update(state: InspectorState): void;
}

/**
 * The Text section: text colour, the shared Text Box options
 * (`TEXT_BOX_SECTIONS`: vertical anchor, direction, autofit mode,
 * wrap-in-shape, columns, drawn by the generic renderer), paragraph spacing
 * and text effects.
 * Text alignment (left/center/right/justify) is already covered by the
 * ribbon's Home > Paragraph group, so it is intentionally not duplicated here.
 */
export function createTextSection(
	doc: Document,
	t: Translator,
	section: (label: string) => HTMLElement,
	handlers: InspectorHandlers,
): TextSection {
	const el = section(t('pptx.inspector.text'));

	// B6 (A3): the inspector's own text-colour picker, distinct from the
	// ribbon's Home > Font colour swatch picker (both write the same
	// `textStyle.color`). `onInput` commits live, matching every other colour
	// field here; the "Recent colours" row below commits through the same
	// `setTextStyle` call and pushes on both the row click and the native
	// picker's `change` (never the continuous `input` a drag fires).
	const colorRow = doc.createElement('div');
	colorRow.className = 'pptxv-inspector-row';
	const colorLabel = doc.createElement('span');
	colorLabel.className = 'pptxv-inspector-row-label';
	colorLabel.textContent = t('pptx.textPanel.color');
	const color = makeColorControl(
		doc,
		{
			label: t('pptx.textPanel.color'),
			onInput: (hex) => handlers.setTextStyle({ color: hex, colorRef: undefined }),
			onCommit: handlers.pushRecentColor,
		},
		'#000000',
	);
	colorRow.append(colorLabel, color.el);
	el.appendChild(colorRow);
	// The deck's real "Theme Colors" grid: a theme-swatch click commits both
	// the resolved hex and the ref (so the text colour keeps following the
	// theme after a later theme change); the native input above and the
	// recent-colours row below always clear it.
	const colorTheme = createThemeColorSwatchGrid(doc, t, (commit) =>
		handlers.setTextStyle({ color: commit.hex, colorRef: commit.ref }),
	);
	el.appendChild(colorTheme.el);
	const colorRecent = createRecentColorsRow(doc, t, (hex) => {
		handlers.setTextStyle({ color: hex, colorRef: undefined });
		handlers.pushRecentColor(hex);
	});
	el.appendChild(colorRecent.el);

	// Text box options (vertical anchor, direction, autofit, wrap, columns):
	// body-level text style, so no inline selection is passed.
	const textBox = createInspectorFields(doc, t, {
		sections: TEXT_BOX_SECTIONS,
		onPatch: (patch) => handlers.setTextStyle(patch),
	});
	el.appendChild(textBox.el);
	const number = (label: string, key: keyof TextAdvancedChanges, min?: number) =>
		makeNumberField(doc, {
			label,
			min,
			onCommit: (value) => handlers.setTextAdvanced({ [key]: value }),
		});
	const characterSpacing = number(t('pptx.textAdvanced.characterSpacing'), 'characterSpacing');
	const lineSpacing = number(t('pptx.textAdvanced.lineSpacing'), 'lineSpacing', 0);
	const exactSpacing = number(t('pptx.textAdvanced.lineSpacingExact'), 'lineSpacingExactPt', 0);
	const spacingBefore = number(t('pptx.textAdvanced.spaceBefore'), 'paragraphSpacingBefore', 0);
	const spacingAfter = number(t('pptx.textAdvanced.spaceAfter'), 'paragraphSpacingAfter', 0);
	const indent = number(t('pptx.textAdvanced.indent'), 'paragraphIndent');
	const margin = number(t('pptx.textAdvanced.marginLeft'), 'paragraphMarginLeft');
	const rtl = makeCheckboxField(doc, {
		label: t('pptx.textAdvanced.rtl'),
		onChange: (enabled) => handlers.setTextAdvanced({ rtl: enabled }),
	});
	const advanced = [
		characterSpacing,
		lineSpacing,
		exactSpacing,
		spacingBefore,
		spacingAfter,
		indent,
		margin,
	];
	el.append(...advanced.map(({ el: node }) => node), rtl.el);
	const effects = createTextEffectsControls(doc, t, handlers);
	el.appendChild(effects.el);

	const gated = [color, ...advanced, rtl];

	return {
		el,
		update(state) {
			el.hidden = !state.hasSelection || !state.canText;
			color.setValue(state.textStyle?.color);
			colorTheme.setThemeColorMap(state.themeColorMap);
			colorTheme.setSelected(state.textStyle?.colorRef, state.textStyle?.color);
			colorTheme.setDisabled(!state.canText);
			colorRecent.setColors(state.recentColors ?? []);
			colorRecent.setDisabled(!state.canText);
			textBox.update(state.textStyle, !state.canText);
			characterSpacing.setValue(state.characterSpacing);
			lineSpacing.setValue(state.lineSpacing);
			exactSpacing.setValue(state.lineSpacingExactPt ?? 0);
			spacingBefore.setValue(state.paragraphSpacingBefore);
			spacingAfter.setValue(state.paragraphSpacingAfter);
			indent.setValue(state.paragraphIndent);
			margin.setValue(state.paragraphMarginLeft);
			rtl.setValue(state.textRtl);
			effects.update(state);
			for (const c of gated) {
				c.setDisabled(!state.canText);
			}
		},
	};
}
