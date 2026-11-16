import {
	CHARACTER_SPACING_SECTIONS,
	PARAGRAPH_SPACING_SECTIONS,
	TEXT_BOX_SECTIONS,
} from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import { makeColorControl } from '../controls';
import { createRecentColorsRow } from '../recent-colors-row';
import { createThemeColorSwatchGrid } from '../theme-color-swatch-grid';
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
 * wrap-in-shape, columns), the shared character and paragraph spacing
 * (`CHARACTER_SPACING_SECTIONS`, `PARAGRAPH_SPACING_SECTIONS`), all drawn by
 * the generic renderer, and text effects.
 * Text alignment (left/center/right/justify) is already covered by the
 * ribbon's Home > Paragraph group, so it is intentionally not duplicated here.
 */
export function createTextSection(
	doc: Document,
	t: Translator,
	section: (label: string) => HTMLElement,
	handlers: InspectorHandlers,
	/** The Format pane's Paragraph home for the text box and spacing fields (default: `el`). */
	paragraph?: HTMLElement,
): TextSection {
	const el = section(t('pptx.inspector.text'));
	const paragraphEl = paragraph ?? el;

	// B6 (A3): the inspector's own text-colour picker, distinct from the
	// ribbon's Home > Font colour swatch picker (both write the same
	// `textStyle.color`). It is live while dragging, like the fill and line
	// pickers (`editor-live-colors.ts`); the "Recent colours" row below
	// commits through `setTextStyle`.
	const colorRow = doc.createElement('div');
	colorRow.className = 'pptxv-inspector-row';
	const colorLabel = doc.createElement('span');
	colorLabel.className = 'pptxv-inspector-row-label';
	colorLabel.textContent = t('pptx.textPanel.color');
	const color = makeColorControl(
		doc,
		{
			label: t('pptx.textPanel.color'),
			onInput: (hex) => handlers.liveColor('text', 'input', hex),
			onCommit: (hex) => handlers.liveColor('text', 'change', hex),
			onBlur: (hex) => handlers.liveColor('text', 'blur', hex),
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
	paragraphEl.appendChild(textBox.el);
	// Character spacing (Font) and indentation, spacing, line spacing and
	// direction (Paragraph): the shared descriptors, through the same
	// body-level text-style route.
	const characterSpacing = createInspectorFields(doc, t, {
		sections: CHARACTER_SPACING_SECTIONS,
		onPatch: (patch) => handlers.setTextStyle(patch),
	});
	const paragraphSpacing = createInspectorFields(doc, t, {
		sections: PARAGRAPH_SPACING_SECTIONS,
		onPatch: (patch) => handlers.setTextStyle(patch),
	});
	el.append(characterSpacing.el);
	paragraphEl.append(paragraphSpacing.el);
	const effects = createTextEffectsControls(doc, t, handlers);
	el.appendChild(effects.el);

	return {
		el,
		update(state) {
			el.hidden = !state.hasSelection || !state.canText;
			paragraphEl.hidden = el.hidden;
			color.setValue(state.textStyle?.color);
			colorTheme.setThemeColorMap(state.themeColorMap);
			colorTheme.setSelected(state.textStyle?.colorRef, state.textStyle?.color);
			colorTheme.setDisabled(!state.canText);
			colorRecent.setColors(state.recentColors ?? []);
			colorRecent.setDisabled(!state.canText);
			textBox.update(state.textStyle, !state.canText);
			characterSpacing.update(state.textStyle, !state.canText);
			paragraphSpacing.update(state.textStyle, !state.canText);
			effects.update(state);
			color.setDisabled(!state.canText);
		},
	};
}
