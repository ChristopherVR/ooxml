import { SHAPE_LINE_SECTIONS } from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import { makeColorControl } from '../controls';
import { createRecentColorsRow } from '../recent-colors-row';
import { createThemeColorSwatchGrid } from '../theme-color-swatch-grid';
import { createInspectorFields } from './inspector-fields';
import { createShapeEffectsControls } from './shape-effects-controls';
import { appendColorGroup, createShapeFillControls } from './shape-fill-controls';
import type { InspectorHandlers, InspectorState } from './types';

export interface FillSection {
	el: HTMLElement;
	update(state: InspectorState): void;
}

/**
 * Where the Format pane puts the parts of this section: the line controls
 * (and connector arrowheads) under Outline, the effects under Effects, the
 * shape type under Arrange. Omitted, everything stays in the one section.
 */
export interface FillSectionSplit {
	outline: HTMLElement;
	effects: HTMLElement;
	shapeType?: HTMLElement;
}

/**
 * The Fill & Stroke section: the Fill part is the shared fill descriptors and
 * actions (`createShapeFillControls`); the line colour with its theme grid and
 * recent colours, the shared Line options and the effects follow.
 */
export function createFillSection(
	doc: Document,
	t: Translator,
	section: (label: string) => HTMLElement,
	handlers: InspectorHandlers,
	split?: FillSectionSplit,
): FillSection {
	const el = section(t('pptx.shape.fillStroke'));
	const outlineEl = split?.outline ?? el;
	const effectsEl = split?.effects ?? el;
	const parts = [
		...new Set([el, outlineEl, effectsEl, ...(split?.shapeType ? [split.shapeType] : [])]),
	];
	// Scroll target for the "Format Shape" element context-menu command; see
	// `pptx-viewer-shared`'s `context-menu-inspector-anchor.ts`.
	el.setAttribute('data-pptx-inspector-section', 'fill-stroke');

	const fillControls = createShapeFillControls(doc, t, el, handlers);

	// The native picker reports every `input`, `change` and `blur` to the live
	// colour actions (`editor-live-colors.ts`).
	const stroke = makeColorControl(
		doc,
		{
			label: t('pptx.inspector.line'),
			onInput: (hex) => handlers.liveColor('line', 'input', hex),
			onCommit: (hex) => handlers.liveColor('line', 'change', hex),
			onBlur: (hex) => handlers.liveColor('line', 'blur', hex),
		},
		'#1e3a8a',
	);
	// The deck's "Theme Colors" grid commits BOTH the hex and the ref; the
	// recent-colours row commits through the picker's own handler.
	const strokeTheme = createThemeColorSwatchGrid(doc, t, (commit) =>
		handlers.setShapeStroke(commit.hex, commit.ref),
	);
	const strokeRecent = createRecentColorsRow(doc, t, handlers.setShapeStroke);
	appendColorGroup(
		doc,
		outlineEl,
		t('pptx.inspector.line'),
		stroke.el,
		strokeTheme.el,
		strokeRecent.el,
	);

	// Line options (width, transparency, dash, compound, join, cap): the shared
	// descriptors, drawn by the generic renderer.
	const lineFields = createInspectorFields(doc, t, {
		sections: SHAPE_LINE_SECTIONS,
		onPatch: (patch) => handlers.setShapeStyle(patch),
	});
	outlineEl.appendChild(lineFields.el);

	const effects = createShapeEffectsControls(
		doc,
		t,
		handlers,
		split ? { arrows: outlineEl, ...(split.shapeType ? { shapeType: split.shapeType } : {}) } : {},
	);
	effectsEl.appendChild(effects.el);

	return {
		el,
		update(state) {
			for (const part of parts) {
				part.hidden = !state.hasSelection || !state.canShape;
			}
			fillControls.update(state);
			stroke.setValue(state.strokeColor);
			stroke.setDisabled(!state.canShape);
			lineFields.update(state.shapeStyle, !state.canShape);
			strokeTheme.setThemeColorMap(state.themeColorMap);
			strokeTheme.setSelected(state.strokeColorRef, state.strokeColor);
			strokeTheme.setDisabled(!state.canShape);
			strokeRecent.setColors(state.recentColors ?? []);
			strokeRecent.setDisabled(!state.canShape);
			effects.update(state);
		},
	};
}
