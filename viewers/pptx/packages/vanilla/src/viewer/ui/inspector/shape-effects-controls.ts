import { SHAPE_EFFECT_SECTIONS, SHAPE_PRESET_DEFS } from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import { createConnectorArrowControls } from './connector-arrow-controls';
import { makeSelectField } from './controls-extra';
import { createInspectorFields } from './inspector-fields';
import type { InspectorHandlers, InspectorState } from './types';

export interface ShapeEffectsControls {
	el: HTMLElement;
	update(state: InspectorState): void;
}

/**
 * The shape's outline preset and its effect stack (outer and inner shadow,
 * glow, soft edge, reflection, blur, top bevel), plus the connector arrowhead
 * pickers.
 *
 * The effects are the shared `SHAPE_EFFECT_SECTIONS` descriptors drawn by the
 * generic `createInspectorFields` renderer, whose rows come from the panel's
 * own field factories so the panel's stylesheet reaches every one of them.
 */
export function createShapeEffectsControls(
	doc: Document,
	t: Translator,
	handlers: InspectorHandlers,
): ShapeEffectsControls {
	const el = createEl(doc, 'div', 'pptxv-inspector-shape-effects');

	const shapeType = makeSelectField<string>(doc, {
		label: t('pptx.shape.type'),
		options: SHAPE_PRESET_DEFS.map((preset) => ({
			value: preset.type as string,
			label: t(preset.i18nKey),
		})),
		onChange: (value) => handlers.setShapeType(value),
	});
	el.appendChild(shapeType.el);

	const effects = createInspectorFields(doc, t, {
		sections: SHAPE_EFFECT_SECTIONS,
		onPatch: (patch) => handlers.setShapeStyle(patch),
		onPreview: (patch) => handlers.previewShapeStyle(patch),
		onColorCommit: handlers.pushRecentColor,
	});
	el.appendChild(effects.el);

	// The six `a:headEnd`/`a:tailEnd` dropdowns are their own card; see
	// `connector-arrow-controls`.
	const arrows = createConnectorArrowControls(doc, t, handlers);
	el.appendChild(arrows.el);

	return {
		el,
		update(state) {
			effects.update(state.shapeStyle, !state.canShape);
			shapeType.setValue(state.shapeType ?? 'rect');
			shapeType.el.hidden = !state.shapeType;
			shapeType.setDisabled(!state.canShape);
			arrows.update(state);
		},
	};
}
