/**
 * TextAdvancedSections: the text-effect half of the element inspector, shown
 * whenever the selection carries text.
 *
 * WHY this file exists: these four sections (advanced character formatting,
 * the warp gallery, shadow/glow/reflection and 3D extrusion) used to hang off
 * `TextProperties`, which nothing rendered, so React shipped them as dead
 * code. They are grouped here rather than inlined into `ShapeTextPanels` so
 * that file stays a thin presentational shell and neither grows past the
 * 300-line limit.
 *
 * The other bindings show the same Text Warp, Text Effects and 3D text: Vue
 * in `TextPanel.vue`, Angular in its inspector's Text Warp gallery and Text
 * Effects and 3D Text sections, Svelte in `TextEffectsSection.svelte`, vanilla
 * in `text-effects-controls.ts` and `text-3d-section.ts`. Vue, Angular and
 * vanilla draw shadow/glow/reflection from the shared `TEXT_EFFECT_SECTIONS`
 * descriptors; `TextEffectsPanel` here still builds its own controls, with the
 * same defaults.
 *
 * The font/size/colour/alignment controls of the old `TextProperties` are
 * deliberately NOT here: the ribbon's Home tab already owns them, and a second
 * copy in the inspector is the duplication this cluster was cleaned up to
 * remove.
 */
import type { PptxElement, TextStyle } from 'pptx-viewer-core';
import { hasTextProperties } from 'pptx-viewer-core';
import React from 'react';

import { Text3DProperties } from './properties/Text3DProperties';
import { TextEffectsPanel } from './TextEffectsPanel';
import { AdvancedTextFormatting, createNumericChangeHandler } from './TextPropertiesHelpers';
import { TextWarpGallery } from './TextWarpGallery';

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

interface TextAdvancedSectionsProps {
	selectedElement: PptxElement;
	canEdit: boolean;
	onUpdateTextStyle: (updates: Partial<TextStyle>) => void;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function TextAdvancedSections({
	selectedElement,
	canEdit,
	onUpdateTextStyle,
}: TextAdvancedSectionsProps): React.ReactElement | null {
	// Only text-capable elements (text boxes, shapes with a text body, ...)
	// expose a `textStyle`; everything else renders nothing at all.
	if (!hasTextProperties(selectedElement)) {
		return null;
	}

	const ts = selectedElement.textStyle;
	// Shared factory so each numeric field rejects a non-finite entry instead of
	// writing NaN into the deck.
	const numChange = createNumericChangeHandler(onUpdateTextStyle);

	return (
		<div data-pptx-text-advanced>
			<AdvancedTextFormatting
				ts={ts}
				canEdit={canEdit}
				onUpdateTextStyle={onUpdateTextStyle}
				numChange={numChange}
			/>

			<TextWarpGallery ts={ts} onUpdateTextStyle={onUpdateTextStyle} />

			<TextEffectsPanel ts={ts} onUpdateTextStyle={onUpdateTextStyle} numChange={numChange} />

			<Text3DProperties ts={ts} onUpdateTextStyle={onUpdateTextStyle} />
		</div>
	);
}
