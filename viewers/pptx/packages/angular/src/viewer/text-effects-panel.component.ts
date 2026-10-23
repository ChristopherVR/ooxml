/**
 * text-effects-panel.component.ts: the inspector's Text Effects (text shadow,
 * glow and reflection) for a selected text-bearing element, as React's
 * `TextEffectsPanel` and Svelte's `TextEffectsSection` offer them. These are
 * the effects of the TEXT (`a:rPr/a:effectLst`); the shape's own effects are
 * `pptx-effects-panel`.
 *
 * Selector: `pptx-text-effects-panel`
 *
 * The fields are the shared `TEXT_EFFECT_SECTIONS` descriptors drawn by the
 * generic `pptx-inspector-fields` renderer. These are run properties, and a run
 * is painted from its own style, so each `Partial<TextStyle>` goes to the body
 * AND every run through the shared `applyTextStyleUpdate` (the ribbon's
 * whole-text update, as React's inspector applies them); this component holds
 * no field logic of its own.
 *
 * Contract:
 *   [element] : the selected PptxElement (required)
 *   (patch)   : emits a Partial<PptxElement> for the orchestrator to commit via
 *               EditorStateService.updateElement (one history entry per edit)
 *
 * @module viewer/text-effects-panel
 */
import { ChangeDetectionStrategy, Component, computed, input, output } from '@angular/core';
import { TEXT_EFFECT_SECTIONS, applyTextStyleUpdate } from 'ooxml-ui/pptx';
import type { PptxElement, TextStyle } from 'pptx-viewer-core';
import { hasTextProperties } from 'pptx-viewer-core';

import { InspectorFieldsComponent } from './inspector-fields.component';

@Component({
	selector: 'pptx-text-effects-panel',
	standalone: true,
	imports: [InspectorFieldsComponent],
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		<pptx-inspector-fields
			class="pptx-ng-text-fx"
			[sections]="sections"
			[target]="textStyle()"
			(patch)="onEffectPatch($event)"
		/>
	`,
})
export class TextEffectsPanelComponent {
	/** The element whose text effects are being edited. */
	readonly element = input.required<PptxElement>();

	/** A partial-element patch for the orchestrator to commit (one history entry). */
	readonly patch = output<Partial<PptxElement>>();

	protected readonly sections = TEXT_EFFECT_SECTIONS;

	/** The element's text style, the value the text effect descriptors read. */
	protected readonly textStyle = computed<TextStyle | undefined>(() => {
		const el = this.element();
		return hasTextProperties(el) ? el.textStyle : undefined;
	});

	/** Apply a text effect field's text-style patch to the body and every run. */
	protected onEffectPatch(patch: Partial<TextStyle>): void {
		this.patch.emit(applyTextStyleUpdate(this.element(), patch, null).patch);
	}
}
