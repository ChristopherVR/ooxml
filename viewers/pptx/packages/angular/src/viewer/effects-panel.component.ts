/**
 * effects-panel.component.ts: Standalone Angular component for editing the
 * visual effects (outer and inner shadow, glow, soft edge, reflection, blur,
 * top bevel) of a selected PPTX element.
 *
 * Selector: `pptx-effects-panel`
 *
 * The fields are the shared `SHAPE_EFFECT_SECTIONS` descriptors drawn by the
 * generic `pptx-inspector-fields` renderer; each `Partial<ShapeStyle>` patch
 * is merged into the element with the shared `shapeQuickStylePatch`, so this
 * component holds no field logic of its own.
 *
 * Contract:
 *   [element]     : the selected PptxElement (required)
 *   (patch)       : emits a Partial<PptxElement> for the orchestrator to
 *                   commit via EditorStateService.updateElement
 *   (preview)     : emits a colour drag's live Partial<PptxElement>, for
 *                   EditorStateService.previewElement (no undo step)
 *
 * @module viewer/effects-panel
 */

import { ChangeDetectionStrategy, Component, computed, input, output } from '@angular/core';
import { SHAPE_EFFECT_SECTIONS, shapeQuickStylePatch } from 'ooxml-ui/pptx';
import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import { hasShapeProperties } from 'pptx-viewer-core';

import { InspectorFieldsComponent } from './inspector-fields.component';

@Component({
	selector: 'pptx-effects-panel',
	standalone: true,
	imports: [InspectorFieldsComponent],
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		<pptx-inspector-fields
			class="pptx-ng-fx"
			[sections]="sections"
			[target]="shapeStyle()"
			(patch)="onEffectPatch($event)"
			(preview)="onEffectPreview($event)"
		/>
	`,
})
export class EffectsPanelComponent {
	/** The element whose effects are being edited. */
	readonly element = input.required<PptxElement>();

	/**
	 * Emits a Partial<PptxElement> patch each time the user commits a change.
	 * The orchestrator should call EditorStateService.updateElement(slideIndex, element().id, patch).
	 */
	readonly patch = output<Partial<PptxElement>>();
	/** Emits a colour drag's live patch, to apply without an undo step. */
	readonly preview = output<Partial<PptxElement>>();

	protected readonly sections = SHAPE_EFFECT_SECTIONS;

	/** The element's shape style, the value the effect descriptors read. */
	protected readonly shapeStyle = computed<ShapeStyle | undefined>(() => {
		const el = this.element();
		return hasShapeProperties(el) ? el.shapeStyle : undefined;
	});

	/** Merge an effect field's shape-style patch into the element. */
	protected onEffectPatch(patch: Partial<ShapeStyle>): void {
		this.patch.emit(shapeQuickStylePatch(this.element(), patch));
	}

	/** Merge an effect colour drag's live shape-style patch into the element. */
	protected onEffectPreview(patch: Partial<ShapeStyle>): void {
		this.preview.emit(shapeQuickStylePatch(this.element(), patch));
	}
}
