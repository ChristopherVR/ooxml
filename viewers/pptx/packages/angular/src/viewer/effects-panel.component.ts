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
 *   (preview)     : emits a colour drag's live shape-style patch and the id of
 *                   the element it is for, to merge with no undo step
 *   [lookup]      : the deck's element by id now, which a drag the selection
 *                   left is put back on (none: such a drag is not put back)
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
			[targetKey]="element().id"
			[lookup]="styleLookup"
			(patch)="onEffectPatch($event)"
			(preview)="onEffectPreview($event)"
		/>
	`,
})
export class EffectsPanelComponent {
	/** The element whose effects are being edited. */
	readonly element = input.required<PptxElement>();
	/** The deck's element `id` now (what a colour drag the selection left is put back on). */
	readonly lookup = input<(id: string) => PptxElement | undefined>(() => undefined);

	/**
	 * Emits a Partial<PptxElement> patch each time the user commits a change.
	 * The orchestrator should call EditorStateService.updateElement(slideIndex, element().id, patch).
	 */
	readonly patch = output<Partial<PptxElement>>();
	/**
	 * Emits a colour drag's live shape-style patch with the id of the element it
	 * is for (the selection, or one the drag has left), to merge without an undo step.
	 */
	readonly preview = output<{ patch: Partial<ShapeStyle>; target: string }>();

	protected readonly sections = SHAPE_EFFECT_SECTIONS;

	/** The shape style of element `id` now, for the effect colour fields. */
	protected readonly styleLookup = (id: string): ShapeStyle | undefined => {
		const el = this.lookup()(id);
		return el && hasShapeProperties(el) ? el.shapeStyle : undefined;
	};

	/** The element's shape style, the value the effect descriptors read. */
	protected readonly shapeStyle = computed<ShapeStyle | undefined>(() => {
		const el = this.element();
		return hasShapeProperties(el) ? el.shapeStyle : undefined;
	});

	/** Merge an effect field's shape-style patch into the element. */
	protected onEffectPatch(patch: Partial<ShapeStyle>): void {
		this.patch.emit(shapeQuickStylePatch(this.element(), patch));
	}

	/** Pass an effect colour drag's shade up as it is: the inspector merges it into its element. */
	protected onEffectPreview(preview: { patch: Partial<ShapeStyle>; target: string }): void {
		this.preview.emit(preview);
	}
}
