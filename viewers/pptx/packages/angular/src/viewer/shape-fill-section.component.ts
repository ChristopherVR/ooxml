/**
 * shape-fill-section.component.ts: the Format pane's Fill section.
 *
 * Selector: `pptx-shape-fill-section`
 *
 * The fill kind, transparency, gradient type and angle and picture tiling are
 * the shared fill descriptors (`SHAPE_FILL_MODE_SECTIONS`,
 * `SHAPE_FILL_OPTIONS_SECTIONS`) drawn by the generic `pptx-inspector-fields`
 * renderer. What else the kind shows comes from the shared
 * `resolveShapeFillView`: the solid colour (projected by the inspector, which
 * owns its live colour drag), the gradient stops (`pptx-gradient-picker`), the
 * pattern swatches (`pptx-pattern-fill-panel`) and the picture source, whose
 * events are the shared `shapeFillPatch` actions.
 *
 * Contract:
 *   [element]     : the selected PptxElement (required)
 *   (patch)       : emits a Partial<PptxElement> for the orchestrator to
 *                   commit via EditorStateService.updateElement
 *   content       : the solid fill colour controls, shown for a solid fill
 *
 * @module viewer/shape-fill-section
 */
import { ChangeDetectionStrategy, Component, computed, input, output } from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import {
	linkedContentUrl,
	resolveShapeFillView,
	SHAPE_FILL_MODE_SECTIONS,
	SHAPE_FILL_OPTIONS_SECTIONS,
	shapeFillPatch,
	shapeQuickStylePatch,
} from 'ooxml-ui/pptx';
import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import { hasShapeProperties } from 'pptx-viewer-core';

import { injectLinkedContentAllowed } from './external-content.service';
import { GradientPickerComponent } from './gradient-picker.component';
import { InspectorFieldsComponent } from './inspector-fields.component';
import { PatternFillPanelComponent } from './pattern-fill-panel.component';

@Component({
	selector: 'pptx-shape-fill-section',
	standalone: true,
	imports: [
		TranslatePipe,
		InspectorFieldsComponent,
		GradientPickerComponent,
		PatternFillPanelComponent,
	],
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		<pptx-inspector-fields
			[sections]="modeSections"
			[target]="style()"
			(patch)="onStylePatch($event)"
		/>
		@if (view().color) {
			<ng-content />
		}
		<pptx-inspector-fields
			[sections]="optionSections"
			[target]="style()"
			(patch)="onStylePatch($event)"
		/>
		@if (view().stops) {
			<pptx-gradient-picker [element]="element()" (patch)="patch.emit($event)" />
		}
		@if (view().pattern) {
			<pptx-pattern-fill-panel [element]="element()" (patch)="patch.emit($event)" />
		}
		@if (view().image; as image) {
			<div class="pptx-ng-fill-image">
				<label class="pptx-ng-fill-image__field">
					<span class="pptx-ng-fill-image__label">{{
						'pptx.fillAdvanced.imageUrl' | translate
					}}</span>
					<input
						type="text"
						class="pptx-ng-fill-image__input"
						data-testid="fill-image-url"
						placeholder="https://example.com/image.png"
						[value]="image.url"
						(change)="onImageUrl($event)"
					/>
				</label>
				<label class="pptx-ng-fill-image__field">
					<span class="pptx-ng-fill-image__label">{{
						'pptx.fillAdvanced.orChooseFile' | translate
					}}</span>
					<input type="file" accept="image/*" (change)="onImageFile($event)" />
				</label>
				@if (previewUrl(image.url); as preview) {
					<img
						class="pptx-ng-fill-image__preview"
						[src]="preview"
						[alt]="'pptx.fillAdvanced.fillPreviewAlt' | translate"
					/>
				}
			</div>
		}
	`,
	styles: `
		.pptx-ng-fill-image {
			display: flex;
			flex-direction: column;
			gap: 0.4rem;
			padding: 0.5rem;
			font-size: 12px;
		}
		.pptx-ng-fill-image__field {
			display: flex;
			flex-direction: column;
			gap: 0.2rem;
		}
		.pptx-ng-fill-image__label {
			font-size: 10px;
			color: var(--pptx-inspector-muted, #888);
		}
		.pptx-ng-fill-image__input {
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			border: 1px solid var(--pptx-inspector-border, #444);
			color: inherit;
			border-radius: 3px;
			padding: 2px 4px;
			font-size: 12px;
		}
		.pptx-ng-fill-image__preview {
			width: 100%;
			height: 64px;
			object-fit: cover;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 3px;
		}
	`,
})
export class ShapeFillSectionComponent {
	readonly element = input.required<PptxElement>();
	readonly patch = output<Partial<PptxElement>>();

	protected readonly modeSections = SHAPE_FILL_MODE_SECTIONS;
	protected readonly optionSections = SHAPE_FILL_OPTIONS_SECTIONS;

	/** The element's shape style, the value the fill descriptors read. */
	protected readonly style = computed<ShapeStyle | undefined>(() => {
		const el = this.element();
		return hasShapeProperties(el) ? el.shapeStyle : undefined;
	});

	protected readonly view = computed(() => resolveShapeFillView(this.style()));
	private readonly linkedContentAllowed = injectLinkedContentAllowed();
	/** The fill picture preview; a linked one only once this viewer allows linked content. */
	protected previewUrl(url: string | undefined): string | undefined {
		return linkedContentUrl(url, this.linkedContentAllowed());
	}

	/** Merge a fill field's shape-style patch into the element. */
	protected onStylePatch(patch: Partial<ShapeStyle>): void {
		this.patch.emit(shapeQuickStylePatch(this.element(), patch));
	}

	protected onImageUrl(event: Event): void {
		const target = event.target;
		if (target instanceof HTMLInputElement) {
			this.setImage(target.value);
		}
	}

	protected onImageFile(event: Event): void {
		const file = event.target instanceof HTMLInputElement ? event.target.files?.[0] : undefined;
		if (!file) {
			return;
		}
		const reader = new FileReader();
		reader.onload = () => this.setImage(String(reader.result));
		reader.readAsDataURL(file);
	}

	private setImage(url: string): void {
		const patch = shapeFillPatch(this.style(), { type: 'imageUrl', url });
		if (patch) {
			this.onStylePatch(patch);
		}
	}
}
