/**
 * line-format-panel.component.ts: Standalone Angular component for editing a
 * shape's line options (width, transparency, dash, compound, join, cap); the
 * line colour stays with the inspector's compact stroke colour control.
 *
 * Selector: `pptx-line-format-panel`
 *
 * The fields are the shared `SHAPE_LINE_SECTIONS` descriptors drawn by the
 * generic `pptx-inspector-fields` renderer; each `Partial<ShapeStyle>` patch
 * is merged into the element with the shared `shapeQuickStylePatch`.
 *
 * Contract:
 *   [element]     : the selected PptxElement (required)
 *   (patch)       : emits a Partial<PptxElement> for the orchestrator to
 *                   commit via EditorStateService.updateElement
 *
 * @module viewer/line-format-panel
 */
import { ChangeDetectionStrategy, Component, computed, input, output } from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import { SHAPE_LINE_SECTIONS, shapeQuickStylePatch } from 'ooxml-ui/pptx';
import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import { hasShapeProperties } from 'pptx-viewer-core';

import { InspectorFieldsComponent } from './inspector-fields.component';

@Component({
	selector: 'pptx-line-format-panel',
	standalone: true,
	imports: [TranslatePipe, InspectorFieldsComponent],
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		@if (!applicable()) {
			<p class="pptx-ng-lf__muted">{{ 'pptx.stroke.noBorderProperties' | translate }}</p>
		} @else {
			<pptx-inspector-fields
				class="pptx-ng-lf"
				[sections]="sections"
				[target]="style()"
				(patch)="onLinePatch($event)"
			/>
		}
	`,
	styles: `
		.pptx-ng-lf__muted {
			color: var(--pptx-inspector-muted, #888);
			font-style: italic;
			margin: 0;
			padding: 0.5rem;
			font-size: 12px;
		}
	`,
})
export class LineFormatPanelComponent {
	readonly element = input.required<PptxElement>();
	readonly patch = output<Partial<PptxElement>>();

	protected readonly sections = SHAPE_LINE_SECTIONS;

	protected readonly applicable = computed(() => hasShapeProperties(this.element()));

	/** The element's shape style, the value the line descriptors read. */
	protected readonly style = computed<ShapeStyle | undefined>(() => {
		const el = this.element();
		return hasShapeProperties(el) ? el.shapeStyle : undefined;
	});

	/** Merge a line field's shape-style patch into the element. */
	protected onLinePatch(patch: Partial<ShapeStyle>): void {
		this.patch.emit(shapeQuickStylePatch(this.element(), patch));
	}
}
