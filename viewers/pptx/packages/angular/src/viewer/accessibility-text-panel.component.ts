import { ChangeDetectionStrategy, Component, computed, input, output } from '@angular/core';
import {
	ELEMENT_ALT_TEXT_SECTIONS,
	getNonVisualDescriptionFields,
	shouldShowAccessibilitySection,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import { InspectorFieldsComponent } from './inspector-fields.component';

/**
 * The Alt Text fields (description and title) of a shape, text box,
 * connector or graphic frame: the shared `ELEMENT_ALT_TEXT_SECTIONS` drawn by
 * the generic `pptx-inspector-fields` renderer. A picture's alt text is drawn
 * in `ImagePropertiesPanelComponent` from the same descriptors;
 * `shouldShowAccessibilitySection` (shared) decides which other kinds get this
 * panel at all.
 */
@Component({
	selector: 'pptx-accessibility-text-panel',
	standalone: true,
	imports: [InspectorFieldsComponent],
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		@if (fields().showAltText || fields().showTitle) {
			<div
				class="pptx-ng-accessibility-text"
				data-pptx-accessibility-text
				data-pptx-inspector-section="alt-text"
			>
				<pptx-inspector-fields
					[sections]="sections"
					[target]="element()"
					[disabled]="!canEdit()"
					(patch)="patch.emit($event)"
				/>
			</div>
		}
	`,
	styles: `
		.pptx-ng-accessibility-text {
			font-size: 11px;
		}
	`,
})
export class AccessibilityTextPanelComponent {
	readonly element = input.required<PptxElement>();
	readonly canEdit = input<boolean>(true);
	readonly patch = output<Partial<PptxElement>>();

	protected readonly sections = ELEMENT_ALT_TEXT_SECTIONS;
	protected readonly fields = computed(() => getNonVisualDescriptionFields(this.element()));

	/** Whether the selected element kind should show this panel at all. */
	static supports(element: PptxElement): boolean {
		return shouldShowAccessibilitySection(element);
	}
}
