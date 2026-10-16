/**
 * ribbon-contextual-section.component.ts: the content of a contextual ribbon
 * tab (Shape Format, Picture Format, Table Design, Chart Design, SmartArt
 * Design). Renders shared `CONTEXTUAL_TAB_GROUPS[tab]` as ordinary ribbon
 * groups (caption under the controls, `data-ribbon-group` on the wrapper),
 * each holding its galleries in the mode the shared placement names. No
 * placement is decided here, so a gallery shared adds to a tab appears in
 * all five bindings at once. A command group (Picture Format > Size, Chart
 * Format > Current Selection) draws its shared strip after the galleries.
 */
import {
	ChangeDetectionStrategy,
	Component,
	CUSTOM_ELEMENTS_SCHEMA,
	computed,
	input,
} from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import { CONTEXTUAL_TAB_GROUPS } from 'ooxml-ui/pptx';
import type { RibbonContextualTabId, ToolbarActionId } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import { RibbonChartSelectionComponent } from './ribbon-chart-selection.component';
import { RibbonGalleryComponent } from './ribbon-gallery.component';
import { RibbonPictureSizeComponent } from './ribbon-picture-size.component';

@Component({
	selector: 'pptx-ribbon-contextual-section',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	host: { class: 'contents' },
	imports: [
		TranslatePipe,
		RibbonGalleryComponent,
		RibbonPictureSizeComponent,
		RibbonChartSelectionComponent,
	],
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `
		@for (group of groups(); track group.group) {
			<pptx-ui-ribbon-group
				[attr.label]="group.labelKey | translate"
				[attr.data-ribbon-group]="group.group"
			>
				@for (placement of group.galleries; track placement.control) {
					<pptx-ribbon-gallery
						[gallery]="placement.gallery"
						[mode]="placement.mode"
						[control]="placement.control"
						[element]="selectedElement()"
						[slideIndex]="slideIndex()"
						[canEdit]="canEdit()"
					/>
				}
				@if (group.strip === 'picture-size') {
					<pptx-ribbon-picture-size
						[selectedElement]="selectedElement()"
						[slideIndex]="slideIndex()"
						[canEdit]="canEdit()"
						[hiddenActions]="hiddenActions()"
					/>
				}
				@if (group.strip === 'chart-selection') {
					<pptx-ribbon-chart-selection
						[selectedElement]="selectedElement()"
						[slideIndex]="slideIndex()"
						[canEdit]="canEdit()"
					/>
				}
			</pptx-ui-ribbon-group>
		}
	`,
})
export class RibbonContextualSectionComponent {
	readonly tab = input.required<RibbonContextualTabId>();
	readonly selectedElement = input<PptxElement | null>(null);
	readonly slideIndex = input<number>(0);
	readonly canEdit = input<boolean>(false);
	/** Toolbar buttons the host hides (gates the Size group's Crop). */
	readonly hiddenActions = input<readonly ToolbarActionId[]>([]);

	readonly groups = computed(() => CONTEXTUAL_TAB_GROUPS[this.tab()]);
}
