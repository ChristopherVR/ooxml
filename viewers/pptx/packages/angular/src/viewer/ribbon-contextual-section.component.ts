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
	inject,
	input,
	output,
} from '@angular/core';
import { TranslatePipe, TranslateService } from '@ngx-translate/core';
import { CONTEXTUAL_TAB_GROUPS } from 'ooxml-ui/pptx';
import type {
	PictureRibbonState,
	RibbonContextualTabId,
	RibbonGalleryCommandId,
	ToolbarActionId,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import { EditorStateService } from './editor-state.service';
import { LoadContentService } from './load-content.service';
import { RibbonArrangeSectionComponent } from './ribbon-arrange-section.component';
import { RibbonChartSelectionComponent } from './ribbon-chart-selection.component';
import { RibbonGalleryComponent } from './ribbon-gallery.component';
import { homeLanguage, ngxTranslate } from './ribbon-home-lang';
import { RibbonPictureSizeComponent } from './ribbon-picture-size.component';
import { ViewerInspectorPanelService } from './viewer-inspector-panel.service';

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
		RibbonArrangeSectionComponent,
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
						(command)="onGalleryCommand($event)"
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
				@if (group.pictureStrip === 'picture-arrange') {
					<pptx-ribbon-arrange-section
						[embedded]="true"
						[hiddenActions]="hiddenActions()"
						[slideIndex]="slideIndex()"
						[selectedElement]="selectedElement()"
						[canEdit]="canEdit()"
					/>
				}
				@if (group.pictureStrip; as family) {
					<pptx-ui-picture-ribbon [family]="family" [state]="picture()" />
				}
			</pptx-ui-ribbon-group>
		}
	`,
})
export class RibbonContextualSectionComponent {
	private readonly translation = inject(TranslateService, { optional: true });
	private readonly language = homeLanguage(this.translation);
	private readonly editor = inject(EditorStateService);
	private readonly loader = inject(LoadContentService, { optional: true });
	private readonly inspector = inject(ViewerInspectorPanelService, { optional: true });
	readonly tab = input.required<RibbonContextualTabId>();
	readonly selectedElement = input<PptxElement | null>(null);
	readonly slideIndex = input<number>(0);
	readonly canEdit = input<boolean>(false);
	/** Toolbar buttons the host hides (gates the Size group's Crop). */
	readonly hiddenActions = input<readonly ToolbarActionId[]>([]);
	/** Toggle the Inspector pane (SmartArt Design > Text Pane). */
	readonly toggleInspector = output<void>();

	readonly groups = computed(() => CONTEXTUAL_TAB_GROUPS[this.tab()]);
	protected readonly picture = computed<PictureRibbonState>(() => {
		this.language();
		const slideIndex = this.slideIndex();
		return {
			element: this.selectedElement(),
			editable: this.canEdit(),
			translate: ngxTranslate(this.translation),
			slides: this.editor.slides(),
			patch: (id, patch) => this.editor.updateElement(slideIndex, id, patch),
			loadImage: (path) =>
				this.loader?.getHandler()?.getImageData(path) ?? Promise.resolve(undefined),
			updateSlides: (update) => this.editor.applyReplacement(update([...this.editor.slides()])),
			altText: this.inspector
				? () => this.inspector?.formatPane.requestFormatPane('picture', { focus: true })
				: undefined,
			selectionPane: this.inspector ? () => this.inspector?.togglePanel('selection') : undefined,
		};
	});

	protected onGalleryCommand(command: RibbonGalleryCommandId): void {
		if (command === 'smartArtTextPane') {
			this.toggleInspector.emit();
		}
	}
}
