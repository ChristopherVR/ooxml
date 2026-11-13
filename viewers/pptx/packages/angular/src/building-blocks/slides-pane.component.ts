import { ChangeDetectionStrategy, Component, computed, inject, input } from '@angular/core';

import { ElementRenderErrors } from '../viewer/element-boundary.directive';
import { createSlidesPanelEditorCommands } from '../viewer/slides-panel-commands';
import { SlidesPanelComponent } from '../viewer/slides-panel.component';
import type { ViewerBuildingBlocks } from './viewer-building-blocks';

/** The slide thumbnail rail block, fed from the blocks' own deck and active slide. */
@Component({
	selector: 'pptx-slides-pane',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [SlidesPanelComponent],
	host: { 'data-pptx-block': 'slides-pane' },
	providers: [ElementRenderErrors],
	styles: `
		:host {
			display: block;
			overflow: auto;
		}
	`,
	template: `
		@let shell = blocks().shell;
		<pptx-slides-panel
			[slides]="shell.editor.slides()"
			[templateElementsBySlideId]="shell.editor.templateElementsBySlideId()"
			[sections]="shell.editor.sections()"
			[canvasSize]="shell.loader.canvasSize()"
			[mediaDataUrls]="shell.loader.mediaDataUrls()"
			[activeIndex]="shell.activeSlideIndex()"
			[canEdit]="shell.canEdit()"
			[headerFooter]="shell.loader.headerFooter()"
			[customProperties]="shell.loader.customProperties()"
			[firstSlideNumber]="shell.loader.firstSlideNumber()"
			[theme]="shell.loader.theme()"
			[tableStyleMap]="shell.loader.tableStyleMap()"
			(select)="blocks().goTo($event)"
			(addSlide)="commands().addSlideAfter($event)"
			(addSlideAfter)="commands().addSlideAfter($event)"
			(moveSlide)="commands().moveSlide($event)"
			(duplicateSlides)="commands().duplicateSlides($event)"
			(deleteSlides)="commands().deleteSlides($event)"
			(toggleHideSlides)="commands().toggleHideSlides($event)"
			(addSectionAt)="commands().addSectionAt($event)"
			(renameSection)="commands().renameSection($event)"
			(deleteSection)="commands().deleteSection($event)"
			(moveSectionUp)="commands().moveSectionUp($event)"
			(moveSectionDown)="commands().moveSectionDown($event)"
			(toggleSectionCollapse)="commands().toggleSectionCollapse($event)"
			(openLayoutGallery)="blocks().options().onOpenLayoutGallery?.($event)"
		/>
	`,
})
export class BuildingBlocksSlidesPaneComponent {
	readonly blocks = input.required<ViewerBuildingBlocks>();

	constructor() {
		inject(ElementRenderErrors, { self: true }).reporter = () =>
			this.blocks().options().onElementRenderError;
	}

	protected readonly commands = computed(() => {
		const { shell } = this.blocks();
		return createSlidesPanelEditorCommands(shell.editor, () => shell.canEdit());
	});
}
