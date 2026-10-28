import {
	ChangeDetectionStrategy,
	Component,
	computed,
	inject,
	input,
	signal,
	viewChild,
} from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import {
	CollaborationCursorsComponent,
	createSlidesPanelEditorCommands,
	describeCollaborationShellState,
	POWER_POINT_VIEWER_PROVIDERS,
	RemoteSelectionOverlayComponent,
	SlideCanvasComponent,
	SlidesPanelComponent,
	ViewerCollaborationShellService,
} from 'pptx-angular-viewer';

import type { HostOwnedDemo } from '../../shared/host-owned-collaboration';

/** Host-defined chrome and the public canvas, without mounting the full viewer. */
@Component({
	selector: 'app-host-owned-headless-editor',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	providers: [...POWER_POINT_VIEWER_PROVIDERS, ViewerCollaborationShellService],
	imports: [
		SlideCanvasComponent,
		SlidesPanelComponent,
		CollaborationCursorsComponent,
		RemoteSelectionOverlayComponent,
	],
	styles: `
		:host {
			display: block;
			height: 100%;
		}
		section {
			height: 100%;
			display: flex;
			flex-direction: column;
		}
		.custom-body {
			flex: 1;
			min-height: 0;
			display: flex;
		}
		pptx-slides-panel {
			flex: none;
		}
		pptx-slide-canvas {
			flex: 1;
			min-width: 0;
			min-height: 0;
		}
	`,
	template: `
		<section
			data-host-custom-shell
			[attr.aria-busy]="shell.loader.loading()"
			(pointermove)="shell.cursor.onPointerMove($event)"
		>
			<output [attr.aria-label]="translate('pptx.collaboration.shellStatusLabel')">{{
				statusText()
			}}</output>
			@if (shell.loader.error(); as error) {
				<p role="alert">{{ error }}</p>
			}
			<div class="custom-body">
				<!-- The public thumbnail rail, fed from the shell's own state. -->
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
					(select)="goTo($event)"
					(addSlide)="rail.addSlideAfter($event)"
					(addSlideAfter)="rail.addSlideAfter($event)"
					(moveSlide)="rail.moveSlide($event)"
					(duplicateSlides)="rail.duplicateSlides($event)"
					(deleteSlides)="rail.deleteSlides($event)"
					(toggleHideSlides)="rail.toggleHideSlides($event)"
					(addSectionAt)="rail.addSectionAt($event)"
					(renameSection)="rail.renameSection($event)"
					(deleteSection)="rail.deleteSection($event)"
					(moveSectionUp)="rail.moveSectionUp($event)"
					(moveSectionDown)="rail.moveSectionDown($event)"
					(toggleSectionCollapse)="rail.toggleSectionCollapse($event)"
				/>
				<pptx-slide-canvas
					[slide]="shell.activeSlide()"
					[canvasSize]="shell.loader.canvasSize()"
					[mediaDataUrls]="shell.loader.mediaDataUrls()"
					[editable]="shell.canEdit()"
					[autoFit]="false"
					[zoom]="scale()"
					[selectedIds]="shell.editor.selectedIds()"
					[editingId]="shell.canvasEditing.editingId()"
					[templateElements]="shell.activeTemplateElements()"
					(elementSelect)="shell.canvasEditing.onElementSelect($event)"
					(backgroundClick)="shell.canvasEditing.onBackgroundClick()"
					(transformStart)="shell.canEdit() && shell.editor.beginTransform($event.label)"
					(transformUpdate)="
						shell.canEdit() &&
							shell.editor.applyTransform(shell.activeSlideIndex(), $event.id, $event.box)
					"
					(transformEnd)="
						shell.canEdit() && shell.editor.rerouteConnectors(shell.activeSlideIndex(), $event.ids)
					"
					(rotateUpdate)="
						shell.canEdit() &&
							shell.editor.applyTransform(shell.activeSlideIndex(), $event.id, {
								rotation: $event.rotation,
							})
					"
					(textEditStart)="shell.canvasEditing.onTextEditStart($event.id)"
					(textInput)="shell.canvasEditing.onTextInput($event)"
					(textCommit)="shell.canvasEditing.onTextCommit($event)"
					(textCancel)="shell.canvasEditing.editingId.set(null)"
					(listSession)="shell.canvasEditing.onListSession($event)"
				>
					<pptx-collaboration-cursors [cursors]="shell.cursor.cursors()" />
					<pptx-remote-selection-overlay
						[presences]="shell.collaboration.presence()"
						[elements]="shell.activeSlide()?.elements ?? []"
						[activeSlideIndex]="shell.activeSlideIndex()"
					/>
				</pptx-slide-canvas>
			</div>
		</section>
	`,
})
export class HostOwnedHeadlessEditorComponent {
	readonly host = input.required<HostOwnedDemo>();
	readonly shell = inject(ViewerCollaborationShellService);
	readonly scale = signal(1);
	private readonly canvas = viewChild(SlideCanvasComponent);
	private readonly translateService = inject(TranslateService);
	readonly translate = (key: string, params?: Record<string, string | number>): string =>
		this.translateService.instant(key, params);
	readonly statusText = computed(() =>
		describeCollaborationShellState(this.shell.state(), this.translate),
	);
	/** The rail's commands, through the same editor (and so the same sync) as the canvas. */
	readonly rail = createSlidesPanelEditorCommands(this.shell.editor, () => this.shell.canEdit());

	constructor() {
		this.shell.bind({
			content: () => this.host().source,
			collaboration: () => this.host().config,
			canEdit: () => this.host().editable,
			stageElement: () => this.canvas()?.getStageElement(),
		});
		// A slide added from the rail is shown, as PowerPoint does.
		this.shell.editor.bindSlideNavigation((index) => this.goTo(index));
	}

	goTo(index: number): void {
		if (index >= 0 && index < this.shell.editor.slides().length) {
			this.shell.activeSlideIndex.set(index);
		}
	}

	getContent(): Promise<Uint8Array> {
		return this.shell.getContent();
	}

	setScale(value: number): void {
		if (Number.isFinite(value) && value > 0) {
			this.scale.set(value);
		}
	}
}
