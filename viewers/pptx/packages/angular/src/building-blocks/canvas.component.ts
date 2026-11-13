import {
	ChangeDetectionStrategy,
	Component,
	effect,
	ElementRef,
	HostListener,
	inject,
	input,
	viewChild,
	ViewEncapsulation,
} from '@angular/core';

import { CollaborationCursorsComponent } from '../viewer/collaboration-cursors.component';
import { RemoteSelectionOverlayComponent } from '../viewer/remote-selection-overlay.component';
import { SlideCanvasComponent } from '../viewer/slide-canvas.component';
import { syncFindHighlights } from 'ooxml-ui/pptx';
import type { ViewerBuildingBlocks } from './viewer-building-blocks';

/**
 * The slide canvas block: the editable canvas over the blocks' deck, with no chrome around it.
 * The public `SlideCanvasComponent` owns `pptx-slide-canvas` and Angular refuses two components on
 * one tag, so this block is `pptx-building-blocks-canvas`.
 */
@Component({
	selector: 'pptx-building-blocks-canvas',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	encapsulation: ViewEncapsulation.None,
	imports: [SlideCanvasComponent, CollaborationCursorsComponent, RemoteSelectionOverlayComponent],
	host: { 'data-pptx-block': 'canvas' },
	styles: `
		[data-pptx-block='canvas'] {
			display: block;
			position: relative;
			min-width: 0;
			min-height: 0;
		}
		[data-pptx-block='canvas'] [data-pptx-find-match] {
			outline: 2px solid #f5a623;
			outline-offset: 1px;
		}
		[data-pptx-block='canvas'] [data-pptx-find-match='current'] {
			outline: 3px solid #e53935;
		}
	`,
	template: `
		@let shell = blocks().shell;
		<pptx-slide-canvas
			style="display: block; width: 100%; height: 100%"
			[slide]="shell.activeSlide()"
			[canvasSize]="shell.loader.canvasSize()"
			[mediaDataUrls]="shell.loader.mediaDataUrls()"
			[editable]="shell.canEdit()"
			[zoom]="blocks().zoom.zoom()"
			[selectedIds]="shell.editor.selectedIds()"
			[editingId]="shell.canvasEditing.editingId()"
			[templateElements]="shell.activeTemplateElements()"
			(pointermove)="shell.cursor.onPointerMove($event)"
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
	`,
})
export class BuildingBlocksCanvasComponent {
	readonly blocks = input.required<ViewerBuildingBlocks>();
	private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
	private readonly canvas = viewChild(SlideCanvasComponent);

	constructor() {
		effect((onCleanup) => {
			const canvas = this.canvas();
			if (canvas) {
				onCleanup(
					this.blocks().attachCanvas(this.host.nativeElement, () => canvas.getStageElement()),
				);
			}
		});
		// The stage repaints on change: outline the find matches again after every paint.
		effect((onCleanup) => {
			const blocks = this.blocks();
			const { session, open } = blocks.findBarState();
			const slideIndex = blocks.shell.activeSlideIndex();
			const draw = () =>
				syncFindHighlights(
					this.host.nativeElement,
					open ? session.matches : [],
					session.current,
					slideIndex,
				);
			draw();
			const observer = new MutationObserver(draw);
			observer.observe(this.host.nativeElement, { childList: true, subtree: true });
			onCleanup(() => observer.disconnect());
		});
	}

	@HostListener('document:keydown', ['$event'])
	protected onKeyDown(event: KeyboardEvent): void {
		this.blocks().handleKeyDown(event);
	}
}
