import { NgStyle } from '@angular/common';
import { ChangeDetectionStrategy, Component, input } from '@angular/core';
import type { CanvasSize } from 'ooxml-ui/pptx';

import type { StyleMap } from './element-style';
import type {
	MorphCrossfadeGroupSlides,
	MorphOverlayLayerSlide,
} from './presentation-transition-overlay-morph';
import { SlideCanvasComponent } from './slide-canvas.component';

/**
 * The morph-only layers `PresentationTransitionOverlayComponent` paints,
 * split out to keep that file under the project's per-file LOC budget:
 *
 *  - the **overlay layers**: the plan's ghosts and lifted arrivals, one layer
 *    per run in the plan's single back-to-front order (issue #141). A lifted
 *    run holds arriving shapes a ghost would otherwise hide for the whole
 *    morph (issue #146); their copy on the live stage is held invisible.
 *  - the **crossfade groups**: pairs the overlay paints BOTH halves of, each
 *    as one isolated group so the halves are SUMMED rather than stacked
 *    (issue #161).
 *
 * Selector: `pptx-morph-extra-layers`.
 */
@Component({
	selector: 'pptx-morph-extra-layers',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [NgStyle, SlideCanvasComponent],
	styles: `
		.pptx-ng-transition-layer {
			position: absolute;
			inset: 0;
			display: flex;
			align-items: center;
			justify-content: center;
		}
	`,
	template: `
		<!-- Ghosts and lifted arrivals, one layer per run of the plan's single
		     back-to-front order. Every layer is background-free, or its slide fill
		     would cover the morph beneath it. -->
		@for (layer of overlayLayers(); track layer.key) {
			<div
				class="pptx-ng-transition-layer"
				[attr.data-pptx-transition-layer]="layer.side === 'outgoing' ? 'outgoing' : 'lifted'"
				[attr.data-pptx-morph-outgoing]="layer.side === 'outgoing' ? 'true' : null"
				[attr.data-pptx-morph-lifted]="layer.side === 'incoming' ? 'true' : null"
				[ngStyle]="layer.style"
			>
				<div [ngStyle]="slideBoxStyle()">
					<pptx-slide-canvas
						[slide]="layer.slide"
						[canvasSize]="canvasSize()"
						[mediaDataUrls]="mediaDataUrls()"
						[zoom]="zoom()"
						[autoFit]="false"
						[interactive]="false"
						[transparentBackground]="true"
					/>
				</div>
			</div>
		}

		<!-- A pair dissolving in place, painted as ONE isolated group whose two
		     halves sum instead of stacking (issue #161). -->
		@for (group of crossfadeGroups(); track group.key) {
			<div [attr.data-pptx-morph-crossfade]="group.key" [ngStyle]="group.style">
				<div
					class="pptx-ng-transition-layer"
					data-pptx-transition-layer="outgoing"
					data-pptx-morph-outgoing="true"
					[ngStyle]="group.outgoingStyle"
				>
					<div [ngStyle]="slideBoxStyle()">
						<pptx-slide-canvas
							[slide]="group.outgoing"
							[canvasSize]="canvasSize()"
							[mediaDataUrls]="mediaDataUrls()"
							[zoom]="zoom()"
							[autoFit]="false"
							[interactive]="false"
							[transparentBackground]="true"
						/>
					</div>
				</div>
				<div
					class="pptx-ng-transition-layer"
					data-pptx-transition-layer="lifted"
					data-pptx-morph-lifted="true"
					[ngStyle]="group.incomingStyle"
				>
					<div [ngStyle]="slideBoxStyle()">
						<pptx-slide-canvas
							[slide]="group.incoming"
							[canvasSize]="canvasSize()"
							[mediaDataUrls]="mediaDataUrls()"
							[zoom]="zoom()"
							[autoFit]="false"
							[interactive]="false"
							[transparentBackground]="true"
						/>
					</div>
				</div>
			</div>
		}
	`,
})
export class MorphExtraLayersComponent {
	readonly overlayLayers = input<readonly MorphOverlayLayerSlide[]>([]);
	readonly crossfadeGroups = input<readonly MorphCrossfadeGroupSlides[]>([]);
	readonly canvasSize = input.required<CanvasSize>();
	readonly mediaDataUrls = input<Map<string, string>>(new Map());
	readonly zoom = input<number>(1);
	/** Slide box sized to the zoomed slide footprint; computed once by the parent. */
	readonly slideBoxStyle = input.required<StyleMap>();
}
