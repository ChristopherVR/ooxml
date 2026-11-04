/**
 * slide-thumbnail.component.ts: one slide drawn the way the slides pane draws
 * it, for a host that keeps its own thumbnail strip.
 *
 * Selector: `pptx-slide-thumbnail`
 *
 * @module angular-viewer/slide-thumbnail
 */
import { ChangeDetectionStrategy, Component, computed, inject, input } from '@angular/core';
import { buildPreviewElements, EDITOR_THUMBNAIL_WIDTH, slideThumbnailLayout } from 'ooxml-ui/pptx';
import type { CanvasSize } from 'ooxml-ui/pptx';
import type {
	ParsedTableStyleMap,
	PptxCustomProperty,
	PptxElement,
	PptxHeaderFooter,
	PptxSlide,
	PptxTheme,
} from 'pptx-viewer-core';

import { FieldContextService } from './field-context.service';
import { LoadContentService } from './load-content.service';
import { SlideCanvasComponent } from './slide-canvas.component';
import { createSlidesPanelContexts } from './slides-panel-contexts';
import { TABLE_STYLE_CONTEXT_OVERRIDE } from './table-style-context';

/** No template elements: the stage gets them merged into the preview slide instead. */
const NO_TEMPLATE_ELEMENTS: readonly PptxElement[] = [];

/**
 * SlideThumbnailComponent: a static, scaled-down drawing of one slide, as the
 * slides pane (`SlidesPanelComponent`) shows it. It renders from the slide
 * model, so it follows unsaved edits.
 *
 * The slide's template (layout and master) elements are drawn beneath its own,
 * with the slide background, through the rail's non-editable
 * `SlideCanvasComponent` and the rail's field and table style contexts. The
 * box fills its container's width and is as tall as `previewWidth` at the
 * deck's aspect ratio; the stage inside is drawn at the full canvas size and
 * scaled from the top-left. Nothing in it is selectable, editable or focusable,
 * and its elements carry no `data-element-id`, so it never competes with the
 * editing canvas.
 *
 * Feed it from the values a host already binds on `<pptx-slides-panel>`:
 * `templateElementsBySlideId[slide.id]`, `canvasSize`, `mediaDataUrls`,
 * `headerFooter`, `customProperties`, `firstSlideNumber`, `locale`, `theme`
 * and `tableStyleMap` (each context falls back to the enclosing deck loader
 * when unset, as the rail's do).
 * ```html
 * @for (slide of shell.editor.slides(); track slide.id) {
 *   <pptx-slide-thumbnail
 *     [slide]="slide"
 *     [templateElements]="shell.editor.templateElementsBySlideId()[slide.id] ?? []"
 *     [canvasSize]="shell.loader.canvasSize()"
 *     [mediaDataUrls]="shell.loader.mediaDataUrls()"
 *     [previewWidth]="160"
 *   />
 * }
 * ```
 */
@Component({
	selector: 'pptx-slide-thumbnail',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [SlideCanvasComponent],
	// The stage resolves fields and table styles from these inputs (else the
	// deck loader), exactly as the rail's thumbnails do.
	viewProviders: [
		{
			provide: FieldContextService,
			useFactory: () => inject(SlideThumbnailComponent).contexts.fields,
		},
		{
			provide: TABLE_STYLE_CONTEXT_OVERRIDE,
			useFactory: () => inject(SlideThumbnailComponent).contexts.tableStyle,
		},
	],
	template: `
		<div class="pptx-ng-slide-thumbnail" [style.height.px]="layout().height">
			<pptx-slide-canvas
				[slide]="previewSlide()"
				[templateElements]="noTemplateElements"
				[canvasSize]="stageSize()"
				[mediaDataUrls]="mediaDataUrls()"
				[zoom]="layout().scale"
				[editable]="false"
				[autoFit]="false"
				[interactive]="false"
				[exposeElementIds]="false"
			/>
		</div>
	`,
	styles: `
		:host {
			display: block;
			width: 100%;
		}
		.pptx-ng-slide-thumbnail {
			position: relative;
			width: 100%;
			overflow: hidden;
		}
		/* Anchor the stage at the top-left instead of the canvas's centring margin. */
		.pptx-ng-slide-thumbnail ::ng-deep .pptx-ng-canvas-wrapper {
			margin: 0 !important;
		}
	`,
})
export class SlideThumbnailComponent {
	/** The slide to draw. A new slide object (an edit) redraws the thumbnail. */
	readonly slide = input.required<PptxSlide>();
	/** The slide's inherited layout and master elements, drawn beneath its own. */
	readonly templateElements = input<readonly PptxElement[]>([]);
	/** The deck's slide size in px; the thumbnail keeps its aspect ratio. */
	readonly canvasSize = input.required<CanvasSize>();
	/** Thumbnail width in px; the height follows the deck's aspect ratio. Default: the rail's. */
	readonly previewWidth = input<number>(EDITOR_THUMBNAIL_WIDTH);
	/** Media asset lookup table (package path to data URL), as the rail takes it. */
	readonly mediaDataUrls = input<Map<string, string>>(new Map());

	// ── Thumbnail contexts (each falls back to the enclosing deck loader) ────

	/** Deck header/footer settings for date, footer and header fields. */
	readonly headerFooter = input<PptxHeaderFooter | undefined>(undefined);
	/** Custom document properties for `docproperty` fields. */
	readonly customProperties = input<readonly PptxCustomProperty[] | undefined>(undefined);
	/** `p:presentation/@firstSlideNum`: the number slide 1 shows (1 when unset). */
	readonly firstSlideNumber = input<number | undefined>(undefined);
	/** Locale for date and time fields (the browser default when unset). */
	readonly locale = input<string | undefined>(undefined);
	/** Deck theme, for table style colours and fonts. */
	readonly theme = input<PptxTheme | undefined>(undefined);
	/** Parsed `ppt/tableStyles.xml`, for table banding. */
	readonly tableStyleMap = input<ParsedTableStyleMap | undefined>(undefined);

	/** Field and table style contexts the stage is rendered with. */
	readonly contexts = createSlidesPanelContexts(
		{
			headerFooter: this.headerFooter,
			customProperties: this.customProperties,
			firstSlideNumber: this.firstSlideNumber,
			locale: this.locale,
			theme: this.theme,
			tableStyleMap: this.tableStyleMap,
		},
		inject(LoadContentService, { optional: true }),
	);

	protected readonly noTemplateElements = NO_TEMPLATE_ELEMENTS;

	readonly stageSize = computed<CanvasSize>(() => ({
		width: Math.max(1, this.canvasSize().width),
		height: Math.max(1, this.canvasSize().height),
	}));
	readonly layout = computed(() =>
		slideThumbnailLayout(this.stageSize().width, this.stageSize().height, this.previewWidth()),
	);
	/**
	 * Template first, slide elements on top. `buildPreviewElements` honours
	 * "Hide Background Graphics", so the stage gets no separate template layer.
	 */
	readonly previewSlide = computed<PptxSlide>(() => {
		const slide = this.slide();
		return { ...slide, elements: buildPreviewElements(slide, this.templateElements()) };
	});
}
