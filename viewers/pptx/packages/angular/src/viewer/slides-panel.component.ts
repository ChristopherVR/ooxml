import { NgStyle } from '@angular/common';
import {
	ChangeDetectionStrategy,
	Component,
	computed,
	effect,
	ElementRef,
	HostListener,
	inject,
	input,
	output,
	signal,
	viewChild,
} from '@angular/core';
import { LucidePlus } from '@lucide/angular';
import { TranslatePipe } from '@ngx-translate/core';
import {
	computeVirtualRange,
	EDITOR_THUMBNAIL_WIDTH,
	EDITOR_SLIDE_RAIL_WIDTH,
	editorThumbnailStep,
	groupSlidesBySection,
	HIDDEN_SLIDE_DIM_OPACITY,
	HIDDEN_SLIDE_LABEL_KEY,
	HIDDEN_SLIDE_SLASH_GRADIENT,
	hiddenSlideCue,
	openSlidesPaneLayoutGallery,
	sectionAddAfterSlideIndex,
	SLIDE_VIRTUALIZATION_THRESHOLD,
} from 'ooxml-ui/pptx';
import type {
	CanvasSize,
	HiddenSlideCue,
	SectionContextMenuCommandId,
	SlidesPaneLayoutGalleryRequest,
} from 'ooxml-ui/pptx';
import type {
	ParsedTableStyleMap,
	PptxCustomProperty,
	PptxHeaderFooter,
	PptxSection,
	PptxSlide,
	PptxTheme,
} from 'pptx-viewer-core';

import { EditorStateService } from './editor-state.service';
import { FieldContextService } from './field-context.service';
import { LoadContentService } from './load-content.service';
import { SectionContextMenuComponent } from './section-context-menu.component';
import { SlideCanvasComponent } from './slide-canvas.component';
import { SlidePaneContextMenuComponent } from './slide-pane-context-menu.component';
import { SlidePaneRailSelection } from './slide-pane-rail-selection';
import { thumbnailHeight, thumbnailZoom } from './slide-sorter-overlay-helpers';
import { createSlidesPanelContexts } from './slides-panel-contexts';
import { TABLE_STYLE_CONTEXT_OVERRIDE } from './table-style-context';
import type { TemplateElementsBySlideId } from './template-mode';

/** Pixel width of each thumbnail clipping box inside the panel. Slightly
 *  narrower than the panel so the slide-number column fits to its left
 *  (React SlideItem lays the number beside, not below, the preview). */
const THUMB_W = EDITOR_THUMBNAIL_WIDTH;

/** A slide reorder from the rail's drag and drop. */
export interface SlidesPanelMoveSlide {
	fromIndex: number;
	toIndex: number;
}

/** A section rename committed in the rail's section header. */
export interface SlidesPanelRenameSection {
	sectionId: string;
	name: string;
}

/**
 * SlidesPanelComponent: the slide thumbnail rail (vertical slide strip).
 *
 * Lists every slide as a numbered thumbnail (grouped by section when the deck
 * has sections) and highlights the active one. Clicking a card emits
 * `select(index)`. Slide commands live on the shared thumbnail right-click
 * menu, cards reorder by drag and drop, section commands live on the shared
 * section-header menu, and a footer "Add Slide" button adds a slide after the
 * active one. Thumbnails render through `SlideCanvasComponent` with the same
 * field and table style contexts as the canvas, so slide numbers (including
 * `p:presentation/@firstSlideNum`), header/footer text and table colours match.
 *
 * It runs in one of two modes:
 *
 * - **Host-driven** (a custom shell; set `[slides]`): the rail renders only its
 *   inputs and never edits state itself. Every command is an output, so the host
 *   applies it through its own operations (for the viewer services, see
 *   `createSlidesPanelEditorCommands`). The context menu's Layout entry makes
 *   the slide active (`select`) and emits `openLayoutGallery`.
 * - **Editor-bound** (inside `PowerPointViewerComponent`; `[slides]` unset): the
 *   rail reads the enclosing `EditorStateService` and applies reorder, the footer
 *   add and section commands to it directly, as it always has.
 *
 * Host-driven usage beside a canvas (both under `POWER_POINT_VIEWER_PROVIDERS`):
 * ```html
 * <pptx-slides-panel
 *   [slides]="shell.editor.slides()"
 *   [templateElementsBySlideId]="shell.editor.templateElementsBySlideId()"
 *   [sections]="shell.editor.sections()"
 *   [canvasSize]="shell.loader.canvasSize()"
 *   [mediaDataUrls]="shell.loader.mediaDataUrls()"
 *   [activeIndex]="shell.activeSlideIndex()"
 *   [canEdit]="shell.canEdit()"
 *   (select)="shell.activeSlideIndex.set($event)"
 *   (addSlide)="commands.addSlideAfter($event)"
 * />
 * ```
 */
@Component({
	selector: 'pptx-slides-panel',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [
		NgStyle,
		SlideCanvasComponent,
		SectionContextMenuComponent,
		SlidePaneContextMenuComponent,
		TranslatePipe,
		LucidePlus,
	],
	// The thumbnails resolve fields and table styles from the rail's own
	// contexts (host inputs, else the deck loader), never from the active slide.
	viewProviders: [
		{
			provide: FieldContextService,
			useFactory: () => inject(SlidesPanelComponent).thumbnailContexts.fields,
		},
		{
			provide: TABLE_STYLE_CONTEXT_OVERRIDE,
			useFactory: () => inject(SlidesPanelComponent).thumbnailContexts.tableStyle,
		},
	],
	templateUrl: './slides-panel.component.html',
	styleUrl: './slides-panel.component.css',
})
export class SlidesPanelComponent {
	readonly railWidth = EDITOR_SLIDE_RAIL_WIDTH;

	// ── Deck inputs ───────────────────────────────────────────────────────────

	/**
	 * The slides to list, in deck order, without their inherited layout and
	 * master elements (pass those through {@link templateElementsBySlideId}).
	 * Setting this makes the rail host-driven: it renders these inputs and emits
	 * every command instead of editing the enclosing `EditorStateService`.
	 */
	readonly slides = input<readonly PptxSlide[] | undefined>(undefined);
	/**
	 * Inherited layout and master elements keyed by slide id, drawn behind each
	 * thumbnail's own elements exactly as the canvas draws them. Falls back to the
	 * enclosing `EditorStateService` when unset.
	 */
	readonly templateElementsBySlideId = input<TemplateElementsBySlideId | undefined>(undefined);
	/**
	 * The deck's sections. When any exist the rail groups the thumbnails under
	 * collapsible section headers. Host-driven rails default to none; editor-bound
	 * rails read the enclosing `EditorStateService`.
	 */
	readonly sections = input<readonly PptxSection[] | undefined>(undefined);
	/** Natural (100 %) canvas dimensions, forwarded to each thumbnail canvas. */
	readonly canvasSize = input.required<CanvasSize>();
	/** Media asset lookup table (package path to data URL), forwarded to each thumbnail. */
	readonly mediaDataUrls = input<Map<string, string>>(new Map());
	/** Zero-based index of the currently active slide (highlighted, `aria-current`). */
	readonly activeIndex = input<number>(0);
	/**
	 * Whether the deck can be edited. When false the rail only navigates: no
	 * footer, context menus, drag reorder, Enter-to-add or section rename.
	 */
	readonly canEdit = input<boolean>(true);

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

	// ── Outputs ───────────────────────────────────────────────────────────────

	/** Emits the zero-based index of the card the user clicked (navigation only; no edit). */
	readonly select = output<number>();
	/** Insert a new slide after `index` (thumbnail menu's New Slide, and Enter on the panel). */
	readonly addSlideAfter = output<number>();
	/** Duplicate every listed slide (the rail's whole multi-selection). */
	readonly duplicateSlides = output<number[]>();
	/** Delete every listed slide (the rail's whole multi-selection). */
	readonly deleteSlides = output<number[]>();
	/** Toggle the hidden flag of every listed slide. */
	readonly toggleHideSlides = output<number[]>();
	/** Add a section starting after slide `index`. */
	readonly addSectionAt = output<number>();
	/**
	 * Editor-bound rails only: makes `index` active, then opens the Layout
	 * gallery at (x, y). Host-driven rails emit `select` and
	 * {@link openLayoutGallery} instead.
	 */
	readonly openLayoutForSlide = output<{ index: number; x: number; y: number }>();
	/**
	 * Host-driven rails: the context menu's Layout entry. The rail first emits
	 * `select(slideIndex)`, then this request so the host can open its own Layout
	 * gallery anchored at the click point.
	 */
	readonly openLayoutGallery = output<SlidesPaneLayoutGalleryRequest>();
	/** Host-driven rails: the footer's Add Slide button; carries the active index to add after. */
	readonly addSlide = output<number>();
	/** Host-driven rails: a drag-and-drop reorder. */
	readonly moveSlide = output<SlidesPanelMoveSlide>();
	/** Host-driven rails: a section rename committed in its header. */
	readonly renameSection = output<SlidesPanelRenameSection>();
	/** Host-driven rails: delete the section with this id (its slides stay). */
	readonly deleteSection = output<string>();
	/** Host-driven rails: move the section with this id up one place. */
	readonly moveSectionUp = output<string>();
	/** Host-driven rails: move the section with this id down one place. */
	readonly moveSectionDown = output<string>();
	/** Host-driven rails: collapse or expand the section with this id. */
	readonly toggleSectionCollapse = output<string>();

	/** Ctrl/Shift multi-select + the thumbnail context menu's open state. */
	protected readonly rail = new SlidePaneRailSelection();

	/** The viewer's editor state; required only by an editor-bound rail. */
	protected readonly editor = inject(EditorStateService, { optional: true });
	private readonly loader = inject(LoadContentService, { optional: true });
	/** Field and table style contexts the thumbnails are rendered with. */
	readonly thumbnailContexts = createSlidesPanelContexts(
		{
			headerFooter: this.headerFooter,
			customProperties: this.customProperties,
			firstSlideNumber: this.firstSlideNumber,
			locale: this.locale,
			theme: this.theme,
			tableStyleMap: this.tableStyleMap,
		},
		this.loader,
	);
	private readonly scrollViewport = viewChild<ElementRef<HTMLElement>>('scrollViewport');
	private readonly scrollTop = signal(0);
	private readonly viewportHeight = signal(600);

	// ── Resolved deck (inputs, else the enclosing editor) ────────────────────

	/** True when the host drives the rail through `[slides]`. */
	readonly hostDriven = computed(() => this.slides() !== undefined);
	readonly deck = computed<readonly PptxSlide[]>(
		() => this.slides() ?? this.editor?.slides() ?? [],
	);
	readonly templateMap = computed<TemplateElementsBySlideId>(
		() => this.templateElementsBySlideId() ?? this.editor?.templateElementsBySlideId() ?? {},
	);
	readonly sectionList = computed<readonly PptxSection[]>(
		() => this.sections() ?? (this.hostDriven() ? [] : this.editor?.sections()) ?? [],
	);
	readonly sectionGroups = computed(() => groupSlidesBySection(this.sectionList(), this.deck()));

	// ── Derived thumbnail dimensions ──────────────────────────────────────────

	/** Zoom level that fits the full canvas width into THUMB_W pixels. */
	readonly thumbZoom = computed(() => thumbnailZoom(this.canvasSize().width, THUMB_W));

	/** Pixel height of the clipping box (aspect-correct). */
	readonly thumbH = computed(() =>
		thumbnailHeight(this.canvasSize().width, this.canvasSize().height, THUMB_W),
	);

	/** ngStyle object for the thumbnail clipping box. */
	readonly clipStyle = computed<Record<string, string>>(() => ({
		width: `${THUMB_W}px`,
		height: `${this.thumbH()}px`,
	}));

	readonly itemHeight = computed(() =>
		editorThumbnailStep(this.canvasSize().width, this.canvasSize().height),
	);
	readonly shouldVirtualize = computed(
		() => this.sectionList().length === 0 && this.deck().length >= SLIDE_VIRTUALIZATION_THRESHOLD,
	);
	readonly virtualRange = computed(() =>
		computeVirtualRange(
			this.deck().length,
			this.itemHeight(),
			this.scrollTop(),
			this.viewportHeight(),
		),
	);
	readonly renderedSlides = computed(() => {
		const slides = this.deck();
		if (this.sectionList().length > 0) {
			return this.sectionGroups().flatMap((group) =>
				group.slides.map((slide, offset) => ({
					slide,
					index: group.slideIndexes[offset],
					section: group.section,
					sectionStart: offset === 0,
				})),
			);
		}
		const start = this.shouldVirtualize() ? this.virtualRange().startIndex : 0;
		const end = this.shouldVirtualize() ? this.virtualRange().endIndex : slides.length - 1;
		return slides.slice(start, end + 1).map((slide, offset) => ({
			slide,
			index: start + offset,
			section: undefined,
			sectionStart: false,
		}));
	});

	constructor() {
		effect(() => {
			const viewport = this.scrollViewport()?.nativeElement;
			const index = this.activeIndex();
			if (!viewport || !this.shouldVirtualize()) {
				return;
			}
			const itemHeight = this.itemHeight();
			const top = index * itemHeight;
			const bottom = top + itemHeight;
			if (top < viewport.scrollTop) {
				viewport.scrollTop = top;
			} else if (bottom > viewport.scrollTop + viewport.clientHeight) {
				viewport.scrollTop = Math.max(0, bottom - viewport.clientHeight);
			}
			this.syncViewport(viewport);
		});
	}

	onScroll(): void {
		const viewport = this.scrollViewport()?.nativeElement;
		if (viewport) {
			this.syncViewport(viewport);
		}
	}

	private syncViewport(viewport: HTMLElement): void {
		this.scrollTop.set(viewport.scrollTop);
		this.viewportHeight.set(viewport.clientHeight || 600);
	}

	// ── Drag-to-reorder (the rail's reorder path; there are no per-card buttons) ──
	private dragFrom: number | null = null;

	onDragStart(event: DragEvent, index: number): void {
		if (!this.canEdit()) {
			event.preventDefault();
			return;
		}
		this.dragFrom = index;
		if (event.dataTransfer) {
			event.dataTransfer.effectAllowed = 'move';
			event.dataTransfer.setData('text/plain', String(index));
		}
	}

	onDrop(event: DragEvent, index: number): void {
		event.preventDefault();
		const from = this.dragFrom;
		this.dragFrom = null;
		if (from === null || from === index || !this.canEdit()) {
			return;
		}
		if (this.hostDriven()) {
			this.moveSlide.emit({ fromIndex: from, toIndex: index });
		} else {
			this.editor?.moveSlide(from, index);
		}
	}

	onAddSlide(): void {
		if (this.hostDriven()) {
			this.addSlide.emit(this.activeIndex());
		} else {
			this.editor?.addSlide(this.activeIndex());
		}
	}

	/** A click (even a Ctrl/Shift one) both updates the multi-selection AND
	 * moves the canvas to the slide clicked, matching the sorter overlay. */
	onThumbClick(event: MouseEvent, index: number): void {
		const slide = this.deck()[index];
		if (slide) {
			this.rail.onClick(event, slide.id, this.orderedIds());
		}
		this.select.emit(index);
	}

	onThumbContextMenu(event: MouseEvent, index: number): void {
		event.preventDefault();
		if (this.canEdit()) {
			this.rail.openContextMenu(event.clientX, event.clientY, index, this.orderedIds());
		}
	}

	/** The thumbnail menu's Layout entry (see {@link openLayoutGallery}). */
	onOpenLayout(request: { index: number; x: number; y: number }): void {
		if (!this.hostDriven()) {
			this.openLayoutForSlide.emit(request);
			return;
		}
		openSlidesPaneLayoutGallery(
			{ slideIndex: request.index, x: request.x, y: request.y },
			(index) => this.select.emit(index),
			(gallery) => this.openLayoutGallery.emit(gallery),
		);
	}

	/** PowerPoint's Enter on a focused thumbnail inserts a new slide after it. */
	@HostListener('keydown.enter', ['$event'])
	onPanelKeydown(event: Event): void {
		const target = event.target;
		if (!this.canEdit() || (target instanceof HTMLElement && target.tagName === 'INPUT')) {
			return;
		}
		event.preventDefault();
		this.addSlideAfter.emit(this.activeIndex());
	}

	private orderedIds(): string[] {
		return this.deck().map((s) => s.id);
	}

	// ── Section header menu + inline rename ──
	/** Open state of the section-header menu, or null when closed. */
	protected readonly sectionMenu = signal<{ x: number; y: number; sectionId: string } | null>(null);
	protected readonly renamingSectionId = signal<string | null>(null);
	protected readonly renameValue = signal('');
	private readonly renameInput = viewChild<ElementRef<HTMLInputElement>>('renameInput');

	onSectionContextMenu(event: MouseEvent, sectionId: string): void {
		event.preventDefault();
		if (this.canEdit()) {
			this.sectionMenu.set({ x: event.clientX, y: event.clientY, sectionId });
		}
	}

	onSectionCommand(id: SectionContextMenuCommandId, sectionId: string): void {
		const section = this.sectionList().find((candidate) => candidate.id === sectionId);
		if (!section) {
			return;
		}
		switch (id) {
			case 'rename':
				this.startRenameSection(sectionId, section.name);
				break;
			case 'delete':
				if (this.hostDriven()) {
					this.deleteSection.emit(sectionId);
				} else {
					this.editor?.sectionOps.delete(sectionId);
				}
				break;
			case 'move-up':
			case 'move-down':
				this.moveSection(sectionId, id === 'move-up' ? 'up' : 'down');
				break;
			case 'add-after': {
				const group = this.sectionGroups().find((g) => g.section?.id === sectionId);
				this.addSectionAt.emit(
					sectionAddAfterSlideIndex(
						group?.slideIndexes[group.slideIndexes.length - 1],
						this.deck().length,
					),
				);
				break;
			}
			default:
				break;
		}
	}

	private moveSection(sectionId: string, direction: 'up' | 'down'): void {
		if (!this.hostDriven()) {
			this.editor?.sectionOps.move(sectionId, direction);
		} else if (direction === 'up') {
			this.moveSectionUp.emit(sectionId);
		} else {
			this.moveSectionDown.emit(sectionId);
		}
	}

	/** The section header's collapse toggle. */
	onToggleSection(sectionId: string): void {
		if (this.hostDriven()) {
			this.toggleSectionCollapse.emit(sectionId);
		} else {
			this.editor?.sectionOps.toggle(sectionId);
		}
	}

	startRenameSection(sectionId: string, currentName: string): void {
		if (!this.canEdit()) {
			return;
		}
		this.renamingSectionId.set(sectionId);
		this.renameValue.set(currentName);
		setTimeout(() => {
			const field = this.renameInput()?.nativeElement;
			field?.focus();
			field?.select();
		});
	}

	commitRenameSection(): void {
		const id = this.renamingSectionId();
		if (id === null) {
			return;
		}
		const name = this.renameValue().trim();
		this.renamingSectionId.set(null);
		if (name.length === 0) {
			return;
		}
		if (this.hostDriven()) {
			this.renameSection.emit({ sectionId: id, name });
		} else {
			this.editor?.sectionOps.rename(id, name);
		}
	}

	onRenameKeydown(event: KeyboardEvent): void {
		if (event.key === 'Enter') {
			event.preventDefault();
			this.commitRenameSection();
		} else if (event.key === 'Escape') {
			event.preventDefault();
			this.renamingSectionId.set(null);
		}
		event.stopPropagation();
	}

	sectionIndex(sectionId: string): number {
		return this.sectionList().findIndex((section) => section.id === sectionId);
	}

	/** Dictionary key for the word shown and announced on a hidden slide's card. */
	readonly hiddenLabelKey = HIDDEN_SLIDE_LABEL_KEY;

	/**
	 * The hidden-slide slash and dim, bound as inline styles rather than written
	 * into `slides-panel.component.css`. A component stylesheet cannot read a TS
	 * constant, so a literal copy there would be free to drift from the four
	 * other bindings; binding the shared values makes drift impossible.
	 */
	readonly slashGradient = HIDDEN_SLIDE_SLASH_GRADIENT;
	readonly dimOpacity = HIDDEN_SLIDE_DIM_OPACITY;

	/**
	 * The shared rail/sorter cue for one card. A hidden slide is still LISTED
	 * here (hiding only removes it from the show), so without this the panel gave
	 * a user no way to tell that a slide will be skipped.
	 */
	hiddenCue(slide: { hidden?: boolean }, index: number): HiddenSlideCue {
		return hiddenSlideCue(slide.hidden, 'rail', index);
	}
}
