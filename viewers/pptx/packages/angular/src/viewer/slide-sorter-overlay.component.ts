import { NgStyle } from '@angular/common';
import {
	ChangeDetectionStrategy,
	CUSTOM_ELEMENTS_SCHEMA,
	DestroyRef,
	inject,
	Component,
	HostListener,
	computed,
	input,
	output,
} from '@angular/core';
import type { OnInit } from '@angular/core';
import { TranslatePipe, TranslateService } from '@ngx-translate/core';
import {
	createSlideSorterController,
	createSlideSorterState,
	HIDDEN_SLIDE_LABEL_KEY,
	HIDDEN_SLIDE_SLASH_GRADIENT,
	hiddenSlideCue,
	slideClipboardLabels,
	slideSorterContextMenuLabel,
	slideSorterView,
} from 'ooxml-ui/pptx';
import type {
	CanvasSize,
	HiddenSlideCue,
	SlideSorterContextMenuCommandId,
	SlideSorterContextMenuEntry,
	SlidePasteRequest,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';

import { controllerSignal } from './controller-signal';
import { SlideCanvasComponent } from './slide-canvas.component';
import { thumbnailHeight, thumbnailZoom } from './slide-sorter-overlay-helpers';

/** Pixel width of each thumbnail cell (the clipping box, not the canvas). */
const THUMB_W = 200;

/**
 * SlideSorterOverlayComponent: Angular port of the React `SlideSorterOverlay`.
 *
 * Renders a fixed full-screen modal overlay containing a responsive grid of
 * scaled slide previews. Every decision (selection, clipboard, zoom, keys, the
 * right-click menu and drag reorder) is the shared `createSlideSorterController`;
 * this component only draws it. Double-clicking a thumbnail emits `select(index)`;
 * dropping one on another emits `moveSlide`.
 *
 * Usage:
 * ```html
 * <pptx-slide-sorter-overlay
 *   [slides]="slides()"
 *   [canvasSize]="canvasSize()"
 *   [mediaDataUrls]="mediaDataUrls()"
 *   [activeIndex]="activeSlideIndex()"
 *   (select)="goTo($event)"
 *   (closed)="showSorter.set(false)"
 * />
 * ```
 */
@Component({
	selector: 'pptx-slide-sorter-overlay',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [NgStyle, SlideCanvasComponent, TranslatePipe],
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	templateUrl: './slide-sorter-overlay.component.html',
	styleUrl: './slide-sorter-overlay.component.css',
})
export class SlideSorterOverlayComponent implements OnInit {
	private readonly translation = inject(TranslateService);
	readonly copySystemSlides = input<((indexes: number[]) => Promise<void>) | undefined>();
	readonly pasteSystemSlides = input<((insertAt: number) => Promise<void>) | undefined>();
	clipboardLabels = () => slideClipboardLabels((key) => this.translation.instant(key));
	copyPackage = async (): Promise<void> => {
		await this.copySystemSlides()?.(this.controller.selectedIndexes());
	};
	pastePackage = async (): Promise<void> => {
		await this.pasteSystemSlides()?.(this.controller.pasteInsertAt());
	};
	/** Full list of slides to display. */
	readonly slides = input.required<PptxSlide[]>();
	readonly clipboardSourceSlides = input<readonly PptxSlide[]>();
	readonly clipboardTemplates = input<Record<string, PptxElement[]>>({});

	/** Natural (100 %) canvas dimensions, passed through to SlideCanvasComponent. */
	readonly canvasSize = input.required<CanvasSize>();

	/** Media asset lookup table, forwarded to each SlideCanvasComponent. */
	readonly mediaDataUrls = input<Map<string, string>>(new Map());

	/** Zero-based index of the currently active slide (highlighted in blue). */
	readonly activeIndex = input<number>(0);

	/** Whether the host allows edits; gates the deck-writing shortcuts. */
	readonly canEdit = input<boolean>(false);

	/** Emits the zero-based index of the thumbnail the user clicked. */
	readonly select = output<number>();

	/** Emits when the user closes the overlay (✕ button or Escape key). */
	readonly closed = output<void>();

	/** Delete the active slide (Delete / Backspace). */
	readonly deleteSlide = output<number>();

	/** Duplicate the active slide (Ctrl/Cmd+D). */
	readonly duplicateSlide = output<number>();
	readonly pasteSlides = output<SlidePasteRequest>();

	/** Toggle the hidden flag on a slide (context-menu only, no keyboard chord). */
	readonly toggleHiddenSlide = output<number>();

	/** Move one slide (drag a thumbnail onto another). */
	readonly moveSlide = output<{ from: number; to: number }>();

	// -------------------------------------------------------------------------
	// Derived display values
	// -------------------------------------------------------------------------

	/**
	 * The shared sorter state machine; the deck stays this component's inputs. Required inputs
	 * cannot be read at construction, so it starts empty and `ngOnInit` anchors it on the active
	 * slide.
	 */
	protected readonly controller = createSlideSorterController(
		{
			slides: () => this.slides(),
			activeIndex: () => this.activeIndex(),
			canEdit: () => this.canEdit(),
			clipboardSourceSlides: () => this.clipboardSourceSlides(),
			clipboardTemplates: () => this.clipboardTemplates(),
			onOperation: (operation, indexes) => {
				for (const index of indexes) {
					if (operation === 'delete') {
						this.deleteSlide.emit(index);
					} else if (operation === 'duplicate') {
						this.duplicateSlide.emit(index);
					} else {
						this.toggleHiddenSlide.emit(index);
					}
				}
			},
			onPaste: (request) => this.pasteSlides.emit(request),
			onMove: (from, to) => this.moveSlide.emit({ from, to }),
			onClose: () => this.closed.emit(),
		},
		{ initial: createSlideSorterState([], 0) },
	);
	private readonly snapshot = controllerSignal(this.controller, inject(DestroyRef));
	readonly view = computed(() => slideSorterView(this.snapshot(), this.slides()));
	/** Selection, clipboard and zoom (the shared `SlideSorterState`). */
	readonly state = computed(() => this.snapshot().sorter);
	readonly dragOverIndex = computed(() => this.snapshot().dragOverIndex);

	/** Zoom level that fits the full canvas width into THUMB_W pixels. */
	readonly thumbZoom = computed(() =>
		thumbnailZoom(this.canvasSize().width, (THUMB_W * this.state().zoom) / 100),
	);

	/** Pixel height of the clipping box (aspect-correct). */
	readonly thumbH = computed(() =>
		thumbnailHeight(
			this.canvasSize().width,
			this.canvasSize().height,
			(THUMB_W * this.state().zoom) / 100,
		),
	);

	/** ngStyle object for the thumbnail clipping box. */
	readonly clipStyle = computed<Record<string, string>>(() => ({
		width: `${(THUMB_W * this.state().zoom) / 100}px`,
		height: `${this.thumbH()}px`,
	}));

	/** ngStyle object for the grid: responsive auto-fill columns. */
	readonly gridStyle = computed<Record<string, string>>(() => ({
		'grid-template-columns': `repeat(${this.view().gridColumns}, minmax(0, 1fr))`,
	}));

	// -------------------------------------------------------------------------
	// Event handlers
	// -------------------------------------------------------------------------

	ngOnInit(): void {
		this.controller.reset();
	}

	@HostListener('document:keydown', ['$event'])
	onKeydown(event: KeyboardEvent): void {
		this.controller.handleKeydown(event);
	}
	setZoom(event: Event): void {
		this.controller.setZoom(Number((event.target as HTMLInputElement).value));
	}

	/** Clicking the backdrop (outside the panel) closes the overlay. */
	onBackdropClick(event: MouseEvent): void {
		// Only close when the click target IS the backdrop element itself.
		if (event.target === event.currentTarget) {
			this.closed.emit();
		}
	}

	/** Clicking selects within the sorter; double-clicking opens the canvas. */
	onThumbClick(index: number, event: MouseEvent): void {
		this.controller.select(index, event);
	}

	onDragStart(index: number, event: DragEvent): void {
		event.dataTransfer?.setData('text/plain', String(index));
		if (event.dataTransfer) {
			event.dataTransfer.effectAllowed = 'move';
		}
		this.controller.dragStart(index);
	}

	onDragOver(index: number, event: DragEvent): void {
		event.preventDefault();
		this.controller.dragOver(index);
	}

	onDrop(index: number, event: DragEvent): void {
		event.preventDefault();
		this.controller.drop(
			index,
			Number.parseInt(event.dataTransfer?.getData('text/plain') ?? '', 10),
		);
	}

	onDragEnd(): void {
		this.controller.dragEnd();
	}

	// -------------------------------------------------------------------------
	// Context menu (right-click a thumbnail)
	// -------------------------------------------------------------------------

	/** Open state + screen position of the context menu, or null when closed. */
	readonly contextMenu = computed(() => this.snapshot().menu);

	/**
	 * Right-clicking a thumbnail opens the menu for THAT slide.
	 *
	 * Deliberately does not also emit `select`: the host's `select` handler
	 * closes the whole overlay (it navigates the canvas and dismisses the
	 * sorter), so doing that here would tear down the menu before a single
	 * mouse action against it was reachable.
	 */
	onThumbContextMenu(event: MouseEvent, index: number): void {
		if (this.controller.openMenu(index, event.clientX, event.clientY)) {
			event.preventDefault();
		}
	}

	closeContextMenu(): void {
		this.controller.closeMenu();
	}

	readonly menuEntries = computed<SlideSorterContextMenuEntry[]>(() => this.view().menuEntries);
	menuLabelSuffix(translated: string, entry: SlideSorterContextMenuEntry): string {
		return slideSorterContextMenuLabel(translated, entry, this.view().selectedIndexes.length);
	}
	runMenuCommand(id: SlideSorterContextMenuCommandId): void {
		this.controller.run(id);
	}

	// -------------------------------------------------------------------------
	// Utilities
	// -------------------------------------------------------------------------

	/** Returns true when a slide has been marked as hidden in the presentation. */
	isHiddenSlide(slide: PptxSlide): boolean {
		// PptxSlide carries a `hidden` boolean when the slide is set to hidden in
		// the OpenXML package. Cast via unknown to avoid accessing a field that
		// may not exist on all versions of the core type.
		const s = slide as unknown as Record<string, unknown>;
		return s['hidden'] === true;
	}

	/** Dictionary key for the word shown and announced on a hidden slide's cell. */
	readonly hiddenLabelKey = HIDDEN_SLIDE_LABEL_KEY;

	/** Shared slash mark, bound inline so a stylesheet copy cannot drift. */
	readonly slashGradient = HIDDEN_SLIDE_SLASH_GRADIENT;

	/**
	 * The shared cue for one cell. The dim already came off `.is-hidden`, but
	 * opacity is a colour-only signal and said nothing to a screen reader, so
	 * this adds the number slash, the word, and the neutral marker attribute.
	 */
	hiddenCue(slide: PptxSlide, index: number): HiddenSlideCue {
		return hiddenSlideCue(this.isHiddenSlide(slide), 'sorter', index);
	}
}
