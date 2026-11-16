/**
 * viewer-presentation-mode.service.ts: Viewer-scoped state + logic for the
 * fullscreen slideshow and presenter (speaker) view overlays: their visibility
 * flags, the presenter-view elapsed-timer start time, opening/closing a
 * separate audience window, mapping a (possibly custom-show-filtered)
 * presentation-overlay index back to the full-deck active slide, and the
 * keep/discard-annotations prompt shown when a slideshow with ink on it exits.
 *
 * Extracted from {@link PowerPointViewerComponent}: the component binds the few
 * accessors it alone owns via {@link bind}; the template reads the signals /
 * invokes the handlers off the injected instance directly.
 *
 * This service owns no DOM node, so the real browser Fullscreen API
 * request/exit is driven by `PresentationOverlayComponent` itself off this
 * `presenting` signal's mount/unmount, not from here.
 *
 * Provide it once on the viewer component (`providers: [ViewerPresentationModeService]`).
 */

import { computed, inject, Injectable, signal } from '@angular/core';
import { createPresentationRehearsalController } from 'ooxml-ui/pptx';
import type { AuthoredSlideRange } from 'ooxml-ui/pptx';

import { controllerSignal } from './controller-signal';
import type { SlideAnnotationMap } from './presentation-annotations-helpers';
import { endShowMediaCleanup } from './presentation-overlay-helpers';
import { PresenterWindowService } from './presenter-window.service';
import { ViewerCustomShowsService } from './viewer-custom-shows.service';

/** Live host accessors the presentation-mode controller needs. */
interface PresentationModeHost {
	readonly slideCount: () => number;
	readonly activeSlideIndex: () => number;
	readonly setActiveSlideIndex: (index: number) => void;
	readonly clearEditing: () => void;
	readonly clearSelection: () => void;
	readonly sourceContent: () => Uint8Array | ArrayBuffer | null;
	readonly canEdit: () => boolean;
	readonly promptKeepAnnotations: (map: SlideAnnotationMap) => void;
	readonly applyRehearsalTimings: (timings: Record<number, number>) => void;
	/** The deck's authored `p:showPr/p:sldRg` range, or null for the whole deck. */
	readonly authoredRange: () => AuthoredSlideRange | null;
}

export type RehearsalStart = 'beginning' | 'current';

/** Resolve the editor slide that a Record ribbon command should start from. */
export function resolveRehearsalStartIndex(
	start: RehearsalStart,
	currentIndex: number,
	slideCount: number,
): number {
	if (slideCount <= 0) {
		return 0;
	}
	return start === 'beginning' ? 0 : Math.min(Math.max(currentIndex, 0), slideCount - 1);
}

@Injectable()
export class ViewerPresentationModeService {
	private readonly presenterWindow = inject(PresenterWindowService);
	private readonly customShowsCtl = inject(ViewerCustomShowsService);

	/** Fullscreen presentation-mode overlay visibility. */
	readonly presenting = signal(false);
	/** Presenter-view (speaker) overlay visibility. */
	readonly presentingPresenter = signal(false);
	/** Epoch ms when presenter view started (drives the elapsed timer). */
	readonly presenterStartTime = signal<number | null>(null);
	/** Rehearse Timings: the shared controller, read as a signal. */
	readonly rehearsal = createPresentationRehearsalController();
	private readonly rehearsalState = controllerSignal(this.rehearsal);
	readonly rehearsing = computed(() => this.rehearsalState().active);
	readonly rehearsalPaused = computed(() => this.rehearsalState().paused);
	readonly showRehearsalSummary = computed(() => this.rehearsalState().summaryOpen);
	readonly rehearsalSlideElapsedMs = computed(() => this.rehearsalState().slideElapsedMs);
	readonly rehearsalTotalElapsedMs = computed(() => this.rehearsalState().totalElapsedMs);
	readonly recordedTimings = computed(() => this.rehearsalState().timings);
	/** Live subtitle preference shared by the ribbon and presentation overlay. */
	readonly subtitlesVisible = signal(false);

	private host: PresentationModeHost | null = null;

	/** Wire the host accessors (called once from the component constructor). */
	bind(host: PresentationModeHost): void {
		this.host = host;
	}

	private requireHost(): PresentationModeHost {
		if (!this.host) {
			throw new Error('ViewerPresentationModeService.bind() was not called');
		}
		return this.host;
	}

	/**
	 * Open the presentation overlay at `startIndex`, deselecting first so no
	 * edit chrome (selection outline / resize + rotate "Adjust shape" handles)
	 * leaks over the slideshow. The overlay itself
	 * (`PresentationOverlayComponent`) requests real browser fullscreen once it
	 * mounts as a result of `presenting` flipping true.
	 */
	private openShow(startIndex: number): void {
		const host = this.requireHost();
		if (host.slideCount() > 0) {
			host.clearSelection();
			host.clearEditing();
			// Seeded here, once, rather than derived by the overlay's `startIndex`
			// input: that input is live, and a value permanently pinned to a fixed
			// slide re-adopted itself over every advance the show made.
			host.setActiveSlideIndex(startIndex);
			this.presenting.set(true);
		}
	}

	/**
	 * Open the presentation overlay from the current slide ("From Current
	 * Slide", the status-bar "Slide show" button, `setMode('present')`). Opens
	 * on the active slide when the show (active custom show + the deck's
	 * authored `p:showPr/p:sldRg` range) includes it, else the closest show
	 * slide at or after it, else the show's own first slide.
	 */
	present(): void {
		const host = this.requireHost();
		this.openShow(this.customShowsCtl.showEntryIndex(host.authoredRange()));
	}

	/**
	 * "From Beginning" / F5: always opens the show's own first slide, ignoring
	 * where the editor is currently parked.
	 */
	presentFromBeginning(): void {
		const host = this.requireHost();
		this.openShow(this.customShowsCtl.showFirstIndex(host.authoredRange()));
	}

	presentFromCurrent(): void {
		this.present();
	}

	toggleSubtitles(): void {
		this.subtitlesVisible.update((visible) => !visible);
	}

	startRehearsalFromBeginning(): void {
		this.startRehearsal('beginning');
	}

	startRehearsalFromCurrent(): void {
		this.startRehearsal('current');
	}

	private startRehearsal(start: RehearsalStart): void {
		const host = this.requireHost();
		if (host.slideCount() <= 0) {
			return;
		}
		host.setActiveSlideIndex(
			resolveRehearsalStartIndex(start, host.activeSlideIndex(), host.slideCount()),
		);
		this.present();
		this.rehearsal.start(host.activeSlideIndex());
	}

	/**
	 * Adopt a presentation-overlay index as the editor's `activeSlideIndex`.
	 *
	 * No remap is needed (and none may be done): the overlay is handed the WHOLE
	 * live deck and applies custom-show membership as a navigation rule, so its
	 * index is already a deck index. The id-lookup this used to perform existed
	 * only because Angular pre-filtered the slide array to the show's members,
	 * and it resolved against the PRISTINE loaded deck, so it mislanded whenever
	 * the session had inserted or deleted a slide.
	 */
	onPresentationIndexChange(index: number): void {
		const host = this.requireHost();
		if (index < 0 || index >= this.customShowsCtl.presentationSlides().length) {
			return;
		}
		this.rehearsal.moveTo(index);
		host.setActiveSlideIndex(index);
	}

	toggleRehearsalPause(): void {
		this.rehearsal.togglePause();
	}

	closePresentation(): void {
		this.presenting.set(false);
		// The show has ended (this path never runs on a slide change), so its
		// cross-slide "play across slides" audio ends with it. The presenter-view
		// swap goes through togglePresenterView instead, which keeps it playing.
		endShowMediaCleanup();
		this.rehearsal.finish();
	}

	saveRehearsalTimings(): void {
		this.requireHost().applyRehearsalTimings({ ...this.rehearsal.save() });
	}

	dismissRehearsalSummary(): void {
		this.rehearsal.discard();
	}

	/**
	 * Open a separate audience tab and hand off the deck via the shared
	 * IndexedDB store. Mirrors React's presenter "open audience window".
	 */
	openAudienceWindow(): void {
		const host = this.requireHost();
		this.presenterWindow.openAudienceWindow(host.sourceContent(), host.activeSlideIndex());
	}

	/** Open the presenter (speaker) view: current+next slide, notes, timer. */
	presentPresenter(): void {
		const host = this.requireHost();
		if (host.slideCount() > 0) {
			this.presenterStartTime.set(Date.now());
			this.presentingPresenter.set(true);
		}
	}

	/**
	 * Swap between the fullscreen show and the presenter (speaker) console, the
	 * show toolbar's presenter-view toggle and PowerPoint's `N`. Mirrors React's
	 * `togglePresenterView`.
	 *
	 * The two are mutually exclusive rather than stacked: the show overlay is
	 * `position: fixed; z-index: 10000` while the console sits inside the viewer
	 * at `z-index: 50`, so leaving both up would paint the show straight over the
	 * console and the toggle would look inert. The full-deck `activeSlideIndex`
	 * is what both read, so the swap keeps the presenter on the same slide.
	 */
	togglePresenterView(): void {
		if (this.presentingPresenter()) {
			this.presentingPresenter.set(false);
			this.presenting.set(true);
			return;
		}
		if (this.presenterStartTime() === null) {
			this.presenterStartTime.set(Date.now());
		}
		this.presentingPresenter.set(true);
		this.presenting.set(false);
	}

	/** Close the presenter view (and any audience overlay/window it opened). */
	exitPresenter(): void {
		this.presentingPresenter.set(false);
		this.presenting.set(false);
		this.presenterWindow.closeAudienceWindow();
		// Leaving the presenter console ends the whole show, and with it any
		// cross-slide "play across slides" audio.
		endShowMediaCleanup();
	}

	/** Presentation exited with ink on it: offer the keep/discard prompt. */
	onPresentationAnnotationsExit(map: SlideAnnotationMap): void {
		const host = this.requireHost();
		if (host.canEdit()) {
			host.promptKeepAnnotations(map);
		}
	}
}
