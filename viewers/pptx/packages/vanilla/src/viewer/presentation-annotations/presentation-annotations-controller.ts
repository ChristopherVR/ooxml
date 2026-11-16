import {
	annotationOverlayZIndex,
	applyPresentationInk,
	createPresentationInkController,
	presentationInkSlideCount,
} from 'ooxml-ui/pptx';
import type {
	CanvasSize,
	PresentationBlackout,
	PresentationInkPoint,
	PresentationInkStroke,
	PresentationPointerState,
} from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';

import type { Translator } from '../i18n';
import { mountAnnotationOverlay } from './annotation-overlay';
import { promptKeepAnnotations } from './keep-annotations-dialog';

export interface PresentationAnnotationsControllerOptions {
	doc: Document;
	t: Translator;
	getSlides(): PptxSlide[];
	/** Commit accepted ink with the viewer's normal history and dirty-state integration. */
	commitSlides(slides: PptxSlide[]): void;
	onStrokesChange?(strokes: PresentationInkStroke[]): void;
	onPointerMove?(point: PresentationInkPoint): void;
	/**
	 * File > Options > Advanced > "Prompt to keep ink annotations when
	 * exiting". Read fresh on every `finishPresentation()` call; defaults to
	 * `true` (always prompt) when omitted. When it returns `false`, the ink is
	 * discarded silently, matching real PowerPoint (no dialog).
	 */
	shouldPromptKeepAnnotations?(): boolean;
}

export interface PresentationAnnotationStage {
	stageWrap: HTMLElement;
	active: boolean;
	slideIndex: number;
	canvasSize: CanvasSize;
	pointer?: PresentationPointerState;
	/** Blackout state; the overlay is raised above the blank sheet while set. */
	blackout: PresentationBlackout;
}

export interface PresentationAnnotationsController {
	syncStage(stage: PresentationAnnotationStage): void;
	getStrokes(): readonly PresentationInkStroke[];
	setStrokes(strokes: readonly PresentationInkStroke[]): void;
	hasAnnotations(): boolean;
	/** Erase All Ink on Slide (E, and the show toolbar): the slide on stage only. */
	clearSlide(): void;
	/** Every slide's ink (leaving the show). */
	clear(): void;
	finishPresentation(): Promise<'none' | 'kept' | 'discarded'>;
	dispose(): void;
}

/**
 * Own temporary slide-show ink and optionally persist it into slide elements.
 * The strokes live in the shared `createPresentationInkController`; this
 * binding's overlay draws them in unit-square coordinates and reports whole
 * stroke lists, and "Keep" converts them through the shared `applyPresentationInk`.
 */
export function createPresentationAnnotationsController(
	options: PresentationAnnotationsControllerOptions,
): PresentationAnnotationsController {
	const ink = createPresentationInkController();
	const strokes = (): readonly PresentationInkStroke[] => ink.getSnapshot().strokes;
	let unmount = (): void => undefined;
	let lastStage: PresentationAnnotationStage | null = null;
	let lastOverlayKey: string | null = null;

	/**
	 * Everything about a stage that changes what the overlay element must be.
	 * The pointer's x/y are deliberately excluded: they change on every mouse
	 * move, and remounting on one destroyed the SVG mid-gesture (the drag reports
	 * its position, the report re-synced the stage, the re-sync replaced the
	 * element that was capturing the pointer), so a pen stroke could never finish
	 * and the show's Clear button stayed disabled for ever.
	 */
	const overlayKey = (stage: PresentationAnnotationStage): string =>
		[
			String(stage.active),
			String(stage.slideIndex),
			stage.pointer?.tool ?? 'none',
			stage.pointer?.color ?? '',
		].join('|');

	const notify = (): void => options.onStrokesChange?.([...strokes()]);
	const mount = (): void => {
		unmount();
		const stage = lastStage;
		const tool = stage?.pointer?.tool ?? 'none';
		if (!stage?.active || tool === 'none') {
			return;
		}
		unmount = mountAnnotationOverlay({
			stageWrap: stage.stageWrap,
			slideIndex: stage.slideIndex,
			tool,
			color: stage.pointer?.color ?? '#ef4444',
			blackout: stage.blackout,
			strokes: [...strokes()],
			onChange(next) {
				ink.setStrokes(next);
				notify();
			},
			onPointerMove: options.onPointerMove,
		});
	};
	const clear = (): void => {
		ink.clearAll();
		notify();
		mount();
	};
	const clearSlide = (): void => {
		ink.setSlide(lastStage?.slideIndex ?? 0);
		ink.clearSlide();
		notify();
		mount();
	};
	const persist = (): boolean => {
		const slides = options.getSlides();
		const next = applyPresentationInk(slides, strokes(), {
			width: lastStage?.canvasSize.width ?? 960,
			height: lastStage?.canvasSize.height ?? 540,
		});
		if (next !== slides) {
			options.commitSlides([...next]);
		}
		return next !== slides;
	};
	return {
		syncStage(stage) {
			const key = overlayKey(stage);
			// A stage rebuild empties `stageWrap`, so the overlay is also remounted
			// whenever it is no longer in the tree, not only when an input changed.
			const missing = stage.stageWrap.querySelector('.pptxv-presentation-annotations') === null;
			const changed = stage.stageWrap !== lastStage?.stageWrap || key !== lastOverlayKey;
			lastStage = stage;
			lastOverlayKey = key;
			if (changed || missing) {
				mount();
			}
			// The blackout state moves the overlay's stacking level (above the blank
			// sheet during a blackout, back down otherwise) without a remount, so a
			// stroke in progress is never cut off by a B/W toggle.
			const overlay = stage.stageWrap.querySelector<SVGSVGElement>(
				'.pptxv-presentation-annotations',
			);
			if (overlay) {
				overlay.style.zIndex = String(annotationOverlayZIndex(stage.blackout));
			}
		},
		getStrokes: strokes,
		setStrokes(next) {
			ink.setStrokes(next);
			notify();
			mount();
		},
		hasAnnotations: () => strokes().length > 0,
		clearSlide,
		clear,
		async finishPresentation() {
			if (strokes().length === 0) {
				return 'none';
			}
			if (options.shouldPromptKeepAnnotations?.() === false) {
				clear();
				return 'discarded';
			}
			const state = ink.getSnapshot();
			const choice = await promptKeepAnnotations(
				options.doc,
				options.t,
				state.strokes.length,
				presentationInkSlideCount(state),
			);
			if (choice === 'keep') {
				persist();
			}
			clear();
			return choice === 'keep' ? 'kept' : 'discarded';
		},
		dispose() {
			unmount();
			lastStage = null;
		},
	};
}
