import { applyPresentationInk } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
/**
 * useAnnotationHandlers: Mode-switching callbacks that interact with
 * presentation annotations.  Extracted to keep the orchestrator lean.
 */
import { useCallback, useRef, useState } from 'react';

import type { ViewerMode } from '../types-core';
import { stopAllPersistentAudio } from '../utils/media';
import { mayLeaveSlideShow } from './presentation-mode/audience-content-store';
import type { EditorHistoryResult } from './useEditorHistory';
import type {
	SlideAnnotationMap,
	UsePresentationAnnotationsResult,
} from './usePresentationAnnotations';
import type { UsePresentationModeResult } from './usePresentationMode';

// ---------------------------------------------------------------------------
// Input
// ---------------------------------------------------------------------------

export interface UseAnnotationHandlersInput {
	mode: ViewerMode;
	presentation: UsePresentationModeResult;
	annotations: UsePresentationAnnotationsResult;
	history: EditorHistoryResult;
	setMode: React.Dispatch<React.SetStateAction<ViewerMode>>;
	setSlides: React.Dispatch<React.SetStateAction<PptxSlide[]>>;
	/**
	 * Options > Advanced > "Prompt to keep ink annotations when exiting"
	 * (default true). When false, exiting a slide show with annotations
	 * skips the dialog and keeps them silently.
	 */
	promptKeepInkAnnotations?: boolean;
}

// ---------------------------------------------------------------------------
// Output
// ---------------------------------------------------------------------------

export interface AnnotationHandlersResult {
	showKeepAnnotationsDialog: boolean;
	handleSetMode: (nextMode: ViewerMode) => void;
	handleKeepAnnotations: () => void;
	handleDiscardAnnotations: () => void;
	handleEnterPresenterView: () => void;
	handleEnterRehearsalMode: () => void;
}

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

export function useAnnotationHandlers(input: UseAnnotationHandlersInput): AnnotationHandlersResult {
	const {
		mode,
		presentation,
		annotations,
		history,
		setMode,
		setSlides,
		promptKeepInkAnnotations = true,
	} = input;

	const [showKeepAnnotationsDialog, setShowKeepAnnotationsDialog] = useState(false);
	const pendingModeAfterAnnotationsRef = useRef<ViewerMode>('edit');

	// ── Keep annotations as ink elements ──────────────────────────
	const handleKeepAnnotations = useCallback(() => {
		const annotationMap: SlideAnnotationMap = annotations.allSlideAnnotations;
		if (annotationMap.size === 0) {
			setShowKeepAnnotationsDialog(false);
			setMode(pendingModeAfterAnnotationsRef.current);
			annotations.clearAllAnnotations();
			return;
		}

		// Kept ink becomes `ink` elements on the slides it was drawn on (the
		// shared conversion every binding uses).
		setSlides((prev) => [...applyPresentationInk(prev, [...annotationMap.values()].flat())]);

		history.markDirty();
		annotations.clearAllAnnotations();
		setShowKeepAnnotationsDialog(false);
		setMode(pendingModeAfterAnnotationsRef.current);
	}, [annotations, setMode, setSlides, history]);

	// ── handleSetMode ─────────────────────────────────────────────
	const handleSetMode = useCallback(
		(nextMode: ViewerMode) => {
			// An audience display mirrors the presenter's screen. Every exit path
			// (Escape, the toolbar, the end-of-show advance) is refused there so the
			// editing chrome can never appear in front of the room.
			if (mode === 'present' && nextMode !== 'present' && !mayLeaveSlideShow()) {
				return;
			}
			if (nextMode === 'present') {
				presentation.enterPresentMode();
			} else if (mode === 'present' && annotations.hasAnyAnnotations) {
				stopAllPersistentAudio();
				if (promptKeepInkAnnotations) {
					pendingModeAfterAnnotationsRef.current = nextMode;
					setShowKeepAnnotationsDialog(true);
				} else {
					// Option off: discard the ink silently, no dialog (PowerPoint parity).
					annotations.clearAllAnnotations();
					setMode(nextMode);
				}
			} else {
				if (mode === 'present') {
					stopAllPersistentAudio();
				}
				setMode(nextMode);
			}
		},
		[presentation, mode, annotations, promptKeepInkAnnotations, setMode],
	);

	// ── Discard annotations ───────────────────────────────────────
	const handleDiscardAnnotations = useCallback(() => {
		annotations.clearAllAnnotations();
		setShowKeepAnnotationsDialog(false);
		setMode(pendingModeAfterAnnotationsRef.current);
	}, [annotations, setMode]);

	// ── Presenter view / rehearsal shortcuts ──────────────────────
	const handleEnterPresenterView = useCallback(() => {
		presentation.enterPresenterView();
	}, [presentation]);

	const handleEnterRehearsalMode = useCallback(() => {
		presentation.enterRehearsalMode();
	}, [presentation]);

	return {
		showKeepAnnotationsDialog,
		handleSetMode,
		handleKeepAnnotations,
		handleDiscardAnnotations,
		handleEnterPresenterView,
		handleEnterRehearsalMode,
	};
}
