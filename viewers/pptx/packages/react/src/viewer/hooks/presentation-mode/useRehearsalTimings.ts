/**
 * Rehearse Timings for the React binding: the shared
 * `createPresentationRehearsalController` (clocks, pause, per-slide recording,
 * display tick) read through `useController`. The hook adds only what is
 * React's: entering the show (fullscreen inside the user gesture) and starting
 * the clock on the slide the show actually opens on.
 */
import { createPresentationRehearsalController } from 'ooxml-ui/pptx';
import type { PresentationRehearsalController } from 'ooxml-ui/pptx';
import { useCallback, useEffect, useRef, useState } from 'react';

import type { ViewerMode } from '../../types';
import { useController } from '../useController';

export interface UseRehearsalTimingsInput {
	containerRef: React.RefObject<HTMLElement | null>;
	onSetMode: (mode: ViewerMode) => void;
	onSaveRehearsalTimings?: (timings: Record<number, number>) => void;
	setPresentationStartTime: (time: number | null) => void;
	setPresenterMode: (mode: boolean) => void;
}

export interface UseRehearsalTimingsResult {
	rehearsing: boolean;
	recordedTimings: Readonly<Record<number, number>>;
	showRehearsalSummary: boolean;
	rehearsalPaused: boolean;
	/** Time on the current slide, paused time excluded (the controller ticks it). */
	slideElapsedMs: number;
	/** Time since the rehearsal started, paused time excluded. */
	totalElapsedMs: number;
	/** The show moved to `slideIndex` (records the slide being left). */
	moveRehearsal: (slideIndex: number) => void;
	/** The show ended: record the current slide and open the summary. */
	finishRehearsal: () => void;
	/** A plain show is starting: drop any rehearsal. */
	cancelRehearsal: () => void;
	/** The show opened on `slideIndex`: start the clock if a rehearsal was requested. */
	beginPendingRehearsal: (slideIndex: number) => void;
	dismissRehearsalSummary: () => void;
	saveRehearsalTimings: () => void;
	enterRehearsalMode: () => void;
	toggleRehearsalPause: () => void;
}

export function useRehearsalTimings(input: UseRehearsalTimingsInput): UseRehearsalTimingsResult {
	const {
		containerRef,
		onSetMode,
		onSaveRehearsalTimings,
		setPresentationStartTime,
		setPresenterMode,
	} = input;
	const [controller] = useState<PresentationRehearsalController>(() =>
		createPresentationRehearsalController(),
	);
	const state = useController(controller);
	const pendingRef = useRef(false);
	useEffect(() => () => controller.dispose(), [controller]);

	const saveRef = useRef(onSaveRehearsalTimings);
	saveRef.current = onSaveRehearsalTimings;

	const enterRehearsalMode = useCallback(() => {
		setPresenterMode(false);
		controller.discard();
		pendingRef.current = true;
		setPresentationStartTime(Date.now());
		try {
			const wrapper = containerRef.current;
			if (wrapper && typeof wrapper.requestFullscreen === 'function') {
				void wrapper.requestFullscreen().catch(() => {
					/* ignore */
				});
			}
		} catch {
			/* fullscreen not supported */
		}
		onSetMode('present');
	}, [containerRef, controller, onSetMode, setPresentationStartTime, setPresenterMode]);

	return {
		rehearsing: state.active,
		recordedTimings: state.timings,
		showRehearsalSummary: state.summaryOpen,
		rehearsalPaused: state.paused,
		slideElapsedMs: state.slideElapsedMs,
		totalElapsedMs: state.totalElapsedMs,
		moveRehearsal: controller.moveTo,
		finishRehearsal: controller.finish,
		cancelRehearsal: useCallback(() => {
			pendingRef.current = false;
			controller.discard();
		}, [controller]),
		beginPendingRehearsal: useCallback(
			(slideIndex: number) => {
				if (pendingRef.current) {
					pendingRef.current = false;
					controller.start(slideIndex);
				}
			},
			[controller],
		),
		dismissRehearsalSummary: controller.discard,
		saveRehearsalTimings: useCallback(
			() => saveRef.current?.({ ...controller.save() }),
			[controller],
		),
		enterRehearsalMode,
		toggleRehearsalPause: controller.togglePause,
	};
}
