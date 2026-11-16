/**
 * useAudienceMode: Detects when the viewer is running as an audience tab
 * (opened by the presenter via the audience window button) and auto-enters
 * fullscreen presentation mode, syncing slides via BroadcastChannel.
 */
import { createPresenterSessionController } from 'ooxml-ui/pptx';
import type { PresentationSnapshot } from 'ooxml-ui/pptx';
import { useEffect, useRef } from 'react';

import type { ViewerMode } from '../../types-core';
import { endAudienceDisplay } from './audience-content-store';
import { isAudienceTab } from './usePresenterWindow';

export interface UseAudienceModeInput {
	mode: ViewerMode;
	onSetMode: (mode: ViewerMode) => void;
	onSetActiveSlideIndex: (index: number) => void;
	containerRef: React.RefObject<HTMLElement | null>;
	onSnapshot?: (snapshot: PresentationSnapshot) => void;
	/**
	 * Raise the end-of-slide-show screen. Called when the presenter ends the
	 * session and the browser refuses to close this tab: the audience must see
	 * the black end screen, never the editor.
	 */
	onPresenterEnded?: () => void;
}

/**
 * When the page was opened with the `#pptx-audience` hash, this hook:
 * 1. Auto-enters fullscreen presentation mode after content loads
 * 2. Follows the presenter through the shared presenter session controller
 * 3. Raises the end screen when the presenter ends the session
 */
export function useAudienceMode(input: UseAudienceModeInput): void {
	const {
		mode: _mode,
		onSetMode,
		onSetActiveSlideIndex,
		containerRef,
		onSnapshot,
		onPresenterEnded,
	} = input;
	const initialised = useRef(false);

	// Auto-enter presentation mode when this is an audience tab
	useEffect(() => {
		if (!isAudienceTab()) {
			return;
		}
		if (initialised.current) {
			return;
		}
		initialised.current = true;

		onSetMode('present');

		// A newly navigated audience tab has no transient user activation, so an
		// automatic fullscreen request is rejected by browsers. Use the first
		// audience interaction as the activation and keep presentation mode active
		// meanwhile as a full-viewport fallback.
		const requestAudienceFullscreen = (): void => {
			const wrapper = containerRef.current;
			if (!document.fullscreenElement && wrapper?.requestFullscreen) {
				void wrapper.requestFullscreen().catch(() => undefined);
			}
		};
		document.addEventListener('pointerdown', requestAudienceFullscreen, { once: true });
		document.addEventListener('keydown', requestAudienceFullscreen, { once: true });
		return () => {
			document.removeEventListener('pointerdown', requestAudienceFullscreen);
			document.removeEventListener('keydown', requestAudienceFullscreen);
		};
	}, [containerRef, onSetMode]);

	// The channel must be opened ONCE. Re-subscribing whenever a caller passes a
	// fresh inline callback closed and reopened the BroadcastChannel on every
	// render, and a presenter message that landed in that gap was dropped.
	const handlersRef = useRef({ onSetActiveSlideIndex, onSnapshot, onPresenterEnded });
	handlersRef.current = { onSetActiveSlideIndex, onSnapshot, onPresenterEnded };

	// The audience half of the shared presenter session controller: adopt each
	// snapshot, follow slide changes and end the show when the presenter does.
	// The channel is opened ONCE; the handlers are read through the ref.
	useEffect(() => {
		if (!isAudienceTab()) {
			return;
		}
		const session = createPresenterSessionController({
			onAudienceSnapshot: (snapshot) => handlersRef.current.onSnapshot?.(snapshot),
			onAudienceSlide: (index) => handlersRef.current.onSetActiveSlideIndex(index),
			// Never fall back to edit mode: a tab the browser refuses to close
			// would otherwise show the full editing chrome to the room. Raise the
			// end-of-slide-show screen instead.
			onAudienceExit: () => {
				if (endAudienceDisplay(window)) {
					handlersRef.current.onPresenterEnded?.();
				}
			},
		});
		session.connect();
		return () => session.dispose();
	}, []);
}
