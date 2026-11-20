/**
 * useActiveSlidePresence: publishes the local active slide to the session's
 * presence whatever the viewer is showing.
 *
 * Peers read it to follow along, and a broadcast's audience follows the
 * presenter by it. The cursor overlay also publishes the slide, but it is
 * mounted only on the editing canvas, so a presenter whose broadcast runs the
 * slide show never told the audience which slide it was on. Call this from a
 * component that stays mounted for the whole session.
 *
 * @module collaboration/useActiveSlidePresence
 */
import { useEffect } from 'react';

import type { CollaborationContextValue } from './types';

export function useActiveSlidePresence(
	collab: CollaborationContextValue | null,
	activeSlideIndex: number,
): void {
	const broadcastPresence = collab?.broadcastPresence;
	useEffect(() => {
		broadcastPresence?.({ activeSlideIndex });
	}, [broadcastPresence, activeSlideIndex]);
}
