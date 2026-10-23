import {
	applyReadOnlySessionEvent,
	INITIAL_READ_ONLY_SESSION,
	readOnlyBannerVisible,
	readOnlyLocked,
	readOnlyRecommendation,
	readOnlySessionActions,
} from 'ooxml-ui/pptx';
import type { ReadOnlyRecommendation, ReadOnlySession, ReadOnlySessionStore } from 'ooxml-ui/pptx';
import type { PptxModifyVerifier } from 'pptx-viewer-core';
import { useEffect, useMemo, useRef, useState } from 'react';

/**
 * useReadOnlyRecommendationState: whether the currently loaded deck asked to
 * be opened read-only (`p:modifyVerifier` or "Mark as Final"), and the
 * banner/lock state a viewer shows for it.
 *
 * When the recommendation's `requiresPassword` is set (a `modifyVerifier`
 * with a hash this viewer can actually check), "Edit anyway" no longer lifts
 * the lock immediately: it opens an inline password prompt, and only a
 * correct password (verified by `checkModifyPassword` against core's ECMA-376
 * hash check) unlocks. A wrong password leaves the deck read-only, matching
 * PowerPoint's own "read-only recommended" prompt.
 *
 * The session (banner hidden, Edit anyway, the prompt and its check) and its
 * transitions are the shared `read-only-session` model; this hook only keeps
 * it in React state.
 *
 * `setRecommendation`/`setModifyVerifier` are threaded into `useLoadContent`
 * exactly like every other per-load setter, so a fresh load always overwrites
 * the previous deck's recommendation. The session resets on the SAME trigger
 * `PowerPointViewer`'s Protected View override already resets on (`content`
 * changing), following that existing pattern rather than inventing a second
 * one.
 */
export function useReadOnlyRecommendationState(content: unknown) {
	const [recommendation, setRecommendation] = useState<ReadOnlyRecommendation>(() =>
		readOnlyRecommendation(undefined),
	);
	const [modifyVerifier, setModifyVerifier] = useState<PptxModifyVerifier | undefined>(undefined);
	const [session, setSession] = useState<ReadOnlySession>(INITIAL_READ_ONLY_SESSION);
	// A password check reads the session again when it finishes, after renders it cannot see,
	// so the store mirrors the latest value it set.
	const latest = useRef(session);
	const store = useMemo<ReadOnlySessionStore>(
		() => ({
			get: () => latest.current,
			set: (next) => {
				latest.current = next;
				setSession(next);
			},
		}),
		[],
	);
	const actions = useMemo(
		() =>
			readOnlySessionActions(store, {
				recommendation: () => recommendation,
				modifyVerifier: () => modifyVerifier,
			}),
		[store, recommendation, modifyVerifier],
	);

	useEffect(() => {
		applyReadOnlySessionEvent(store, { type: 'reset' });
		// `content` is deliberately the trigger (a new load), not a value read in
		// the body; mirrors PowerPointViewer's `protectedViewOverridden` reset
		// effect. `store` never changes.
		// oxlint-disable-next-line react/exhaustive-effect-dependencies -- see comment above
	}, [content, store]);

	return {
		recommendation,
		setRecommendation,
		setModifyVerifier,
		/** Whether the viewer's `canEdit` should be forced off. */
		locked: readOnlyLocked(recommendation, session),
		bannerVisible: readOnlyBannerVisible(recommendation, session),
		/** The banner's "Edit anyway": opens the password prompt when one is required, else unlocks. */
		editAnyway: actions.editAnyway,
		dismiss: actions.dismiss,
		/** Whether the inline password prompt should render in place of the banner's buttons. */
		passwordPromptOpen: session.passwordPromptOpen,
		/** Reason the last password attempt failed, or null before any attempt / after success. */
		passwordError: session.passwordError,
		/** True while `submitPassword`'s check is in flight (disables the form). */
		checkingPassword: session.checkingPassword,
		/** Check `password` against the deck's `modifyVerifier`; unlocks on a match. */
		submitPassword: actions.submitPassword,
		cancelPasswordPrompt: actions.cancelPassword,
	};
}

export type UseReadOnlyRecommendationStateResult = ReturnType<
	typeof useReadOnlyRecommendationState
>;
