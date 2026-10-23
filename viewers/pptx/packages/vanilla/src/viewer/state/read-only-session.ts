/**
 * The read-only banner's session (`read-only-session` in `ooxml-ui/pptx`) over the viewer store.
 * `ViewerState` keeps it in its public `readOnly*` fields; these helpers convert between the two
 * so the shared transitions drive the store.
 */
import { readOnlyLocked } from 'ooxml-ui/pptx';
import type { ReadOnlySession, ReadOnlySessionStore } from 'ooxml-ui/pptx';

import type { Store } from './store';
import type { ViewerState } from './viewer-state';

export type ReadOnlySessionFields = Pick<
	ViewerState,
	| 'readOnlyBannerDismissed'
	| 'readOnlyEditAnyway'
	| 'readOnlyPasswordPromptOpen'
	| 'readOnlyPasswordError'
	| 'readOnlyCheckingPassword'
	| 'readOnlyPasswordCheck'
>;

export function readOnlySessionOf(state: ReadOnlySessionFields): ReadOnlySession {
	return {
		bannerHidden: state.readOnlyBannerDismissed,
		editAnyway: state.readOnlyEditAnyway,
		passwordPromptOpen: state.readOnlyPasswordPromptOpen,
		passwordError: state.readOnlyPasswordError,
		checkingPassword: state.readOnlyCheckingPassword,
		passwordCheck: state.readOnlyPasswordCheck,
	};
}

export function readOnlySessionFields(session: ReadOnlySession): ReadOnlySessionFields {
	return {
		readOnlyBannerDismissed: session.bannerHidden,
		readOnlyEditAnyway: session.editAnyway,
		readOnlyPasswordPromptOpen: session.passwordPromptOpen,
		readOnlyPasswordError: session.passwordError,
		readOnlyCheckingPassword: session.checkingPassword,
		readOnlyPasswordCheck: session.passwordCheck,
	};
}

/** The session as a shared `ReadOnlySessionStore` backed by the viewer store. */
export function readOnlySessionStore(store: Store<ViewerState>): ReadOnlySessionStore {
	return {
		get: () => readOnlySessionOf(store.get()),
		set: (next) => store.set(readOnlySessionFields(next)),
	};
}

/** Whether the deck's read-only recommendation still locks editing ("Edit anyway" lifts it). */
export function isReadOnlyLocked(
	state: Pick<ViewerState, 'readOnlyRecommendation' | 'readOnlyEditAnyway'>,
): boolean {
	return readOnlyLocked(state.readOnlyRecommendation, { editAnyway: state.readOnlyEditAnyway });
}
