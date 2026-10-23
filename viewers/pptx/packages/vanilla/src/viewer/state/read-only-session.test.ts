/**
 * read-only-session.test.ts: the shared read-only session over the viewer store. `ViewerState`
 * is public (a host may build its own state container against it), so a state written before the
 * session's `readOnlyEditAnyway` and `readOnlyPasswordCheck` fields existed must still work.
 */
import { reduceReadOnlySession, runReadOnlyPasswordCheck } from 'ooxml-ui/pptx';
import { describe, expect, it } from 'vitest';

import { isReadOnlyLocked, readOnlySessionStore } from './read-only-session';
import { createStore } from './store';
import { createInitialViewerState } from './viewer-state';
import type { ViewerState } from './viewer-state';

const PROTECTED = {
	kind: 'modifyVerifier',
	messageKey: 'pptx.readOnly.modifyVerifierRecommended',
	defaultReadOnly: true,
	requiresPassword: true,
} as const;

/** A host's own state from before the session fields: everything else, those two left out. */
function hostState(): ViewerState {
	const {
		readOnlyEditAnyway: _editAnyway,
		readOnlyPasswordCheck: _passwordCheck,
		...older
	} = createInitialViewerState();
	return { ...older, readOnlyRecommendation: PROTECTED } as ViewerState;
}

describe('read-only session over the viewer store', () => {
	it('unlocks with the right password on a host state without the session fields', async () => {
		const store = createStore(hostState());
		const session = readOnlySessionStore(store);
		expect(isReadOnlyLocked(store.get())).toBeTruthy();

		session.set(
			reduceReadOnlySession(session.get(), { type: 'editAnyway', requiresPassword: true }),
		);
		await runReadOnlyPasswordCheck(session, undefined, 'right-password', async () => ({
			ok: true,
		}));

		expect(isReadOnlyLocked(store.get())).toBeFalsy();
		expect(store.get().readOnlyPasswordPromptOpen).toBeFalsy();
	});
});
