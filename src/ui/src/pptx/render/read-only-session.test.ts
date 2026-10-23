import { createModifyVerifier } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import type { ModifyPasswordCheckResult } from './modify-password-check';
import {
	dispatchReadOnlyBannerIntent,
	INITIAL_READ_ONLY_SESSION,
	readOnlyBannerState,
	readOnlyBannerVisible,
	readOnlyLocked,
	readOnlySessionActions,
	reduceReadOnlySession,
	runReadOnlyPasswordCheck,
} from './read-only-session';
import type {
	ReadOnlySession,
	ReadOnlySessionEvent,
	ReadOnlySessionStore,
} from './read-only-session';

const MARKED_FINAL = {
	kind: 'markedFinal',
	messageKey: 'pptx.readOnly.markedFinal',
	defaultReadOnly: true,
	requiresPassword: false,
} as const;
const PASSWORD = {
	kind: 'modifyVerifier',
	messageKey: 'pptx.readOnly.modifyVerifierRecommended',
	defaultReadOnly: true,
	requiresPassword: true,
} as const;

const run = (...events: ReadOnlySessionEvent[]): ReadOnlySession =>
	events.reduce(reduceReadOnlySession, INITIAL_READ_ONLY_SESSION);

function memoryStore(initial: ReadOnlySession = INITIAL_READ_ONLY_SESSION): ReadOnlySessionStore {
	let session = initial;
	return { get: () => session, set: (next) => (session = next) };
}

/** A check the test finishes by hand. */
function deferredCheck() {
	const resolvers: Array<(result: ModifyPasswordCheckResult) => void> = [];
	const check = vi.fn(
		() => new Promise<ModifyPasswordCheckResult>((resolve) => resolvers.push(resolve)),
	);
	return {
		check,
		finish: (index: number, result: ModifyPasswordCheckResult) => resolvers[index]!(result),
	};
}

describe('reduceReadOnlySession', () => {
	it('Edit anyway unlocks and hides the banner when no password is required', () => {
		const session = run({ type: 'editAnyway', requiresPassword: false });
		expect(readOnlyLocked(MARKED_FINAL, session)).toBeFalsy();
		expect(readOnlyBannerVisible(MARKED_FINAL, session)).toBeFalsy();
	});

	it('Edit anyway opens the password prompt, clearing an old error, when one is required', () => {
		const failed = run(
			{ type: 'editAnyway', requiresPassword: true },
			{ type: 'checkStarted' },
			{ type: 'checkFinished', check: 1, result: { ok: false, reason: 'wrong-password' } },
			{ type: 'cancelPassword' },
			{ type: 'editAnyway', requiresPassword: true },
		);
		expect(failed).toMatchObject({ passwordPromptOpen: true, passwordError: null });
		expect(readOnlyLocked(PASSWORD, failed)).toBeTruthy();
		expect(readOnlyBannerVisible(PASSWORD, failed)).toBeTruthy();
	});

	it('Dismiss hides the banner and keeps the lock', () => {
		const session = run({ type: 'dismiss' });
		expect(readOnlyBannerVisible(MARKED_FINAL, session)).toBeFalsy();
		expect(readOnlyLocked(MARKED_FINAL, session)).toBeTruthy();
		expect(reduceReadOnlySession(session, { type: 'dismiss' })).toBe(session);
	});

	it('Cancel closes the prompt and forgets the error', () => {
		const session = run(
			{ type: 'editAnyway', requiresPassword: true },
			{ type: 'checkStarted' },
			{ type: 'checkFinished', check: 1, result: { ok: false, reason: 'unsupported-algorithm' } },
			{ type: 'cancelPassword' },
		);
		expect(session).toMatchObject({
			passwordPromptOpen: false,
			passwordError: null,
			checkingPassword: false,
		});
		expect(readOnlyLocked(PASSWORD, session)).toBeTruthy();
	});

	it('a check unlocks on ok and records the reason otherwise', () => {
		const open = run({ type: 'editAnyway', requiresPassword: true }, { type: 'checkStarted' });
		expect(open.checkingPassword).toBeTruthy();
		const ok = reduceReadOnlySession(open, {
			type: 'checkFinished',
			check: open.passwordCheck,
			result: { ok: true },
		});
		expect(ok).toMatchObject({ editAnyway: true, bannerHidden: true, passwordPromptOpen: false });
		expect(ok.checkingPassword).toBeFalsy();
		const wrong = reduceReadOnlySession(open, {
			type: 'checkFinished',
			check: open.passwordCheck,
			result: { ok: false, reason: 'wrong-password' },
		});
		expect(wrong).toMatchObject({
			editAnyway: false,
			passwordPromptOpen: true,
			passwordError: 'wrong-password',
			checkingPassword: false,
		});
	});

	it('a second Unlock while a check runs changes nothing', () => {
		const checking = run({ type: 'editAnyway', requiresPassword: true }, { type: 'checkStarted' });
		expect(reduceReadOnlySession(checking, { type: 'checkStarted' })).toBe(checking);
	});

	it('drops a result that arrives after a new load or Cancel', () => {
		const checking = run({ type: 'editAnyway', requiresPassword: true }, { type: 'checkStarted' });
		const late = {
			type: 'checkFinished',
			check: checking.passwordCheck,
			result: { ok: true },
		} as const;
		for (const next of [{ type: 'reset' }, { type: 'cancelPassword' }] as const) {
			const after = reduceReadOnlySession(checking, next);
			expect(reduceReadOnlySession(after, late)).toBe(after);
			expect(readOnlyLocked(PASSWORD, after)).toBeTruthy();
		}
	});

	it('reset starts over but never reuses a check id', () => {
		const unlocked = run({ type: 'editAnyway', requiresPassword: false });
		const reset = reduceReadOnlySession(unlocked, { type: 'reset' });
		expect(reset).toStrictEqual({ ...INITIAL_READ_ONLY_SESSION, passwordCheck: 1 });
		const checking = run({ type: 'checkStarted' }, { type: 'reset' }, { type: 'checkStarted' });
		expect(checking.passwordCheck).toBe(3);
	});
});

describe('derived flags', () => {
	it('lock and show nothing without a recommendation', () => {
		for (const recommendation of [null, undefined, { kind: null, defaultReadOnly: false }]) {
			expect(readOnlyLocked(recommendation, INITIAL_READ_ONLY_SESSION)).toBeFalsy();
			expect(readOnlyBannerVisible(recommendation, INITIAL_READ_ONLY_SESSION)).toBeFalsy();
		}
	});

	it('builds the banner state, with the translator only when given', () => {
		const session = run({ type: 'editAnyway', requiresPassword: true }, { type: 'checkStarted' });
		const translate = (key: string) => `t:${key}`;
		expect(readOnlyBannerState(PASSWORD, session, translate)).toStrictEqual({
			kind: 'modifyVerifier',
			messageKey: 'pptx.readOnly.modifyVerifierRecommended',
			passwordPromptOpen: true,
			passwordError: null,
			checkingPassword: true,
			translate,
		});
		expect(readOnlyBannerState(MARKED_FINAL, INITIAL_READ_ONLY_SESSION)).not.toHaveProperty(
			'translate',
		);
	});
});

describe('runReadOnlyPasswordCheck', () => {
	const open = () => memoryStore(run({ type: 'editAnyway', requiresPassword: true }));

	it('checks the real verifier by default', async () => {
		const verifier = await createModifyVerifier('right-password', { spinCount: 10 });
		const wrong = open();
		await runReadOnlyPasswordCheck(wrong, verifier, 'nope');
		expect(wrong.get()).toMatchObject({ passwordError: 'wrong-password', checkingPassword: false });
		const right = open();
		await runReadOnlyPasswordCheck(right, verifier, 'right-password');
		expect(readOnlyLocked(PASSWORD, right.get())).toBeFalsy();
	});

	it('is checking while the check runs and ignores a second submit', async () => {
		const store = open();
		const { check, finish } = deferredCheck();
		const first = runReadOnlyPasswordCheck(store, undefined, 'a', check);
		expect(store.get().checkingPassword).toBeTruthy();
		await runReadOnlyPasswordCheck(store, undefined, 'b', check);
		expect(check).toHaveBeenCalledTimes(1);
		finish(0, { ok: true });
		await first;
		expect(readOnlyLocked(PASSWORD, store.get())).toBeFalsy();
	});

	it('drops a result that arrives after a new load or Cancel', async () => {
		for (const next of [{ type: 'reset' }, { type: 'cancelPassword' }] as const) {
			const store = open();
			const { check, finish } = deferredCheck();
			const pending = runReadOnlyPasswordCheck(store, undefined, 'a', check);
			store.set(reduceReadOnlySession(store.get(), next));
			const after = store.get();
			finish(0, { ok: true });
			await pending;
			expect(store.get()).toBe(after);
			expect(readOnlyLocked(PASSWORD, after)).toBeTruthy();
		}
	});

	it('treats a check that throws as one this viewer cannot run', async () => {
		const store = open();
		await runReadOnlyPasswordCheck(store, undefined, 'a', () => Promise.reject(new Error('x')));
		expect(store.get()).toMatchObject({
			passwordError: 'unsupported-algorithm',
			checkingPassword: false,
		});
	});
});

describe('readOnlySessionActions', () => {
	it('reads the recommendation and verifier when each action runs', async () => {
		const store = memoryStore();
		let recommendation: { requiresPassword: boolean } | null = null;
		const verifier = await createModifyVerifier('right-password', { spinCount: 10 });
		const actions = readOnlySessionActions(store, {
			recommendation: () => recommendation,
			modifyVerifier: () => verifier,
		});
		recommendation = PASSWORD;
		actions.editAnyway();
		expect(store.get().passwordPromptOpen).toBeTruthy();
		await actions.submitPassword('right-password');
		expect(store.get().editAnyway).toBeTruthy();
		actions.reset();
		expect(store.get().editAnyway).toBeFalsy();
		actions.dismiss();
		expect(store.get().bannerHidden).toBeTruthy();
		actions.editAnyway();
		actions.cancelPassword();
		expect(store.get().passwordPromptOpen).toBeFalsy();
	});
});

describe('dispatchReadOnlyBannerIntent', () => {
	const request = (detail: unknown) => new CustomEvent('read-only-request', { detail });

	it('routes each request to its handler', () => {
		const seen: string[] = [];
		const handlers = {
			editAnyway: () => seen.push('editAnyway'),
			dismiss: () => seen.push('dismiss'),
			cancelPassword: () => seen.push('cancelPassword'),
			submitPassword: (password: string) => seen.push(`submit:${password}`),
		};
		dispatchReadOnlyBannerIntent(request({ id: 'editAnyway' }), handlers);
		dispatchReadOnlyBannerIntent(request({ id: 'dismiss' }), handlers);
		dispatchReadOnlyBannerIntent(request({ id: 'cancelPassword' }), handlers);
		dispatchReadOnlyBannerIntent(request({ id: 'submitPassword', password: 'pw' }), handlers);
		expect(seen).toStrictEqual(['editAnyway', 'dismiss', 'cancelPassword', 'submit:pw']);
	});

	it('does nothing for a request without a handler or without detail', () => {
		expect(() => {
			dispatchReadOnlyBannerIntent(request({ id: 'submitPassword', password: 'pw' }), {});
			dispatchReadOnlyBannerIntent(new Event('read-only-request'), { dismiss: undefined });
		}).not.toThrow();
	});
});
