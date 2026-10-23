/**
 * Shared session model for the read-only banner (`pptx-ui-read-only-banner`).
 *
 * `readOnlyRecommendation` says whether a loaded deck asks to be opened read-only. What the user
 * does about it during the session (hide the banner, lift the lock with Edit anyway or the right
 * password, open and cancel the password prompt) was written five times. The session is a plain
 * value each host keeps in its own reactive container: {@link reduceReadOnlySession} is the only
 * place it changes, {@link readOnlyLocked} and {@link readOnlyBannerVisible} derive the edit gate
 * and the banner from it, {@link runReadOnlyPasswordCheck} runs Unlock through the transitions and
 * {@link dispatchReadOnlyBannerIntent} routes the element's requests.
 *
 * DOM-free and framework-free.
 */
import type { PptxModifyVerifier } from 'ooxml-core/pptx';

import type {
	ChromeTranslate,
	ReadOnlyBannerIntent,
	ReadOnlyBannerPasswordError,
	ReadOnlyBannerViewState,
} from './chrome-controls-state';
import { checkModifyPassword } from './modify-password-check';
import type { ModifyPasswordCheckResult } from './modify-password-check';
import type { ReadOnlyRecommendation } from './read-only-recommendation';

/** What the user has done about the current deck's read-only recommendation. */
export interface ReadOnlySession {
	/** The banner is closed: Dismiss, Edit anyway or the right password. */
	readonly bannerHidden: boolean;
	/** Edit anyway (or the right password) lifted the lock. Dismiss never does. */
	readonly editAnyway: boolean;
	/** The inline password form replaces Edit anyway and Dismiss. */
	readonly passwordPromptOpen: boolean;
	/** Why the last check failed; null before one and after Cancel, a reopen or success. */
	readonly passwordError: ReadOnlyBannerPasswordError | null;
	/** A submitted password is being checked. */
	readonly checkingPassword: boolean;
	/**
	 * Which check a result belongs to. A new load, Cancel and every new check move it on, so a
	 * check that finishes after one of them is dropped instead of unlocking the wrong deck.
	 */
	readonly passwordCheck: number;
}

/** A session before the user has done anything. */
export const INITIAL_READ_ONLY_SESSION: ReadOnlySession = Object.freeze({
	bannerHidden: false,
	editAnyway: false,
	passwordPromptOpen: false,
	passwordError: null,
	checkingPassword: false,
	passwordCheck: 0,
});

export type ReadOnlySessionEvent =
	/** A new deck: everything starts over and a check still running is dropped. */
	| { readonly type: 'reset' }
	/** Edit anyway: the password prompt when the deck requires one, otherwise unlock. */
	| { readonly type: 'editAnyway'; readonly requiresPassword: boolean }
	/** Dismiss: hide the banner, keep the lock. */
	| { readonly type: 'dismiss' }
	/** The prompt's Cancel: close it, forget the error, drop a check still running. */
	| { readonly type: 'cancelPassword' }
	/** Unlock: start a check. Ignored while one is already running. */
	| { readonly type: 'checkStarted' }
	/** A check finished; applied only while it is still the current one. */
	| {
			readonly type: 'checkFinished';
			readonly check: number;
			readonly result: ModifyPasswordCheckResult;
	  };

function unlocked(session: ReadOnlySession): ReadOnlySession {
	return {
		...session,
		bannerHidden: true,
		editAnyway: true,
		passwordPromptOpen: false,
		passwordError: null,
		checkingPassword: false,
	};
}

/** The one transition function. Returns `session` itself when the event changes nothing. */
export function reduceReadOnlySession(
	session: ReadOnlySession,
	event: ReadOnlySessionEvent,
): ReadOnlySession {
	switch (event.type) {
		case 'reset':
			return { ...INITIAL_READ_ONLY_SESSION, passwordCheck: session.passwordCheck + 1 };
		case 'editAnyway':
			return event.requiresPassword
				? { ...session, passwordPromptOpen: true, passwordError: null }
				: unlocked(session);
		case 'dismiss':
			return session.bannerHidden ? session : { ...session, bannerHidden: true };
		case 'cancelPassword':
			return {
				...session,
				passwordPromptOpen: false,
				passwordError: null,
				checkingPassword: false,
				passwordCheck: session.passwordCheck + 1,
			};
		case 'checkStarted':
			return session.checkingPassword
				? session
				: {
						...session,
						checkingPassword: true,
						passwordCheck: session.passwordCheck + 1,
					};
		case 'checkFinished':
			if (!session.checkingPassword || event.check !== session.passwordCheck) {
				return session;
			}
			return event.result.ok
				? unlocked(session)
				: { ...session, checkingPassword: false, passwordError: event.result.reason };
	}
}

/** The recommendation fields the derived flags read; null or undefined means none. */
export type ReadOnlySessionRecommendation =
	| Pick<ReadOnlyRecommendation, 'kind' | 'defaultReadOnly'>
	| null
	| undefined;

/** Whether editing is locked: the deck recommends read-only and nothing lifted it. */
export function readOnlyLocked(
	recommendation: ReadOnlySessionRecommendation,
	session: Pick<ReadOnlySession, 'editAnyway'>,
): boolean {
	return recommendation?.defaultReadOnly === true && !session.editAnyway;
}

/** Whether the banner shows: the deck recommends read-only and the banner is still open. */
export function readOnlyBannerVisible(
	recommendation: ReadOnlySessionRecommendation,
	session: Pick<ReadOnlySession, 'bannerHidden'>,
): boolean {
	return (recommendation?.kind ?? null) !== null && !session.bannerHidden;
}

/** The banner element's state for a recommendation and the session's prompt. */
export function readOnlyBannerState(
	recommendation: Pick<ReadOnlyRecommendation, 'kind' | 'messageKey'>,
	session: Pick<ReadOnlySession, 'passwordPromptOpen' | 'passwordError' | 'checkingPassword'>,
	translate?: ChromeTranslate,
): ReadOnlyBannerViewState {
	return {
		kind: recommendation.kind ?? null,
		messageKey: recommendation.messageKey,
		passwordPromptOpen: session.passwordPromptOpen,
		passwordError: session.passwordError,
		checkingPassword: session.checkingPassword,
		...(translate ? { translate } : {}),
	};
}

/** Where a host keeps its session. `get` must return what the last `set` stored. */
export interface ReadOnlySessionStore {
	get(): ReadOnlySession;
	set(next: ReadOnlySession): void;
}

/** Apply one event to a store; a no-op event leaves the store untouched. */
export function applyReadOnlySessionEvent(
	store: ReadOnlySessionStore,
	event: ReadOnlySessionEvent,
): void {
	const current = store.get();
	const next = reduceReadOnlySession(current, event);
	if (next !== current) {
		store.set(next);
	}
}

export type ModifyPasswordCheck = (
	verifier: PptxModifyVerifier | undefined,
	password: string,
) => Promise<ModifyPasswordCheckResult>;

/**
 * Unlock: check `password` against the deck's verifier through the transitions. A second submit
 * while a check runs does nothing, and a result that arrives after a new load or Cancel is
 * dropped. `check` defaults to `checkModifyPassword`; a check that throws counts as one this
 * viewer cannot run.
 */
export async function runReadOnlyPasswordCheck(
	store: ReadOnlySessionStore,
	verifier: PptxModifyVerifier | undefined,
	password: string,
	check: ModifyPasswordCheck = checkModifyPassword,
): Promise<void> {
	const before = store.get();
	const started = reduceReadOnlySession(before, { type: 'checkStarted' });
	if (started === before) {
		return;
	}
	store.set(started);
	let result: ModifyPasswordCheckResult;
	try {
		result = await check(verifier, password);
	} catch {
		result = { ok: false, reason: 'unsupported-algorithm' };
	}
	applyReadOnlySessionEvent(store, {
		type: 'checkFinished',
		check: started.passwordCheck,
		result,
	});
}

/** What the actions read from the host when they run. */
export interface ReadOnlySessionSource {
	/** The loaded deck's recommendation; its `requiresPassword` decides what Edit anyway does. */
	recommendation(): Pick<ReadOnlyRecommendation, 'requiresPassword'> | null | undefined;
	/** The loaded deck's `p:modifyVerifier`, which Unlock checks. */
	modifyVerifier(): PptxModifyVerifier | undefined;
}

/** A host's banner actions; also usable as {@link ReadOnlyBannerIntentHandlers}. */
export interface ReadOnlySessionActions {
	/** A new deck loaded. */
	reset(): void;
	editAnyway(): void;
	dismiss(): void;
	cancelPassword(): void;
	submitPassword(password: string): Promise<void>;
}

/** The banner's actions over a host's store, each one a transition. */
export function readOnlySessionActions(
	store: ReadOnlySessionStore,
	source: ReadOnlySessionSource,
): ReadOnlySessionActions {
	const apply = (event: ReadOnlySessionEvent) => applyReadOnlySessionEvent(store, event);
	return {
		reset: () => apply({ type: 'reset' }),
		editAnyway: () =>
			apply({
				type: 'editAnyway',
				requiresPassword: source.recommendation()?.requiresPassword === true,
			}),
		dismiss: () => apply({ type: 'dismiss' }),
		cancelPassword: () => apply({ type: 'cancelPassword' }),
		submitPassword: (password) =>
			runReadOnlyPasswordCheck(store, source.modifyVerifier(), password),
	};
}

/** A host's handler per banner request; a request without one does nothing. */
export interface ReadOnlyBannerIntentHandlers {
	editAnyway?: (() => void) | undefined;
	dismiss?: (() => void) | undefined;
	cancelPassword?: (() => void) | undefined;
	submitPassword?: ((password: string) => void) | undefined;
}

/** Route the banner's `read-only-request` event to the host's handler for it. */
export function dispatchReadOnlyBannerIntent(
	event: Event,
	handlers: ReadOnlyBannerIntentHandlers,
): void {
	const intent = (event as CustomEvent<ReadOnlyBannerIntent | undefined>).detail;
	switch (intent?.id) {
		case 'editAnyway':
			handlers.editAnyway?.();
			break;
		case 'dismiss':
			handlers.dismiss?.();
			break;
		case 'cancelPassword':
			handlers.cancelPassword?.();
			break;
		case 'submitPassword':
			handlers.submitPassword?.(intent.password);
			break;
	}
}
