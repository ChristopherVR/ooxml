import {
	INITIAL_READ_ONLY_SESSION,
	readOnlyBannerVisible,
	readOnlyLocked,
	readOnlyRecommendation,
	readOnlySessionActions,
} from 'ooxml-ui/pptx';
import type {
	ReadOnlyBannerPasswordError,
	ReadOnlyRecommendation,
	ReadOnlySession,
	ReadOnlySessionActions,
} from 'ooxml-ui/pptx';
import type { PptxCustomProperty, PptxModifyVerifier } from 'pptx-viewer-core';

/**
 * ReadOnlyRecommendationState: the Trust-Center-style banner shown when a
 * loaded deck asks to be opened read-only (`p:modifyVerifier`, or `docProps/
 * custom.xml`'s "Mark as Final"), and the lock that goes with it. Svelte port
 * of Vue's `useReadOnlyRecommendation` composable.
 *
 * `locked` is meant to be ANDed into the SAME `getEditable` gate the viewer
 * already reads (Protected View, collaboration read-only), not a second
 * editability mechanism. `reset()` is called from the load-commit hook
 * (`onContentApplied`) so a newly opened document is evaluated fresh even if
 * a previous one was unlocked.
 *
 * The session (banner hidden, Edit anyway, the prompt and its check) and its
 * transitions are the shared `read-only-session` model; this keeps it in
 * `$state.raw` (each transition replaces the whole value).
 */
export class ReadOnlyRecommendationState {
	#session = $state.raw<ReadOnlySession>(INITIAL_READ_ONLY_SESSION);
	#actions: ReadOnlySessionActions;
	#getModifyVerifier: () => PptxModifyVerifier | undefined;
	#getCustomProperties: () => PptxCustomProperty[];

	constructor(deps: {
		getModifyVerifier(): PptxModifyVerifier | undefined;
		getCustomProperties(): PptxCustomProperty[];
	}) {
		this.#getModifyVerifier = deps.getModifyVerifier;
		this.#getCustomProperties = deps.getCustomProperties;
		this.#actions = readOnlySessionActions(
			{
				get: () => this.#session,
				set: (next) => {
					this.#session = next;
				},
			},
			{
				recommendation: () => this.recommendation,
				modifyVerifier: () => this.#getModifyVerifier(),
			},
		);
	}

	/** The current recommendation (`kind: null` when the deck asks for nothing). */
	get recommendation(): ReadOnlyRecommendation {
		return readOnlyRecommendation({
			modifyVerifier: this.#getModifyVerifier(),
			customProperties: this.#getCustomProperties(),
		});
	}

	/** Whether the banner should render. */
	get showBanner(): boolean {
		return readOnlyBannerVisible(this.recommendation, this.#session);
	}

	/** Whether editing should currently be blocked because of this recommendation. */
	get locked(): boolean {
		return readOnlyLocked(this.recommendation, this.#session);
	}

	/** Whether the inline password prompt should render instead of the two buttons. */
	get passwordPromptOpen(): boolean {
		return this.#session.passwordPromptOpen;
	}

	/** Reason the last password attempt failed, or null before any attempt / after success. */
	get passwordError(): ReadOnlyBannerPasswordError | null {
		return this.#session.passwordError;
	}

	/** True while {@link submitPassword}'s check is in flight (disables the form). */
	get checkingPassword(): boolean {
		return this.#session.checkingPassword;
	}

	/**
	 * "Edit anyway": lifts the lock and hides the banner for this document, or
	 * (when `recommendation.requiresPassword` is set) opens the inline
	 * password prompt instead of unlocking immediately.
	 */
	editAnyway(): void {
		this.#actions.editAnyway();
	}

	/** "Dismiss": hides the banner but leaves the lock (if any) in place. */
	dismiss(): void {
		this.#actions.dismiss();
	}

	/** Close the password prompt without unlocking. */
	cancelPasswordPrompt(): void {
		this.#actions.cancelPassword();
	}

	/** Check `password` against the deck's `modifyVerifier`; unlocks on a match. */
	submitPassword(password: string): Promise<void> {
		return this.#actions.submitPassword(password);
	}

	/** Re-arm the recommendation for a newly loaded document. */
	reset(): void {
		this.#actions.reset();
	}
}
