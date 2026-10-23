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
} from 'ooxml-ui/pptx';
import type { PptxCustomProperty, PptxModifyVerifier } from 'pptx-viewer-core';
import { computed, shallowRef, watch } from 'vue';
import type { ComputedRef, Ref } from 'vue';

/**
 * useReadOnlyRecommendation: the Trust-Center-style banner shown when a
 * loaded deck asks to be opened read-only (`p:modifyVerifier`, or `docProps/
 * custom.xml`'s "Mark as Final"), and the lock that goes with it.
 *
 * Mirrors `PowerPointViewer.vue`'s existing Protected View mechanism
 * (`protectedViewActive`/`protectedViewDismissed`/`enableEditing`) rather than
 * inventing a second one: `locked` feeds the SAME `canEditEffective` gate
 * every edit entry point already reads, and `reset()` is called from the same
 * `watch(activeContent, ...)` handler that resets Protected View, so a newly
 * opened document is evaluated fresh even if a previous one was unlocked.
 *
 * The session (banner hidden, Edit anyway, the prompt and its check) and its
 * transitions are the shared `read-only-session` model; this keeps it in a ref.
 */
export interface UseReadOnlyRecommendationInput {
	modifyVerifier: Ref<PptxModifyVerifier | undefined>;
	customProperties: Ref<PptxCustomProperty[]>;
	/**
	 * Bumped when a load finishes applying a deck. The recommendation is read
	 * then, as PowerPoint reads it when the file opens, not again whenever the
	 * deck's custom properties are edited.
	 */
	loadVersion: Ref<number>;
}

export interface UseReadOnlyRecommendationResult {
	/** The current recommendation (`kind: null` when the deck asks for nothing). */
	recommendation: ComputedRef<ReadOnlyRecommendation>;
	/** Whether the banner should render. */
	showBanner: ComputedRef<boolean>;
	/** Whether editing should currently be blocked because of this recommendation. */
	locked: ComputedRef<boolean>;
	/**
	 * "Edit anyway": lifts the lock and hides the banner for this document, or
	 * (when `recommendation.requiresPassword` is set) opens the inline
	 * password prompt instead of unlocking immediately.
	 */
	editAnyway: () => void;
	/** "Dismiss": hides the banner but leaves the lock (if any) in place. */
	dismiss: () => void;
	/** Re-arm the recommendation for a newly loaded document. */
	reset: () => void;
	/** Whether the inline password prompt should render instead of the two buttons. */
	passwordPromptOpen: ComputedRef<boolean>;
	/** Reason the last password attempt failed, or null before any attempt / after success. */
	passwordError: ComputedRef<ReadOnlyBannerPasswordError | null>;
	/** True while `submitPassword`'s check is in flight (disables the form). */
	checkingPassword: ComputedRef<boolean>;
	/** Check `password` against the deck's `modifyVerifier`; unlocks on a match. */
	submitPassword: (password: string) => Promise<void>;
	/** Close the password prompt without unlocking. */
	cancelPasswordPrompt: () => void;
}

export function useReadOnlyRecommendation(
	input: UseReadOnlyRecommendationInput,
): UseReadOnlyRecommendationResult {
	const read = () =>
		readOnlyRecommendation({
			modifyVerifier: input.modifyVerifier.value,
			customProperties: input.customProperties.value,
		});
	const loaded = shallowRef<ReadOnlyRecommendation>(read());
	watch(
		input.loadVersion,
		() => {
			loaded.value = read();
		},
		{ flush: 'sync' },
	);
	const session = shallowRef<ReadOnlySession>(INITIAL_READ_ONLY_SESSION);
	const actions = readOnlySessionActions(
		{
			get: () => session.value,
			set: (next) => {
				session.value = next;
			},
		},
		{ recommendation: () => loaded.value, modifyVerifier: () => input.modifyVerifier.value },
	);

	return {
		recommendation: computed(() => loaded.value),
		showBanner: computed(() => readOnlyBannerVisible(loaded.value, session.value)),
		locked: computed(() => readOnlyLocked(loaded.value, session.value)),
		editAnyway: actions.editAnyway,
		dismiss: actions.dismiss,
		reset: actions.reset,
		passwordPromptOpen: computed(() => session.value.passwordPromptOpen),
		passwordError: computed(() => session.value.passwordError),
		checkingPassword: computed(() => session.value.checkingPassword),
		submitPassword: actions.submitPassword,
		cancelPasswordPrompt: actions.cancelPassword,
	};
}
