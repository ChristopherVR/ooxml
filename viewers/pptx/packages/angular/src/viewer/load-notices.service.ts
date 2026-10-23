/**
 * load-notices.service.ts: Load-time viewer notices, Angular binding.
 *
 * Two independent decisions read straight off the loaded deck's `PptxData`:
 * - The read-only recommendation (`p:modifyVerifier` / "Mark as Final"):
 *   whether editing should default to locked, and the banner shown for it.
 *   Shared: `read-only-recommendation.ts` (`readOnlyRecommendation`).
 * - The compatibility-warning toast list: `data.warnings` plus every slide's
 *   own `warnings`, deduped by code. Shared: `compatibility-warning-toasts.ts`
 *   (`compatibilityWarningToasts`).
 *
 * Both reset on every load (see {@link resetForLoad}, called from the
 * viewer's load-reset effect) and both keep `PowerPointViewerComponent` from
 * growing for a per-load notice, mirroring `ViewerDialogsService`.
 *
 * @module viewer/load-notices
 */
import { Injectable, computed, inject, signal } from '@angular/core';
import {
	compatibilityWarningToasts,
	INITIAL_READ_ONLY_SESSION,
	readOnlyBannerVisible,
	readOnlyLocked,
	readOnlyRecommendation,
	readOnlySessionActions,
} from 'ooxml-ui/pptx';
import type {
	CompatibilityWarningToast,
	ReadOnlyRecommendation,
	ReadOnlySession,
} from 'ooxml-ui/pptx';

import { LoadContentService } from './load-content.service';

@Injectable()
export class LoadNoticesService {
	/** Optional so isolated service tests and stubbed injectors keep working. */
	private readonly loader = inject(LoadContentService, { optional: true });

	// ── Read-only recommendation ────────────────────────────────────────────

	/** Whether the loaded deck recommends opening read-only, and why. */
	readonly recommendation = computed<ReadOnlyRecommendation>(() =>
		readOnlyRecommendation(this.loader?.parsedData()),
	);
	/**
	 * What the user did about it this load: banner hidden, Edit anyway, the
	 * password prompt and its check (the shared `read-only-session` model).
	 */
	private readonly session = signal<ReadOnlySession>(INITIAL_READ_ONLY_SESSION);
	private readonly actions = readOnlySessionActions(
		{ get: () => this.session(), set: (next) => this.session.set(next) },
		{
			recommendation: () => this.recommendation(),
			modifyVerifier: () => this.loader?.parsedData()?.modifyVerifier ?? undefined,
		},
	);
	/** Whether the read-only banner should render. */
	readonly bannerActive = computed(() =>
		readOnlyBannerVisible(this.recommendation(), this.session()),
	);
	/**
	 * Whether the recommendation is still actively locking editing. Gates the
	 * host viewer's `canEdit`, the same way the Protected View lock does.
	 */
	readonly lockActive = computed(() => readOnlyLocked(this.recommendation(), this.session()));

	/** Whether the inline password prompt should render instead of the two buttons. */
	readonly passwordPromptOpen = computed(() => this.session().passwordPromptOpen);
	/** Reason the last password attempt failed, or null before any attempt / after success. */
	readonly passwordError = computed(() => this.session().passwordError);
	/** True while {@link submitPassword}'s check is in flight (disables the form). */
	readonly checkingPassword = computed(() => this.session().checkingPassword);

	/**
	 * "Edit anyway": lifts the recommendation's lock and hides the banner, or
	 * (when `recommendation().requiresPassword` is set) opens the inline
	 * password prompt instead of unlocking immediately.
	 */
	editAnyway(): void {
		this.actions.editAnyway();
	}

	/** "Dismiss": hides the banner but leaves any lock in place. */
	dismissBanner(): void {
		this.actions.dismiss();
	}

	/** Close the password prompt without unlocking. */
	cancelPasswordPrompt(): void {
		this.actions.cancelPassword();
	}

	/** Check `password` against the deck's `modifyVerifier`; unlocks on a match. */
	submitPassword(password: string): Promise<void> {
		return this.actions.submitPassword(password);
	}

	// ── Compatibility-warning toasts ────────────────────────────────────────

	/** Deck-level plus every slide's compatibility warnings, deduped by code. */
	readonly toasts = computed<CompatibilityWarningToast[]>(() => {
		const deckWarnings = this.loader?.parsedData()?.warnings ?? [];
		const slideWarnings = (this.loader?.slides() ?? []).flatMap((slide) => slide.warnings ?? []);
		return [...this.notices(), ...compatibilityWarningToasts([...deckWarnings, ...slideWarnings])];
	});
	/** Notices outside the load diagnostics (a failed Save); see {@link notify}. */
	private readonly notices = signal<CompatibilityWarningToast[]>([]);

	/** Show a notice (a failed Save) in the toast stack; shown again if it was dismissed. */
	notify(toast: CompatibilityWarningToast): void {
		this.notices.update((list) => [toast, ...list.filter((entry) => entry.id !== toast.id)]);
		this.dismissedToastIds.update((ids) => {
			const next = new Set(ids);
			next.delete(toast.id);
			return next;
		});
	}
	private readonly dismissedToastIds = signal<ReadonlySet<string>>(new Set());
	/** {@link toasts} with per-toast / dismiss-all dismissals applied. */
	readonly visibleToasts = computed(() =>
		this.toasts().filter((toast) => !this.dismissedToastIds().has(toast.id)),
	);

	/** Dismiss one toast by id. */
	dismissToast(id: string): void {
		this.dismissedToastIds.update((ids) => new Set(ids).add(id));
	}

	/** Dismiss every currently-visible toast. */
	dismissAllToasts(): void {
		this.dismissedToastIds.set(new Set(this.toasts().map((toast) => toast.id)));
	}

	// ── Reset ────────────────────────────────────────────────────────────────

	/** Reset both notices' dismissed/lifted state for a newly loaded deck. */
	resetForLoad(): void {
		this.actions.reset();
		this.dismissedToastIds.set(new Set());
		this.notices.set([]);
	}
}
