import {
	dispatchReadOnlyBannerIntent,
	readOnlyBannerState,
	readOnlyBannerVisible,
	registerPptxWebControls,
} from 'ooxml-ui/pptx';
import type {
	PptxUiReadOnlyBannerElement,
	ReadOnlyBannerPasswordError,
	ReadOnlyRecommendation,
} from 'ooxml-ui/pptx';

import type { Translator } from '../i18n';

/** Why the last password attempt failed; see `checkModifyPassword` (`ooxml-ui/pptx`). */
export type ModifyPasswordErrorReason = ReadOnlyBannerPasswordError;

/**
 * The `p:modifyVerifier` / "Mark as Final" read-only recommendation banner:
 * a strip below the ribbon telling the user WHY the deck opened locked and
 * offering "Edit anyway" (lifts the lock, hides the banner) or a plain close
 * (hides the banner, keeps the lock). See `readOnlyRecommendation` in
 * `pptx-viewer-shared` and `ViewerState.readOnlyRecommendation`.
 *
 * A thin adapter over the shared `pptx-ui-read-only-banner`, which owns the
 * markup and the inline password form. When the recommendation's
 * `requiresPassword` is set (a `modifyVerifier` with a hash this viewer can
 * check), "Edit anyway" is replaced by that form instead: PowerPoint's own
 * "read-only recommended" prompt keeps the deck locked until the correct
 * password is entered, and a wrong one leaves it locked.
 */
export interface ReadOnlyBanner {
	el: HTMLElement;
	update(
		recommendation: ReadOnlyRecommendation | null,
		dismissed: boolean,
		passwordState: {
			promptOpen: boolean;
			error: ModifyPasswordErrorReason | null;
			checking: boolean;
		},
	): void;
}

export function createReadOnlyBanner(
	doc: Document,
	t: Translator,
	onEditAnyway: () => void,
	onDismiss: () => void,
	onSubmitPassword: (password: string) => void,
	onCancelPassword: () => void,
): ReadOnlyBanner {
	registerPptxWebControls();
	const el = doc.createElement('pptx-ui-read-only-banner') as PptxUiReadOnlyBannerElement;
	el.hidden = true;
	el.addEventListener('read-only-request', (event) =>
		dispatchReadOnlyBannerIntent(event, {
			editAnyway: onEditAnyway,
			dismiss: onDismiss,
			cancelPassword: onCancelPassword,
			submitPassword: onSubmitPassword,
		}),
	);
	return {
		el,
		update(recommendation, dismissed, passwordState) {
			el.hidden = !readOnlyBannerVisible(recommendation, { bannerHidden: dismissed });
			if (el.hidden || !recommendation) {
				return;
			}
			el.state = readOnlyBannerState(
				recommendation,
				{
					passwordPromptOpen: passwordState.promptOpen,
					passwordError: passwordState.error,
					checkingPassword: passwordState.checking,
				},
				t,
			);
		},
	};
}
