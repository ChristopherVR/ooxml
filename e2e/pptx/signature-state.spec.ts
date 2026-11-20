/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { test, expect } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { fixture, loadDeck } from './support/deck';

/**
 * A deck reads as signed only when its signature verifies.
 *
 * Every binding shows the loaded deck's signature state on an element carrying
 * `data-pptx-signature-state` (its text is the state's label from
 * `signatureStatusView`), and shows none for an unsigned deck. A browser trusts
 * no signer, so the best a deck can reach here is "Signed, signer not
 * verified"; the bare "Signed" must never appear for any of these decks. Before
 * OOXML 2.0 a signature part alone was enough for that label, so a tampered or
 * unverifiable signature read as signed.
 *
 * The decks are made by `viewers/pptx/scripts/make-signature-state-decks.mjs`
 * (see `fixtures/signature-states/README.md`).
 */

const deck = (state: string): string => fixture(`signature-states/${state}.pptx`);

const SIGNED_DECKS = [
	{ state: 'verified-untrusted', label: 'Signed, signer not verified' },
	{ state: 'invalid', label: 'Invalid signature' },
	{ state: 'unsupported', label: 'Signature not verified' },
] as const;

/** Visible elements showing a signature state. A closed panel or dialog does not count. */
function stateMarkers(page: Page, selector = '[data-pptx-signature-state]'): Locator {
	return page.locator(selector).filter({ visible: true });
}

/** The signature load notice in the viewer's notice stack (`signatureLoadNotice`). */
function signatureNotice(page: Page): Locator {
	return page.locator('[data-testid="pptx-compat-toast"][data-code="SIGNATURE_STATE"]');
}

/** The verified label shown for a deck whose signature did not get that far. */
function bareSignedLabel(page: Page): Locator {
	return page.getByText('Signed', { exact: true }).filter({ visible: true });
}

test.describe('signature state', () => {
	test('an unsigned deck shows no signature state', async ({ page }) => {
		await loadDeck(page, deck('unsigned'));

		await expect(stateMarkers(page)).toHaveCount(0);
		await expect(bareSignedLabel(page)).toHaveCount(0);
		await expect(signatureNotice(page)).toHaveCount(0);
	});

	for (const { state, label } of SIGNED_DECKS) {
		test(`a ${state} deck reads "${label}", never "Signed"`, async ({ page }) => {
			await loadDeck(page, deck(state));

			const badge = stateMarkers(page, `[data-pptx-signature-state="${state}"]`).first();
			await expect(badge).toBeVisible();
			await expect(badge).toHaveText(label);
			// Every surface showing a state agrees with the deck's.
			await expect(
				stateMarkers(
					page,
					`[data-pptx-signature-state]:not([data-pptx-signature-state="${state}"])`,
				),
			).toHaveCount(0);
			await expect(bareSignedLabel(page)).toHaveCount(0);
		});
	}

	test('a signed deck opens with a dismissible notice of its state', async ({ page }) => {
		await loadDeck(page, deck('invalid'));

		const notice = signatureNotice(page);
		await expect(notice).toBeVisible();
		await expect(notice).toContainText('not valid');
		await expect(notice).toContainText('Saving changes removes the signatures');
		await notice.getByTestId('pptx-compat-toast-dismiss').click();
		await expect(notice).toHaveCount(0);
	});
});
