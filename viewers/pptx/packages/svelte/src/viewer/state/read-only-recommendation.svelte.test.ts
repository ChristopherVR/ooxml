/**
 * read-only-recommendation.svelte.test.ts: `ReadOnlyRecommendationState` on its own, for the
 * password check timing the full-viewer wiring test cannot pin (a check that finishes after the
 * next load or after Cancel).
 */
import { createModifyVerifier } from 'pptx-viewer-core';
import type { PptxModifyVerifier } from 'pptx-viewer-core';
import { beforeAll, describe, expect, it } from 'vitest';

import { ReadOnlyRecommendationState } from './read-only-recommendation.svelte';

let verifier: PptxModifyVerifier;
beforeAll(async () => {
	verifier = await createModifyVerifier('right-password', { spinCount: 10 });
});

function promptOpen(): ReadOnlyRecommendationState {
	const state = new ReadOnlyRecommendationState({
		getModifyVerifier: () => verifier,
		getCustomProperties: () => [],
	});
	expect(state.recommendation.requiresPassword).toBeTruthy();
	state.editAnyway();
	expect(state.passwordPromptOpen).toBeTruthy();
	return state;
}

describe('readOnlyRecommendationState password check', () => {
	it('unlocks on the right password', async () => {
		const state = promptOpen();
		await state.submitPassword('right-password');
		expect(state.locked).toBeFalsy();
		expect(state.showBanner).toBeFalsy();
	});

	it('ignores a password check that finishes after the next deck loads', async () => {
		const state = promptOpen();
		const pending = state.submitPassword('right-password');
		state.reset();
		await pending;

		expect(state.locked).toBeTruthy();
		expect(state.showBanner).toBeTruthy();
		expect(state.passwordPromptOpen).toBeFalsy();
		expect(state.checkingPassword).toBeFalsy();
	});

	it('ignores a password check that finishes after Cancel', async () => {
		const state = promptOpen();
		const pending = state.submitPassword('right-password');
		state.cancelPasswordPrompt();
		await pending;

		expect(state.locked).toBeTruthy();
		expect(state.passwordPromptOpen).toBeFalsy();
		expect(state.passwordError).toBeNull();
		expect(state.checkingPassword).toBeFalsy();
	});
});
