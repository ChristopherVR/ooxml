import type { SignatureState, SignatureSummary } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';
import { effectScope, nextTick, ref } from 'vue';

import { useSignatureWorkflow } from './useSignatureWorkflow';

function deck(state: SignatureState): SignatureSummary {
	return { state, signatures: [{ path: '_xmlsignatures/sig1.xml', state }] };
}

function workflow(summary: SignatureSummary | undefined, locale = 'en') {
	const scope = effectScope();
	const signatures = ref(summary);
	const localeRef = ref(locale);
	const isDirty = ref(false);
	const result = scope.run(() => useSignatureWorkflow({ signatures, locale: localeRef, isDirty }))!;
	return { ...result, signatures, locale: localeRef, isDirty, stop: () => scope.stop() };
}

describe('useSignatureWorkflow', () => {
	it('reads "Signed" only when the signatures verify', () => {
		const w = workflow(deck('verified'));
		expect(w.view.value.label).toBe('Signed');
		w.signatures.value = deck('invalid');
		expect(w.view.value.label).toBe('Invalid signature');
		expect(w.view.value.signed).toBeFalsy();
		w.stop();
	});

	it('follows the UI locale', () => {
		const w = workflow(deck('verified-untrusted'));
		const english = w.view.value.label;
		w.locale.value = 'fr';
		expect(w.view.value.label).not.toBe(english);
		w.stop();
	});

	it('warns once on the first edit of a deck carrying signatures, verified or not', async () => {
		const w = workflow(deck('unsupported'));
		w.isDirty.value = true;
		await nextTick();
		expect(w.showSignatureStripped.value).toBeTruthy();
		w.onAckSignatureStripped();
		w.isDirty.value = false;
		await nextTick();
		w.isDirty.value = true;
		await nextTick();
		expect(w.showSignatureStripped.value).toBeFalsy();
		w.stop();
	});

	it('does not warn for an unsigned deck', async () => {
		const w = workflow(undefined);
		w.isDirty.value = true;
		await nextTick();
		expect(w.showSignatureStripped.value).toBeFalsy();
		expect(w.view.value.present).toBeFalsy();
		w.stop();
	});
});
