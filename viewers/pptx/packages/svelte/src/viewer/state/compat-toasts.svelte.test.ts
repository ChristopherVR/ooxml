import { SAVE_FAILED_TOAST, signatureLoadNotice, signatureStatusView } from 'ooxml-ui/pptx';
import type { PptxCompatibilityWarning, SignatureSummary } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { CompatToastsState } from './compat-toasts.svelte';

function warning(code: string, severity: 'info' | 'warning' = 'warning'): PptxCompatibilityWarning {
	return { code, severity, message: code, scope: 'presentation' } as PptxCompatibilityWarning;
}

describe('compatToastsState', () => {
	it('produces no toasts for an empty warning list', () => {
		const state = new CompatToastsState({ getWarnings: () => [] });
		expect(state.visibleToasts).toHaveLength(0);
		expect(state.overflowCount).toBe(0);
	});

	it('maps warnings through the shared decision function, deduped by code', () => {
		const state = new CompatToastsState({
			getWarnings: () => [warning('SAVE_ELEMENT_SKIPPED'), warning('SAVE_ELEMENT_SKIPPED')],
		});
		expect(state.visibleToasts).toHaveLength(1);
		expect(state.visibleToasts[0]?.code).toBe('SAVE_ELEMENT_SKIPPED');
	});

	it('caps the visible stack at 5 and reports the overflow count', () => {
		const warnings = Array.from({ length: 8 }, (_unused, i) => warning(`CODE_${i}`));
		const state = new CompatToastsState({ getWarnings: () => warnings });
		expect(state.visibleToasts).toHaveLength(5);
		expect(state.overflowCount).toBe(3);
	});

	it('dismiss(id) removes only that toast', () => {
		const state = new CompatToastsState({ getWarnings: () => [warning('A'), warning('B')] });
		state.dismiss('A');
		expect(state.visibleToasts.map((toast) => toast.code)).toStrictEqual(['B']);
	});

	it('dismissAll() clears every toast', () => {
		const state = new CompatToastsState({ getWarnings: () => [warning('A'), warning('B')] });
		state.dismissAll();
		expect(state.visibleToasts).toHaveLength(0);
	});

	it('reset() clears the dismissal state for a newly loaded document', () => {
		const state = new CompatToastsState({ getWarnings: () => [warning('A')] });
		state.dismissAll();
		expect(state.visibleToasts).toHaveLength(0);
		state.reset();
		expect(state.visibleToasts).toHaveLength(1);
	});

	it('notify() shows the save-failed notice first, again after a dismissal, until reset', () => {
		const state = new CompatToastsState({ getWarnings: () => [warning('A')] });
		state.notify(SAVE_FAILED_TOAST);
		expect(state.visibleToasts.map((toast) => toast.id)).toEqual([
			'save-failed',
			expect.any(String),
		]);
		state.dismissAll();
		state.notify(SAVE_FAILED_TOAST);
		expect(state.visibleToasts.map((toast) => toast.id)).toEqual(['save-failed']);
		state.reset();
		expect(state.visibleToasts.map((toast) => toast.id)).not.toContain('save-failed');
	});

	describe('signature load notice', () => {
		const invalid: SignatureSummary = {
			state: 'invalid',
			signatures: [{ path: '_xmlsignatures/sig1.xml', state: 'invalid' }],
		};

		it('shows a signed deck its state on top until dismissed, then again after a load', () => {
			let locale = 'en';
			const state = new CompatToastsState({ getWarnings: () => [warning('A')] });
			state.setSignatureNotice(() => signatureLoadNotice(signatureStatusView(invalid, locale)));
			const first = state.visibleToasts[0];
			expect(first?.code).toBe('SIGNATURE_STATE');
			expect(first?.severity).toBe('warning');
			expect(first?.text).toContain('is not valid');
			expect(first?.text).toContain('Saving changes removes the signatures');
			locale = 'fr';
			expect(state.visibleToasts[0]?.text).not.toContain('Saving changes');
			state.dismiss('SIGNATURE_STATE');
			expect(state.visibleToasts.map((toast) => toast.code)).toStrictEqual(['A']);
			state.reset();
			expect(state.visibleToasts[0]?.code).toBe('SIGNATURE_STATE');
		});

		it('shows nothing for an unsigned deck', () => {
			const state = new CompatToastsState({ getWarnings: () => [] });
			state.setSignatureNotice(() => signatureLoadNotice(signatureStatusView(undefined, 'en')));
			expect(state.visibleToasts).toHaveLength(0);
		});
	});
});
