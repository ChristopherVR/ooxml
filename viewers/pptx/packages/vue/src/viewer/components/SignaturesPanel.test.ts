import { mount } from '@vue/test-utils';
import { signatureStatusView } from 'ooxml-ui/pptx';
import type { SignatureState, SignatureSummary } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import SignaturesPanel from './SignaturesPanel.vue';

/** A deck whose one signature is in `state`, as `PptxData.signatures` carries it. */
function signedDeck(state: SignatureState): SignatureSummary {
	return {
		state,
		signatures: [
			{
				path: '_xmlsignatures/sig1.xml',
				state,
				signer: 'CN=Alice, O=Acme',
				issuer: 'CN=Acme CA',
				signingTime: '2024-03-01T10:00:00Z',
			},
		],
	};
}

function panel(summary: SignatureSummary | undefined) {
	return mount(SignaturesPanel, { props: { view: signatureStatusView(summary, 'en') } });
}

describe('signaturesPanel', () => {
	it('shows "Not signed" for a deck without signatures', () => {
		const wrapper = panel(undefined);
		expect(wrapper.get('.pptx-vue-signatures__title').text()).toBe('Not signed');
		expect(wrapper.get('section').attributes('data-pptx-signature-state')).toBe('unsigned');
		expect(wrapper.find('.pptx-vue-signatures__empty').exists()).toBeTruthy();
		expect(wrapper.findAll('.pptx-vue-signatures__item')).toHaveLength(0);
	});

	it('shows "Signed" and each signer for a verified deck', () => {
		const wrapper = panel({
			state: 'verified',
			signatures: [
				{ path: '_xmlsignatures/sig1.xml', state: 'verified', signer: 'CN=Alice' },
				{ path: '_xmlsignatures/sig2.xml', state: 'verified', signer: 'CN=Bob' },
			],
		});
		expect(wrapper.get('.pptx-vue-signatures__title').text()).toBe('Signed');
		expect(wrapper.get('section').attributes('data-pptx-signature-state')).toBe('verified');
		const rows = wrapper.findAll('[data-pptx-signature-row-state]');
		expect(rows.map((row) => row.attributes('data-pptx-signature-row-state'))).toStrictEqual([
			'verified',
			'verified',
		]);
		expect(rows.map((row) => row.get('.pptx-vue-signatures__signer').text())).toStrictEqual([
			'Alice',
			'Bob',
		]);
	});

	it('reads "Signed, signer not verified" for an untrusted signer', () => {
		const wrapper = panel(signedDeck('verified-untrusted'));
		expect(wrapper.get('.pptx-vue-signatures__title').text()).toBe('Signed, signer not verified');
		expect(wrapper.get('.pptx-vue-signatures__badge').text()).toBe('Signed, signer not verified');
	});

	it.each([
		['invalid', 'Invalid signature'],
		['unsupported', 'Signature not verified'],
	] as const)('never labels an %s deck or its rows "Signed"', (state: SignatureState, label) => {
		const wrapper = panel(signedDeck(state));
		expect(wrapper.get('.pptx-vue-signatures__title').text()).toBe(label);
		expect(wrapper.get('section').attributes('data-pptx-signature-state')).toBe(state);
		const row = wrapper.get('[data-pptx-signature-row-state]');
		expect(row.attributes('data-pptx-signature-row-state')).toBe(state);
		expect(row.get('.pptx-vue-signatures__badge').text()).toBe(label);
	});

	it('shows the issuer, signing time and the removed-on-save note', () => {
		const wrapper = panel(signedDeck('invalid'));
		const meta = wrapper.get('.pptx-vue-signatures__meta').text();
		expect(meta).toContain('CN=Acme CA');
		expect(meta).toContain('Signed on');
		expect(wrapper.get('.pptx-vue-signatures__removed-on-save').text()).toBe(
			'Saving changes removes the signatures, because any change invalidates them.',
		);
	});
});
