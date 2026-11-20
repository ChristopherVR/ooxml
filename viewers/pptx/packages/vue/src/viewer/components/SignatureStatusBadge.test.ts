import { mount } from '@vue/test-utils';
import { signatureStatusView } from 'ooxml-ui/pptx';
import type { SignatureState } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import SignatureStatusBadge from './SignatureStatusBadge.vue';

/** A deck whose one signature is in `state`, as `PptxData.signatures` carries it. */
function badge(state: SignatureState) {
	const summary = { state, signatures: [{ path: '_xmlsignatures/sig1.xml', state }] };
	return mount(SignatureStatusBadge, { props: { view: signatureStatusView(summary, 'en') } });
}

describe('signatureStatusBadge', () => {
	it('renders nothing for an unsigned deck', () => {
		const wrapper = mount(SignatureStatusBadge, {
			props: { view: signatureStatusView(undefined, 'en') },
		});
		expect(wrapper.find('button').exists()).toBeFalsy();
	});

	it('reads "Signed" only for a verified deck', () => {
		const button = badge('verified').get('button');
		expect(button.text()).toBe('Signed');
		expect(button.attributes('data-pptx-signature-state')).toBe('verified');
	});

	it('reads "Signed, signer not verified" when the signer is not trusted', () => {
		const button = badge('verified-untrusted').get('button');
		expect(button.text()).toBe('Signed, signer not verified');
		expect(button.attributes('data-pptx-signature-state')).toBe('verified-untrusted');
	});

	it.each([
		['invalid', 'Invalid signature'],
		['unsupported', 'Signature not verified'],
	] as const)('never labels an %s deck "Signed"', (state, label) => {
		const button = badge(state).get('button');
		expect(button.text()).toBe(label);
		expect(button.attributes('data-pptx-signature-state')).toBe(state);
		expect(button.attributes('title')).not.toBe('');
	});

	it('emits click to open the signatures panel', async () => {
		const wrapper = badge('invalid');
		await wrapper.get('button').trigger('click');
		expect(wrapper.emitted('click')).toHaveLength(1);
	});
});
