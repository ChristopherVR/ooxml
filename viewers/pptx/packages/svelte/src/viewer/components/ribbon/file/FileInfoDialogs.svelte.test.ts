import { signatureStatusView } from 'ooxml-ui/pptx';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { footerAction } from '../../dialog-footer.test-support';
import DigitalSignaturesDialog from './DigitalSignaturesDialog.svelte';
import PasswordProtectionDialog from './PasswordProtectionDialog.svelte';

const cleanups: Array<() => void> = [];
afterEach(() => cleanups.splice(0).forEach((cleanup) => cleanup()));

describe('file info dialogs', () => {
	it('lists each signature with its own verification state', () => {
		const target = document.createElement('div');
		const view = signatureStatusView(
			{
				state: 'invalid',
				signatures: [
					{ path: '_xmlsignatures/sig1.xml', state: 'verified-untrusted' },
					{ path: '_xmlsignatures/sig2.xml', state: 'invalid' },
				],
			},
			'en',
		);
		const instance = mount(DigitalSignaturesDialog, { target, props: { view, onclose: vi.fn() } });
		cleanups.push(() => unmount(instance));
		const rows = [...target.querySelectorAll<HTMLElement>('[data-pptx-signature-row-state]')];
		expect(rows.map((row) => row.dataset.pptxSignatureRowState)).toEqual([
			'verified-untrusted',
			'invalid',
		]);
		expect(rows[0]?.textContent).toContain('Signed, signer not verified');
		expect(rows[1]?.textContent).toContain('Invalid signature');
	});

	it('validates and accepts a protection password', () => {
		const target = document.createElement('div');
		const onset = vi.fn();
		const instance = mount(PasswordProtectionDialog, {
			target,
			props: { protected: false, onset, onremove: vi.fn(), onclose: vi.fn() },
		});
		cleanups.push(() => unmount(instance));
		const inputs = target.querySelectorAll<HTMLInputElement>('input');
		inputs[0].value = 'Safe123!';
		inputs[0].dispatchEvent(new Event('input', { bubbles: true }));
		inputs[1].value = 'Safe123!';
		inputs[1].dispatchEvent(new Event('input', { bubbles: true }));
		flushSync();
		footerAction(target, 'submit')!.click();
		expect(onset).toHaveBeenCalledWith('Safe123!');
	});

	it('refuses a mismatched or too short password with the shared rule', () => {
		const target = document.createElement('div');
		const onset = vi.fn();
		const instance = mount(PasswordProtectionDialog, {
			target,
			props: { protected: false, onset, onremove: vi.fn(), onclose: vi.fn() },
		});
		cleanups.push(() => unmount(instance));
		const inputs = target.querySelectorAll<HTMLInputElement>('input');
		const enter = (first: string, second: string) => {
			inputs[0].value = first;
			inputs[0].dispatchEvent(new Event('input', { bubbles: true }));
			inputs[1].value = second;
			inputs[1].dispatchEvent(new Event('input', { bubbles: true }));
			flushSync();
			footerAction(target, 'submit')!.click();
			flushSync();
		};
		enter('Safe123!', 'Other456!');
		expect(target.querySelector('[role="alert"]')).not.toBeNull();
		enter('abc', 'abc');
		expect(target.querySelector('[role="alert"]')).not.toBeNull();
		expect(onset).not.toHaveBeenCalled();
	});
});
