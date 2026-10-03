// @vitest-environment jsdom
import { createWorkbook } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createTestContext } from '../commands/test-support.js';
import { showDialog } from './frame.js';

afterEach(() => (document.body.innerHTML = ''));

function open(submit: () => Promise<string | undefined>) {
	const ctx = createTestContext(createWorkbook());
	const input = document.createElement('input');
	const result = showDialog(ctx, { heading: 'Test', name: 'frame-test', body: input, submit });
	const dialog = ctx.root.querySelector<HTMLElement>('[data-dialog="frame-test"]')!;
	const ok = dialog.querySelector<HTMLButtonElement>('.xve-btn-primary')!;
	const enter = () =>
		input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
	return { result, dialog, ok, enter };
}

describe('dialog frame', () => {
	it('runs an async submit once when OK is clicked twice or Enter is pressed twice', async () => {
		let release!: (value: string) => void;
		const submit = vi.fn(() => new Promise<string | undefined>((resolve) => (release = resolve)));
		const { result, ok, enter } = open(submit);
		ok.click();
		expect(ok.disabled).toBe(true);
		ok.click();
		enter();
		enter();
		expect(submit).toHaveBeenCalledTimes(1);
		release('done');
		await expect(result).resolves.toBe('done');
		expect(submit).toHaveBeenCalledTimes(1);
	});

	it('re-enables OK when the submit keeps the dialog open or fails', async () => {
		const submit = vi
			.fn<() => Promise<string | undefined>>()
			.mockResolvedValueOnce(undefined)
			.mockRejectedValueOnce(new Error('Bad input'))
			.mockResolvedValueOnce('ok');
		const { result, ok, enter, dialog } = open(submit);
		enter();
		await vi.waitFor(() => expect(ok.disabled).toBe(false));
		expect(dialog.isConnected).toBe(true);
		ok.click();
		await vi.waitFor(() => expect(ok.disabled).toBe(false));
		ok.click();
		await expect(result).resolves.toBe('ok');
		expect(submit).toHaveBeenCalledTimes(3);
	});
});
