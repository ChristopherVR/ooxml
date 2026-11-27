// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { setInkMathRecognizer } from '../../ink/math-recognizer';
import { installCommands } from './index';
import { createTestContext } from './test-support';

afterEach(() => document.body.replaceChildren());

it('turns Ink to Math on with the host recognizer and inserts an equation text box', async () => {
	const ctx = createTestContext();
	const stop = installCommands(ctx);
	try {
		expect(ctx.commands.isEnabled('draw.ink-to-math')).toBe(false);
		const remove = setInkMathRecognizer(async () => 'x^2');
		try {
			expect(ctx.commands.isEnabled('draw.ink-to-math')).toBe(true);
			await ctx.commands.run('draw.ink-to-math');
			const dialog = await vi.waitFor(() => {
				const open = ctx.root.querySelector('office-ui-dialog');
				expect(open).not.toBeNull();
				return open!;
			});
			const source = dialog.querySelector('textarea')!;
			source.value = 'x^2';
			source.dispatchEvent(new Event('input'));
			dialog
				.querySelector('office-ui-dialog-footer')!
				.dispatchEvent(
					new CustomEvent('office-dialog-footer-request', { detail: { id: 'insert' } }),
				);
			expect(ctx.workbook()?.sheets[0]?.drawings[0]).toMatchObject({
				shape: { equationParagraphs: [[{ fallback: 'x^2' }]] },
			});
		} finally {
			remove();
		}
		expect(ctx.commands.isEnabled('draw.ink-to-math')).toBe(false);
	} finally {
		stop();
	}
});
