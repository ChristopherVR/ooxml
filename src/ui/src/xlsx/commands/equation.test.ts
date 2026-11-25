// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { installCommands } from './index';
import { createTestContext } from './test-support';

afterEach(() => document.body.replaceChildren());

it('validates equation input and edits a selected equation through the shared dialog', async () => {
	const ctx = createTestContext();
	const stop = installCommands(ctx);
	try {
		expect(ctx.commands.isEnabled('drawing.edit-equation')).toBe(false);
		async function edit(command: string, latex: string, previous?: RegExp) {
			const pending = ctx.commands.run(command);
			const dialog = await vi.waitFor(() => {
				const open = ctx.root.querySelector('[data-dialog="equation"]');
				expect(open).not.toBeNull();
				return open!;
			});
			const input = dialog.querySelector('textarea')!;
			if (previous) expect(input.value).toMatch(previous);
			const confirm = dialog.querySelector<HTMLButtonElement>('button.xve-btn-primary')!;
			input.value = '}';
			input.dispatchEvent(new Event('input'));
			expect(confirm.disabled).toBe(true);
			input.value = latex;
			input.dispatchEvent(new Event('input'));
			expect(confirm.disabled).toBe(false);
			confirm.click();
			expect(await pending).toBe(true);
		}
		await edit('insert.equation', '\\frac{x}{2}');
		expect(ctx.commands.isEnabled('drawing.edit-equation')).toBe(true);
		await edit('drawing.edit-equation', '\\sqrt{2}', /frac/);
		const drawings = ctx.workbook()?.sheets[0]?.drawings;
		expect(drawings).toHaveLength(1);
		expect(drawings?.[0]).toMatchObject({
			shape: { equationParagraphs: [[{ fallback: '\\sqrt{2}' }]] },
		});
	} finally {
		stop();
	}
});
