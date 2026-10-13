import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook } from 'ooxml-core/xlsx';
import { createTestContext } from '../commands/test-support';
import { menuEntries, type RenderScope } from './controls';

afterEach(() => {
	document.body.innerHTML = '';
});

describe('menu command focus', () => {
	it('does not steal focus from a different ribbon input after an asynchronous command', async () => {
		const focus = vi.fn();
		const ctx = createTestContext(createWorkbook(), { grid: { focus } });
		const ribbon = document.createElement('div');
		ribbon.setAttribute('part', 'ribbon');
		ctx.root.append(ribbon);
		const anchor = document.createElement('button');
		const input = document.createElement('input');
		ribbon.append(anchor, input);
		anchor.focus();
		let complete!: () => void;
		ctx.commands.register({
			id: 'test.async',
			label: 'Async',
			run: () =>
				new Promise<void>((resolve) => {
					complete = resolve;
				}),
		});
		const scope: RenderScope = { ctx, doc: document, updates: [], isHidden: () => false };
		const entry = menuEntries(scope, [{ command: 'test.async' }])[0]!;
		if ('separator' in entry) throw new Error('Expected a command');
		entry.run();
		input.focus();
		complete();
		await Promise.resolve();
		await Promise.resolve();
		await Promise.resolve();
		expect(ctx.root.activeElement).toBe(input);
		expect(focus).not.toHaveBeenCalled();
	});
	it('returns to the grid before a command can leave its disabled trigger unfocusable', async () => {
		const focus = vi.fn();
		const ctx = createTestContext(createWorkbook(), { grid: { focus } });
		const ribbon = document.createElement('div');
		ribbon.setAttribute('part', 'ribbon');
		ctx.root.append(ribbon);
		const anchor = document.createElement('button');
		ribbon.append(anchor);
		anchor.focus();
		ctx.commands.register({
			id: 'test.clear',
			label: 'Clear',
			run: () => {
				anchor.disabled = true;
			},
		});
		const scope: RenderScope = { ctx, doc: document, updates: [], isHidden: () => false };
		const entry = menuEntries(scope, [{ command: 'test.clear' }])[0]!;
		if ('separator' in entry) throw new Error('Expected a command');
		entry.run();
		await vi.waitFor(() => expect(focus).toHaveBeenCalledOnce());
	});
	it('keeps focus in a dialog opened by the menu command', async () => {
		const focus = vi.fn();
		const ctx = createTestContext(createWorkbook(), { grid: { focus } });
		const input = document.createElement('input');
		ctx.root.append(input);
		ctx.commands.register({
			id: 'test.dialog',
			label: 'Dialog',
			run: () => {
				input.focus();
			},
		});
		const scope: RenderScope = { ctx, doc: document, updates: [], isHidden: () => false };
		const entry = menuEntries(scope, [{ command: 'test.dialog' }])[0]!;
		if ('separator' in entry) throw new Error('Expected a command');
		entry.run();
		await Promise.resolve();
		await Promise.resolve();
		expect(ctx.root.activeElement).toBe(input);
		expect(focus).not.toHaveBeenCalled();
	});
});
