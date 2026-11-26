import { expect, it, vi } from 'vitest';
import type { Command } from 'ooxml-core/xlsx/ui';
import { createTestContext } from '../commands/test-support';
import { renderGallery } from './controls-input';
import { closeRibbonPopover, openMenu } from './popover';

it('loads a gallery once, reports failures and replaces the open popover on retry', async () => {
	const ctx = createTestContext();
	const command: Command = { id: 'test.gallery', label: 'Styles', run() {} };
	type Items = { id: string; label: string }[];
	let rejectLoad!: (error: Error) => void;
	let resolveLoad!: (items: Items) => void;
	const failed = new Promise<Items>((_, reject) => {
		rejectLoad = reject;
	});
	const ready = new Promise<Items>((resolve) => {
		resolveLoad = resolve;
	});
	const loadItems = vi.fn().mockReturnValueOnce(failed).mockReturnValueOnce(ready);
	const button = renderGallery(
		{ ctx, doc: document, updates: [], isHidden: () => false },
		{ kind: 'gallery', command: command.id, items: () => [], loadItems },
		command,
	);
	const other = document.createElement('button');
	ctx.root.append(button, other);
	try {
		button.click();
		button.click();
		expect(loadItems).toHaveBeenCalledTimes(1);
		expect(button.getAttribute('aria-busy')).toBe('true');
		rejectLoad(new Error('chunk unavailable'));
		await vi.waitFor(() =>
			expect(ctx.toasts).toEqual([{ message: 'Could not load the gallery.', kind: 'error' }]),
		);
		expect(button.hasAttribute('aria-busy')).toBe(false);
		expect(ctx.root.querySelector('[role=menu]')).toBeNull();

		button.click();
		openMenu(other, [{ label: 'Other option', run() {} }], 'Other menu');
		resolveLoad([{ id: 'style', label: 'Frame' }]);
		await vi.waitFor(() => expect(button.getAttribute('aria-expanded')).toBe('true'));
		expect(ctx.root.querySelectorAll('[role=menu]')).toHaveLength(1);
		expect(ctx.root.querySelector('[role=menuitem]')?.getAttribute('aria-label')).toBe('Frame');
		expect(other.getAttribute('aria-expanded')).toBe('false');
		expect(button.hasAttribute('aria-busy')).toBe(false);
	} finally {
		closeRibbonPopover(ctx.root);
		ctx.host.remove();
	}
});
