// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { createWorkbook } from 'ooxml-core/xlsx';
import { reviewCommands } from './commands/review';
import { createTestContext } from './commands/test-support';
import { createCommentsPane } from './comments-pane';

afterEach(() => document.body.replaceChildren());

it('keeps threaded comments current and returns focus when closed by Escape', async () => {
	const ctx = createTestContext(createWorkbook({ sheets: ['Sheet1', 'Sheet2'] }));
	ctx.commands.registerAll(reviewCommands());
	const pane = createCommentsPane(ctx);
	const opener = document.createElement('button');
	ctx.root.append(opener, pane.element);
	opener.focus();
	await ctx.commands.run('review.show-comments');
	expect(pane.element.hidden).toBe(false);
	expect(ctx.commands.get('review.show-comments')?.checked?.(ctx)).toBe(true);
	const list = pane.element.querySelector('office-ui-comments-pane')?.shadowRoot;
	expect(list?.querySelector('.empty')?.textContent).toBe('There are no comments on this sheet.');
	const session = ctx.session()!;
	session.setComment(0, { row: 1, col: 1 }, 'Legacy note', 'Ben');
	session.setComment(0, { row: 2, col: 1 }, 'Thread', 'Ana', true);
	pane.refresh();
	expect(list?.querySelector('.text')?.textContent).toBe('Thread');
	expect(list?.querySelectorAll('.thread')).toHaveLength(1);
	list?.querySelector<HTMLElement>('.thread')?.click();
	expect(ctx.selection.get().active).toEqual({ row: 2, col: 1 });
	ctx.setActiveSheet(1);
	pane.refresh();
	expect(list?.querySelectorAll('.thread')).toHaveLength(0);
	ctx.setActiveSheet(0);
	session.undo();
	pane.refresh();
	expect(list?.querySelectorAll('.thread')).toHaveLength(0);
	const add = pane.element.querySelector('button.xve-comments-new');
	ctx.setReadOnly(true);
	pane.refresh();
	expect(add?.hasAttribute('disabled')).toBe(true);
	add?.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
	);
	expect(pane.element.hidden).toBe(true);
	expect(ctx.commands.get('review.show-comments')?.checked?.(ctx)).toBe(false);
	expect(ctx.root.activeElement).toBe(opener);
});
