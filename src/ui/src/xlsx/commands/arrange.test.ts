import { afterEach, expect, it } from 'vitest';
import { installCommands } from './index';
import { shapeSections } from './insert-shapes';
import { clickButton, createTestContext, dialogEl, inputByLabel } from './test-support';
import { createSelectionPane } from '../selection-pane';
import { setDrawingSelection } from '../drawing-selection';

afterEach(() => document.body.replaceChildren());

it('adds native drawings through the ribbon and groups selected shapes with undo', async () => {
	const ctx = createTestContext();
	const stop = installCommands(ctx);
	try {
		const galleryItem = shapeSections(ctx)
			.flatMap((section) => section.items)
			.find((item) => item.id === 'rect');
		expect(galleryItem?.preview).toContain('<path');
		await ctx.commands.run('insert.shapes', galleryItem?.id);
		await ctx.commands.run('insert.shapes', 'ellipse');
		const open = ctx.commands.run('insert.text-box');
		const dialog = dialogEl(ctx, 'text-box');
		inputByLabel<HTMLTextAreaElement>(dialog, 'Text').value = 'Ribbon text';
		clickButton(dialog, 'OK');
		expect(await open).toBe(true);

		const sheet = ctx.workbook()?.sheets[0];
		expect(sheet?.drawings[2]).toMatchObject({
			shape: { textBox: true, text: { text: 'Ribbon text' } },
		});
		const pane = createSelectionPane(ctx);
		ctx.root.append(pane.element);
		await ctx.commands.run('arrange.selection-pane');
		expect(pane.element.hidden).toBe(false);
		const top = pane.element.querySelector<HTMLButtonElement>('li button');
		expect(top?.textContent).toBe('Text Box');
		top?.click();
		await ctx.commands.run('arrange.backward');
		expect(sheet?.drawings[1]).toMatchObject({ shape: { textBox: true } });
		setDrawingSelection(ctx, [0, 2]);
		expect(await ctx.commands.run('arrange.group')).toBe(true);
		expect(sheet?.drawings).toHaveLength(2);
		expect(ctx.session()?.undo()).toBe(true);
		expect(sheet?.drawings).toHaveLength(3);
	} finally {
		stop();
	}
});
