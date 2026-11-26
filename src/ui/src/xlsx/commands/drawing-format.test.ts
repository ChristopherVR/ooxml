import { expect, it } from 'vitest';
import { pictureAnchorAt } from 'ooxml-core/xlsx';
import { drawingFormatCommands, formatTabContext } from './drawing-format';
import { createTestContext } from './test-support';

it('limits formatting to a supported selection on an editable sheet', async () => {
	const ctx = createTestContext();
	ctx.commands.registerAll(drawingFormatCommands());
	const session = ctx.session()!;
	const anchor = pictureAnchorAt({ row: 1, col: 1 }, 100, 50);
	const shape = session.addShape(0, anchor, { kind: 'preset', preset: 'rect', name: 'Box' });
	session.addImage(0, Uint8Array.of(0), 'image/png', anchor, 'Logo');
	expect(formatTabContext(ctx, 'shape')).toBe(false);
	ctx.selection.set({ drawing: shape });
	expect(formatTabContext(ctx, 'shape')).toBe(true);
	expect(ctx.commands.isEnabled('drawing.edit-shape')).toBe(true);
	expect(ctx.commands.isEnabled('drawing.word-art-style')).toBe(false);
	expect(ctx.commands.isEnabled('picture.crop')).toBe(false);
	expect(await ctx.commands.run('drawing.edit-shape', 'hexagon')).toBe(true);
	expect(session.workbook.sheets[0]?.drawings[shape]).toMatchObject({
		shape: { geometry: 'hexagon' },
	});
	expect(session.undo()).toBe(true);
	expect(session.workbook.sheets[0]?.drawings[shape]).toMatchObject({
		shape: { geometry: 'rect' },
	});

	ctx.selection.set({ drawing: 1 });
	expect(formatTabContext(ctx, 'picture')).toBe(true);
	expect(ctx.commands.isEnabled('picture.artistic-effects')).toBe(false);
	expect(ctx.commands.isEnabled('picture.corrections')).toBe(true);
	ctx.setReadOnly(true);
	expect(ctx.commands.isEnabled('picture.corrections')).toBe(false);
	expect(ctx.commands.isEnabled('drawing.alt-text')).toBe(false);
	ctx.host.remove();
});

it('reuses native WordArt editing for style and text effects', async () => {
	const ctx = createTestContext();
	ctx.commands.registerAll(drawingFormatCommands());
	const session = ctx.session()!;
	const art = session.addWordArt(0, pictureAnchorAt({ row: 0, col: 0 }, 100, 50), {
		text: 'Sales',
		style: { fill: '000000' },
	});
	ctx.selection.set({ drawing: art });
	expect(ctx.commands.isEnabled('drawing.word-art-style')).toBe(true);
	expect(await ctx.commands.run('drawing.word-art-style', '1')).toBe(true);
	expect(await ctx.commands.run('drawing.text-effects', true)).toBe(true);
	const drawing = () => session.workbook.sheets[0]?.drawings[art];
	expect(drawing()).toMatchObject({
		shape: { wordArt: { fill: '1769A5', outline: '103D63', shadow: '404040' } },
	});
	expect(session.undo()).toBe(true);
	expect(drawing()).not.toHaveProperty('shape.wordArt.shadow');
	ctx.host.remove();
});
