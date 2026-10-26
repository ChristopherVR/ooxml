// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { loadXlsx, saveXlsx, type SmartArtObject } from 'ooxml-core/xlsx';
import { readFileSync } from 'node:fs';
import { URL as NodeURL } from 'node:url';
import { diagramNodeOutline } from 'ooxml-core/diagram';
import { createTestContext } from './commands/test-support';
import { commandTabs } from './commands/index';
import { createSmartArtController, smartArtCommands } from './smartart-controller';
import { paintSmartArt } from './grid/smartart';

afterEach(() => {
	document.body.innerHTML = '';
});

const fixture = () =>
	new Uint8Array(
		readFileSync(
			new NodeURL('../../../core/xlsx/__fixtures__/excel-smartart.xlsx', import.meta.url),
		),
	);
const smartArt = (ctx: ReturnType<typeof createTestContext>) =>
	ctx.workbook()!.sheets[0]!.drawings.find((d) => d.kind === 'smartArt') as SmartArtObject;

it('runs SmartArt Design commands and the text pane on the selected diagram, with undo', async () => {
	const ctx = createTestContext(await loadXlsx(fixture()));
	ctx.commands.registerAll(smartArtCommands());
	const controller = createSmartArtController(ctx);
	document.body.append(controller.element);
	expect(ctx.commands.isEnabled('smartart.addAfter')).toBe(false);
	const tab = commandTabs().find((entry) => entry.id === 'smartart-design')!;
	expect(tab.contextual?.(ctx)).toBe(false);
	ctx.selection.set({ drawing: 0 });
	expect(tab.contextual?.(ctx)).toBe(true);
	expect(ctx.commands.isEnabled('smartart.addAfter')).toBe(true);

	await ctx.commands.run('smartart.textPane');
	controller.refresh();
	expect(controller.element.hidden).toBe(false);
	expect(ctx.commands.get('smartart.textPane')!.checked!(ctx)).toBe(true);

	// With no line chosen the commands act on the last node, as in Word.
	await ctx.commands.run('smartart.addAfter');
	expect(smartArt(ctx).nodes.map((n) => n.text)).toEqual(['Plan', 'Build', 'Ship', '']);
	await ctx.commands.run('smartart.demote');
	expect(smartArt(ctx).nodes[3]!.parentId).toBe(smartArt(ctx).nodes[2]!.id);
	expect(smartArt(ctx).diagram).toBeUndefined();
	expect(smartArt(ctx).layoutSource?.layoutXml).toContain('layoutDef');
	const canvas = document.createElement('div');
	paintSmartArt(ctx, canvas, smartArt(ctx), ctx.workbook()!.theme);
	expect(canvas.dataset.rendering).toBe('interpreted');
	expect(canvas.querySelector('office-ui-smartart')).not.toBeNull();

	const bytes = await saveXlsx(ctx.workbook()!);
	const back = (await loadXlsx(bytes)).sheets[0]!.drawings[0] as SmartArtObject;
	expect(diagramNodeOutline(back.nodes).map((line) => [line.text, line.level])).toEqual([
		['Plan', 0],
		['Build', 0],
		['Ship', 0],
		['', 1],
	]);

	ctx.session()!.undo();
	ctx.session()!.undo();
	expect(smartArt(ctx).nodes.map((n) => n.text)).toEqual(['Plan', 'Build', 'Ship']);
	expect(smartArt(ctx).diagram).toBeDefined();
	await ctx.commands.run('smartart.colors', 'monochromatic2');
	expect(smartArt(ctx).colorsId).toBe('urn:ooxml-core/colors/monochromatic2');
	await ctx.commands.run('smartart.style', 'moderate');
	expect(smartArt(ctx).quickStyleId).toBe('urn:ooxml-core/style/moderate');
	ctx.session()!.undo();
	expect(smartArt(ctx).quickStyleId).not.toBe('urn:ooxml-core/style/moderate');
});
