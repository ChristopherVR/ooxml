import { afterEach, expect, it } from 'vitest';
import { parseVsdx } from 'ooxml-core/visio';
import JSZip from 'jszip';
import { renderText } from './render-text';
import { wrapText } from './text-layout';
import { setupFormattingViewer as setup } from './__fixtures__/formatting-viewer';

afterEach(() => document.body.replaceChildren());

it('applies font color, strike, bullets, indent and justify through the ribbon and source bytes', async () => {
	const ui = await setup();
	ui.selection();
	ui.press('font-color-red');
	await ui.done();
	expect(ui.shape().text.color).toBe('#ff0000');
	expect(ui.button('font-color-red').getAttribute('checked')).toBe('true');
	ui.press('font-color');
	await ui.done();
	expect(ui.shape().text.color).toBe('#ff0000');
	ui.press('strikethrough');
	await ui.done();
	ui.press('underline');
	await ui.done();
	expect(ui.shape().text.runs[0]).toMatchObject({ strikethrough: true, underline: true });
	expect(ui.button('strikethrough').getAttribute('pressed')).toBe('true');
	const svg = renderText(ui.shape().text, new Set());
	expect(svg.querySelector('tspan tspan')!.getAttribute('text-decoration')).toBe(
		'underline line-through',
	);
	expect(ui.button('indent-decrease').disabled).toBe(true);
	ui.press('indent-increase');
	await ui.done();
	expect(ui.shape().text.paragraphs![0]!.indentLeft * 72).toBeCloseTo(18);
	expect(ui.button('indent-decrease').disabled).toBe(false);
	ui.press('bullets');
	await ui.done();
	expect(ui.shape().text.paragraphs![0]!.bullet).toBeDefined();
	expect(ui.button('bullets').getAttribute('pressed')).toBe('true');
	ui.press('justify');
	await ui.done();
	expect(ui.shape().text.horizontalAlign).toBe('justify');
	expect(ui.button('justify').getAttribute('pressed')).toBe('true');
	const saved = (await parseVsdx(ui.controller.exportVsdx().bytes)).pages[0]!.shapes[0]!;
	expect(saved.text.runs[0]).toMatchObject({
		strikethrough: true,
		underline: true,
		color: '#ff0000',
	});
	expect(saved.text.paragraphs![0]).toMatchObject({ horizontalAlign: 'justify', indentLeft: 0.25 });
	ui.press('bullets');
	await ui.done();
	expect(ui.shape().text.paragraphs![0]!.bullet).toBeUndefined();
	ui.press('indent-decrease');
	await ui.done();
	expect(ui.shape().text.paragraphs![0]!.indentLeft).toBe(0);
	ui.dispose();
	ui.controller.destroy();
});

it('sets mixed-selection toggles together and preserves all selected targets through undo', async () => {
	const ui = await setup();
	await ui.controller.applyEdits([
		{
			type: 'create-rectangle',
			pageId: '1',
			shapeId: '2',
			x: 1,
			y: 1,
			width: 2,
			height: 1,
			text: 'Second shape',
		},
		{ type: 'format-text', pageId: '1', shapeId: '1', bold: true, fontSize: 24 },
	]);
	ui.controller.selectShapes([
		{ id: '1', name: 'First', pageId: '1' },
		{ id: '2', name: 'Second', pageId: '1' },
	]);
	expect(ui.button('bold').getAttribute('pressed')).toBe('false');
	expect(ui.combo('font-size').value).toBe('');
	expect(ui.button('bring-to-front').disabled).toBe(true);
	ui.press('bold');
	await ui.done();
	expect(ui.edits.at(-1)).toEqual([
		{ type: 'format-text', pageId: '1', shapeId: '1', bold: true },
		{ type: 'format-text', pageId: '1', shapeId: '2', bold: true },
	]);
	expect(
		ui.controller.state.document!.pages[0]!.shapes.every((shape) => shape.text.runs[0]?.bold),
	).toBe(true);
	expect(ui.controller.state.selectedShapes.map((shape) => shape.id)).toEqual(['1', '2']);
	ui.press('undo');
	await ui.done();
	expect(
		ui.controller.state.document!.pages[0]!.shapes.map((shape) => shape.text.runs[0]?.bold),
	).toEqual([true, false]);
	ui.press('redo');
	await ui.done();
	ui.press('bold');
	await ui.done();
	expect(
		ui.controller.state.document!.pages[0]!.shapes.every((shape) => !shape.text.runs[0]?.bold),
	).toBe(true);
	ui.press('fill-blue');
	await ui.done();
	expect(ui.edits.at(-1)).toHaveLength(2);
	expect(
		ui.controller.state.document!.pages[0]!.shapes.every((shape) => shape.style.fill === '#4472c4'),
	).toBe(true);
	expect(ui.button('fill-blue').getAttribute('checked')).toBe('true');
	ui.dispose();
	ui.controller.destroy();
});

it('refuses the whole incompatible selection instead of editing its primary shape', async () => {
	const ui = await setup();
	await ui.controller.applyEdits([
		{
			type: 'create-rectangle',
			pageId: '1',
			shapeId: '2',
			x: 1,
			y: 1,
			width: 2,
			height: 1,
			text: 'Second',
		},
	]);
	ui.shape().masterId = 'unsupported-master';
	ui.controller.selectShapes([
		{ id: '1', name: 'First', pageId: '1' },
		{ id: '2', name: 'Second', pageId: '1' },
	]);
	expect(ui.button('strikethrough').disabled).toBe(true);
	expect(ui.button('fill-blue').disabled).toBe(true);
	const count = ui.edits.length;
	ui.commands.run({ type: 'text-toggle', property: 'strikethrough' });
	expect(ui.edits).toHaveLength(count);
	ui.dispose();
	ui.controller.destroy();
});

it('retains neighboring strike styles when wrapping and merging text runs', () => {
	const run = {
		text: 'a',
		fontFamily: 'Arial',
		fontSize: 0.2,
		color: '#000000',
		bold: false,
		italic: false,
		underline: false,
	};
	const lines = wrapText(
		[run, { ...run, text: 'b', strikethrough: true }],
		5,
		(text) => text.length * 0.1,
	);
	expect(lines[0]!.runs).toHaveLength(2);
	expect(lines[0]!.runs[1]!.strikethrough).toBe(true);
});

it('preserves the complete source when a later target refuses an atomic formatting batch', async () => {
	const ui = await setup();
	await ui.controller.applyEdits([
		{
			type: 'create-rectangle',
			pageId: '1',
			shapeId: '2',
			x: 1,
			y: 1,
			width: 2,
			height: 1,
			text: 'Locked shape',
		},
	]);
	const zip = await JSZip.loadAsync(ui.controller.exportVsdx().bytes);
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	zip.file(
		'visio/pages/page1.xml',
		xml.replace(/<Shape\b[^>]*\bID="2"[^>]*>/u, '$&<Cell N="LockFormat" V="1"/>'),
	);
	const bytes = await zip.generateAsync({ type: 'uint8array' });
	await ui.controller.load(bytes);
	ui.controller.selectShapes([
		{ id: '1', name: 'First', pageId: '1' },
		{ id: '2', name: 'Locked', pageId: '1' },
	]);
	expect(ui.button('font-color-red').disabled).toBe(false);
	ui.press('font-color-red');
	await ui.done();
	expect(ui.controller.state.edit.error).toBeDefined();
	expect(ui.controller.exportVsdx().bytes).toEqual(bytes);
	expect(ui.controller.state.edit.canUndo).toBe(false);
	expect(ui.controller.state.selectedShapes).toHaveLength(2);
	ui.dispose();
	ui.controller.destroy();
});
