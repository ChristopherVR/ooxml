import { afterEach, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { parseVsdx } from 'ooxml-core/visio';
import { setupFormattingViewer as setup } from './__fixtures__/formatting-viewer';
import type { PaintField, PaintInput, PaintDialog } from './viewer-paint-dialog';
import { createContextMenus } from './viewer-context-menu';

afterEach(() => document.body.replaceChildren());
const dialog = (ui: Awaited<ReturnType<typeof setup>>) =>
	ui.root.querySelector<PaintDialog>('.paint-properties-dialog')!;
const field = (ui: Awaited<ReturnType<typeof setup>>, name: PaintField) =>
	dialog(ui).querySelector<PaintInput>(`[data-paint-field="${name}"]`)!;
async function open(ui: Awaited<ReturnType<typeof setup>>, menu = 'fill-options') {
	ui.press(menu);
	await vi.waitFor(() => expect(dialog(ui).open).toBe(true));
}
async function twoShapes(ui: Awaited<ReturnType<typeof setup>>) {
	ui.selection();
	await ui.controller.duplicateSelection();
	ui.controller.selectShapes([
		{ id: '1', name: 'One', pageId: '1' },
		{ id: '2', name: 'Two', pageId: '1' },
	]);
	ui.edits.length = 0;
}

it('exposes native line choices without color or weight implicitly enabling No Line', async () => {
	const ui = await setup();
	ui.selection();
	ui.press('line-pattern-0');
	await ui.done();
	expect(ui.shape().style.linePattern).toBe(0);
	expect(ui.button('line-pattern-0').getAttribute('checked')).toBe('true');
	ui.press('line-blue');
	await ui.done();
	ui.press('line-weight-3');
	await ui.done();
	expect(ui.shape().style.linePattern).toBe(0);
	for (const value of [1, 23]) {
		ui.press(`line-pattern-${value}`);
		await ui.done();
		expect(ui.shape().style.linePattern).toBe(value);
		expect(ui.button(`line-pattern-${value}`).getAttribute('checked')).toBe('true');
	}
	ui.dispose();
	ui.controller.destroy();
});

it('applies changed paint fields once, preserves untouched fields and exact source history', async () => {
	const ui = await setup();
	ui.selection();
	await open(ui);
	expect(field(ui, 'fillPattern').value).toBe('1');
	field(ui, 'fillPattern').value = '24';
	field(ui, 'fillBackgroundColor').value = '#123456';
	field(ui, 'fillTransparency').value = '17.5';
	field(ui, 'lineTransparency').value = '62.5';
	ui.press('paint-apply');
	await ui.done();
	await vi.waitFor(() => expect(dialog(ui).open).toBe(false));
	expect(ui.edits).toEqual([
		[
			{
				type: 'format-shape',
				pageId: '1',
				shapeId: '1',
				fillPattern: 24,
				fillBackgroundColor: '#123456',
				fillTransparency: 17.5,
				lineTransparency: 62.5,
			},
		],
	]);
	const style = ui.shape().style;
	expect(style.fillForegroundOpacity).toBeCloseTo(0.825);
	expect(style.fillBackgroundOpacity).toBeCloseTo(0.825);
	expect(style.lineColorOpacity).toBeCloseTo(0.375);
	expect(style.fill).toBe('#DAEFE6');
	expect(style.linePattern).toBe(1);
	const accepted = ui.controller.exportVsdx().bytes;
	expect((await parseVsdx(accepted)).pages[0]!.shapes[0]!.style).toEqual(style);
	ui.press('undo');
	await ui.done();
	expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
	ui.press('redo');
	await ui.done();
	expect(ui.controller.exportVsdx().bytes).toEqual(accepted);
	expect(ui.controller.state.selectedShapes.map((shape) => shape.id)).toEqual(['1']);
	ui.dispose();
	ui.controller.destroy();
});

it('shows mixed and mismatched foreground/background values and edits all selected targets atomically', async () => {
	const ui = await setup();
	await twoShapes(ui);
	await ui.controller.applyEdits([
		{ type: 'format-shape', pageId: '1', shapeId: '2', linePattern: 0, fillTransparency: 45 },
	]);
	ui.edits.length = 0;
	const selection = ui.controller.state.selectedShapes;
	await open(ui, 'line-options');
	expect(field(ui, 'linePattern').value).toBe('');
	expect(field(ui, 'fillTransparency').value).toBe('');
	field(ui, 'lineTransparency').value = '25';
	ui.press('paint-apply');
	await ui.done();
	expect(ui.edits).toEqual([
		[
			{ type: 'format-shape', pageId: '1', shapeId: '1', lineTransparency: 25 },
			{ type: 'format-shape', pageId: '1', shapeId: '2', lineTransparency: 25 },
		],
	]);
	expect(ui.controller.state.selectedShapes).toBe(selection);
	expect(
		ui.controller.state.document!.pages[0]!.shapes.map((shape) => shape.style.linePattern),
	).toEqual([1, 0]);
	ui.dispose();
	ui.controller.destroy();
});

it('rejects invalid drafts and any protected target with source and history untouched', async () => {
	const ui = await setup();
	await twoShapes(ui);
	const zip = await JSZip.loadAsync(ui.controller.exportVsdx().bytes);
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	zip.file(
		'visio/pages/page1.xml',
		xml.replace(/(<Shape\b[^>]*\bID="2"[^>]*>)/u, '$1<Cell N="LockFormat" V="1"/>'),
	);
	await ui.controller.load(await zip.generateAsync({ type: 'uint8array' }));
	ui.controller.selectAll();
	const before = ui.controller.exportVsdx().bytes;
	await open(ui);
	field(ui, 'fillTransparency').value = '101';
	ui.press('paint-apply');
	await vi.waitFor(() =>
		expect(dialog(ui).querySelector('[role="alert"]')!.textContent).toMatch(/0 to 100/),
	);
	expect(ui.controller.exportVsdx().bytes).toEqual(before);
	field(ui, 'fillTransparency').value = '25';
	ui.press('paint-apply');
	await ui.done();
	await vi.waitFor(() =>
		expect(dialog(ui).querySelector('[role="alert"]')!.textContent).toMatch(/EDIT_PROTECTED_CELL/),
	);
	expect(dialog(ui).open).toBe(true);
	expect(ui.controller.state.edit.canUndo).toBe(false);
	expect(ui.controller.exportVsdx().bytes).toEqual(before);
	ui.dispose();
	ui.controller.destroy();
});

it('cancels dirty drafts on close, selection/source changes and pending cancellation', async () => {
	const ui = await setup();
	ui.selection();
	await open(ui);
	field(ui, 'lineTransparency').value = '75';
	ui.press('paint-cancel');
	expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
	await open(ui);
	expect(field(ui, 'lineTransparency').value).toBe('0');
	field(ui, 'lineTransparency').value = '75';
	ui.controller.clearSelection();
	expect(dialog(ui).open).toBe(false);
	expect(ui.edits).toHaveLength(0);
	ui.selection();
	await open(ui);
	await ui.controller.load(ui.bytes);
	expect(dialog(ui).open).toBe(false);
	ui.selection();
	await open(ui);
	field(ui, 'lineTransparency').value = '75';
	ui.press('paint-apply');
	expect(ui.controller.state.edit.busy).toBe(true);
	expect(ui.button('paint-cancel').disabled).toBe(false);
	ui.press('paint-cancel');
	await ui.done();
	expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
	expect(ui.controller.state.edit.canUndo).toBe(false);
	expect(ui.feedback).not.toContain('Updated fill and line formatting.');
	ui.dispose();
	ui.controller.destroy();
});

it('does not publish stale success when an accepted edit callback changes selection', async () => {
	const ui = await setup();
	ui.selection();
	await open(ui);
	field(ui, 'lineTransparency').value = '75';
	const original = ui.controller.state.document;
	const unsubscribe = ui.controller.subscribe((state) => {
		if (state.document !== original && state.selectedShapes.length) ui.controller.clearSelection();
	});
	ui.press('paint-apply');
	await ui.done();
	await vi.waitFor(() => expect(dialog(ui).open).toBe(false));
	expect(ui.shape().style.lineColorOpacity).toBe(0.25);
	expect(ui.feedback).not.toContain('Updated fill and line formatting.');
	unsubscribe();
	ui.dispose();
	ui.controller.destroy();
});

it('opens the canonical dialog from context menus and declines model-only/background selections', async () => {
	const ui = await setup();
	ui.root.append(...createContextMenus(document));
	ui.selection();
	await open(ui, 'ctx-format');
	ui.press('paint-apply');
	expect(dialog(ui).open).toBe(false);
	expect(ui.edits).toHaveLength(0);
	expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
	const zip = await JSZip.loadAsync(ui.bytes);
	const pages = await zip.file('visio/pages/pages.xml')!.async('string');
	const background = pages
		.match(/<Page ID="1".*?<\/Page>/su)![0]
		.replace('ID="1"', 'ID="2" Background="1"')
		.replace('rId1', 'rId2');
	zip.file(
		'visio/pages/pages.xml',
		pages.replace('ID="1"', 'ID="1" BackPage="2"').replace('</Pages>', `${background}</Pages>`),
	);
	zip.file('visio/pages/page2.xml', await zip.file('visio/pages/page1.xml')!.async('string'));
	const relationships = await zip.file('visio/pages/_rels/pages.xml.rels')!.async('string');
	zip.file(
		'visio/pages/_rels/pages.xml.rels',
		relationships.replace(
			'</Relationships>',
			'<Relationship Id="rId2" Type="http://schemas.microsoft.com/visio/2010/relationships/page" Target="page2.xml"/></Relationships>',
		),
	);
	await ui.controller.load(await zip.generateAsync({ type: 'uint8array' }));
	ui.controller.selectShape({ id: '1', name: 'Background', pageId: '2' });
	expect(ui.controller.state.selectedShape?.pageId).toBe('2');
	ui.commands.run({ type: 'paint-properties' });
	expect(dialog(ui).open).toBe(false);
	expect(ui.button('ctx-format').disabled).toBe(true);
	ui.controller.setDocument(ui.controller.state.document!);
	ui.selection();
	expect(ui.button('fill-options').disabled).toBe(true);
	ui.commands.run({ type: 'paint-properties' });
	expect(dialog(ui).open).toBe(false);
	ui.dispose();
	ui.controller.destroy();
});

it('shows Mixed for different foreground/background transparency within a single shape', async () => {
	const ui = await setup();
	const zip = await JSZip.loadAsync(ui.bytes);
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	zip.file(
		'visio/pages/page1.xml',
		xml.replace(
			'<Text>',
			'<Cell N="FillForegndTrans" V="0.1"/><Cell N="FillBkgndTrans" V="0.2"/><Text>',
		),
	);
	await ui.controller.load(await zip.generateAsync({ type: 'uint8array' }));
	ui.selection();
	await open(ui);
	expect(field(ui, 'fillTransparency').value).toBe('');
	field(ui, 'lineTransparency').value = '30';
	ui.press('paint-apply');
	await ui.done();
	expect(ui.edits.at(-1)).toEqual([
		{ type: 'format-shape', pageId: '1', shapeId: '1', lineTransparency: 30 },
	]);
	expect(ui.shape().style.fillForegroundOpacity).toBe(0.9);
	expect(ui.shape().style.fillBackgroundOpacity).toBe(0.8);
	ui.dispose();
	ui.controller.destroy();
});

it('does not promote a newer selection intent triggered by a source-state getter', async () => {
	const ui = await setup();
	ui.selection();
	let fired = false;
	const style = ui.shape().style;
	const pattern = style.linePattern;
	Object.defineProperty(style, 'linePattern', {
		configurable: true,
		get() {
			if (!fired) {
				fired = true;
				ui.controller.clearSelection();
			}
			return pattern;
		},
	});
	ui.commands.run({ type: 'paint-properties' });
	expect(dialog(ui).open).toBe(false);
	expect(ui.edits).toHaveLength(0);
	ui.dispose();
	ui.controller.destroy();
});

it('displays unsupported current patterns and preserves them when another field changes', async () => {
	const ui = await setup();
	const zip = await JSZip.loadAsync(ui.bytes);
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	zip.file(
		'visio/pages/page1.xml',
		xml.replace('<Text>', '<Cell N="FillPattern" V="25"/><Cell N="LinePattern" V="254"/><Text>'),
	);
	await ui.controller.load(await zip.generateAsync({ type: 'uint8array' }));
	ui.selection();
	const original = ui.controller.exportVsdx().bytes;
	await open(ui);
	for (const [name, value] of [
		['fillPattern', '25'],
		['linePattern', '254'],
	] as const) {
		const control = field(ui, name);
		expect(control.value).toBe(value);
		expect(
			'options' in control && control.options.find((item) => item.value === value),
		).toMatchObject({ disabled: true, label: `Current pattern ${value}` });
	}
	ui.press('paint-apply');
	expect(ui.controller.exportVsdx().bytes).toEqual(original);
	expect(ui.edits).toHaveLength(0);
	await open(ui);
	field(ui, 'lineTransparency').value = '25';
	ui.press('paint-apply');
	await ui.done();
	expect(ui.edits.at(-1)).toEqual([
		{ type: 'format-shape', pageId: '1', shapeId: '1', lineTransparency: 25 },
	]);
	expect(ui.shape().style.linePattern).toBe(254);
	expect(ui.shape().style.fillPatternIndex).toBe(25);
	ui.dispose();
	ui.controller.destroy();
});

it('suppresses success if restoring focus through close replaces the source', async () => {
	const ui = await setup();
	ui.selection();
	await open(ui);
	field(ui, 'lineTransparency').value = '25';
	dialog(ui).addEventListener('office-dialog-close', () => ui.controller.setDocument(null), {
		once: true,
	});
	ui.press('paint-apply');
	await ui.done();
	expect(ui.controller.state.document).toBe(null);
	expect(ui.feedback).not.toContain('Updated fill and line formatting.');
	ui.dispose();
	ui.controller.destroy();
});

it('waits for shared dialog focus restoration before announcing accepted formatting', async () => {
	const ui = await setup();
	ui.selection();
	const outside = document.createElement('button');
	document.body.append(outside);
	outside.focus();
	ui.commands.run({ type: 'paint-properties' });
	await dialog(ui).updateComplete;
	field(ui, 'lineTransparency').value = '25';
	outside.addEventListener('focus', () => ui.controller.setDocument(null), { once: true });
	ui.press('paint-apply');
	await ui.done();
	await vi.waitFor(() => expect(ui.controller.state.document).toBe(null));
	expect(ui.feedback).not.toContain('Updated fill and line formatting.');
	ui.dispose();
	ui.controller.destroy();
});
