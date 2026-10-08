import { afterEach, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { editVsdx, parseVsdx, type VisioEdit } from 'ooxml-core/visio';
import { ViewerController } from './controller';
import { ViewerCommands } from './viewer-commands';
import { createRibbon } from './ribbon';
import { createRulers } from './viewer-ruler';
import { registerViewerControls } from './office-ui';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';

afterEach(() => document.body.replaceChildren());

async function setup(source = true, mixedText = false) {
	registerViewerControls();
	const zip = await JSZip.loadAsync(await createVsdxFixture('Formatted shape'));
	const documentXml = await zip.file('visio/document.xml')!.async('string');
	zip.file(
		'visio/document.xml',
		documentXml.replace(
			'/>',
			'><FaceNames><FaceName ID="0" Name="Arial"/><FaceName ID="1" Name="Calibri"/></FaceNames></VisioDocument>',
		),
	);
	if (mixedText) {
		const pageXml = await zip.file('visio/pages/page1.xml')!.async('string');
		zip.file(
			'visio/pages/page1.xml',
			pageXml.replace(
				'<Text>Formatted shape</Text>',
				'<Section N="Character"><Row IX="0"><Cell N="Style" V="0"/></Row><Row IX="1"><Cell N="Style" V="1"/></Row></Section><Text><cp IX="0"/>Plain <cp IX="1"/>bold</Text>',
			),
		);
	}
	const bytes = await zip.generateAsync({ type: 'uint8array' });
	const edits: VisioEdit[][] = [];
	const controller = new ViewerController(
		parseVsdx,
		() => {},
		async (bytes, commands) => {
			edits.push([...commands]);
			const result = await editVsdx(bytes, commands);
			return { ...result, document: await parseVsdx(result.bytes) };
		},
	);
	if (source) await controller.load(bytes);
	else controller.setDocument(await parseVsdx(bytes));
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	const viewport = document.createElement('div');
	viewport.tabIndex = 0;
	root.append(createRibbon(document), viewport);
	const feedback: string[] = [];
	const commands = new ViewerCommands({
		root,
		viewport,
		controller,
		rulers: createRulers(viewport),
		fit: () => {},
		togglePane: () => {},
		reveal: () => {},
		focusSearch: () => {},
		togglePanZoom: () => {},
		announce: (message) => feedback.push(message),
	});
	const dispose = commands.wire();
	controller.subscribe((state) => commands.render(state));
	const button = (id: string) =>
		root.querySelector<HTMLElement & { disabled: boolean }>(`[command="${id}"]`)!;
	const press = (id: string) =>
		button(id).shadowRoot!.querySelector<HTMLButtonElement>('button')!.click();
	const combo = (id: string) =>
		root.querySelector<HTMLElement & { value: string; disabled: boolean }>(`[data-combo="${id}"]`)!;
	const select = (id: string, value: string) => {
		combo(id).value = value;
		combo(id).dispatchEvent(new Event('change', { bubbles: true }));
	};
	const done = () => vi.waitFor(() => expect(controller.state.edit.busy).toBe(false));
	const shape = () => controller.state.document!.pages[0]!.shapes[0]!;
	const selection = () => controller.selectShape({ id: '1', name: 'Import test', pageId: '1' });
	return {
		root,
		viewport,
		controller,
		commands,
		edits,
		bytes,
		feedback,
		button,
		press,
		combo,
		select,
		done,
		shape,
		selection,
		dispose,
	};
}

it('routes real ribbon styles and alignment through source history and restores saved bytes', async () => {
	const ui = await setup();
	expect(ui.button('bold').disabled).toBe(true);
	expect(ui.button('bold').title).toMatch(/Select a shape/);
	ui.selection();
	expect(ui.button('bold').disabled).toBe(false);
	ui.press('bold');
	expect(ui.button('italic').disabled).toBe(true);
	expect(ui.combo('font-size').disabled).toBe(true);
	ui.press('italic');
	await ui.done();
	expect(ui.edits).toHaveLength(1);
	expect(ui.edits[0]).toEqual([{ type: 'format-text', pageId: '1', shapeId: '1', bold: true }]);
	expect(ui.shape().text.runs[0]!.bold).toBe(true);
	expect(ui.button('bold').getAttribute('pressed')).toBe('true');
	expect(
		(await parseVsdx(ui.controller.exportVsdx().bytes)).pages[0]!.shapes[0]!.text.runs[0]!.bold,
	).toBe(true);
	ui.press('undo');
	await ui.done();
	expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
	expect(ui.button('bold').getAttribute('pressed')).toBe('false');
	ui.press('redo');
	await ui.done();
	expect(ui.shape().text.runs[0]!.bold).toBe(true);
	for (const id of ['italic', 'underline', 'align-right', 'align-top']) {
		ui.press(id);
		await ui.done();
		expect(ui.button(id).getAttribute('pressed')).toBe('true');
	}
	expect(ui.shape().text.horizontalAlign).toBe('right');
	expect(ui.shape().text.verticalAlign).toBe('top');
	expect(ui.feedback).toContain('Updated shape formatting.');
	ui.dispose();
	ui.controller.destroy();
});

it('uses shared font combos, size stepping, fill colors and line weights', async () => {
	const ui = await setup();
	ui.selection();
	ui.select('font-size', '24');
	await ui.done();
	expect(ui.shape().text.fontSize * 72).toBeCloseTo(24);
	expect(ui.combo('font-size').value).toBe('24');
	ui.press('grow-font');
	await ui.done();
	expect(ui.combo('font-size').value).toBe('28');
	ui.press('shrink-font');
	await ui.done();
	expect(ui.combo('font-size').value).toBe('24');
	ui.select('font', 'Calibri');
	await ui.done();
	expect(ui.shape().text.fontFamily).toBe('Calibri');
	ui.press('fill-red');
	await ui.done();
	expect(ui.shape().style.fill).toBe('#ff0000');
	ui.press('line-blue');
	await ui.done();
	expect(ui.shape().style.lineColor).toBe('#4472c4');
	ui.press('line-weight-3');
	await ui.done();
	expect(ui.shape().style.lineWidth * 72).toBeCloseTo(3);
	expect(ui.button('line-weight-3').getAttribute('checked')).toBe('true');
	ui.press('fill-none');
	await ui.done();
	expect(ui.shape().style.fill).toBe('none');
	ui.dispose();
	ui.controller.destroy();
});

it('disables model-only, mismatched page and unsupported selections and ignores forged events', async () => {
	const ui = await setup(false);
	ui.selection();
	expect(ui.button('bold').disabled).toBe(true);
	expect(ui.combo('font').disabled).toBe(true);
	expect(ui.button('bold').title).toMatch(/Open a .vsdx file/);
	ui.button('bold').dispatchEvent(
		new CustomEvent('office-command', { bubbles: true, detail: { command: 'bold' } }),
	);
	ui.commands.run({ type: 'text-toggle', property: 'bold' });
	expect(ui.edits).toHaveLength(0);
	await ui.controller.load(ui.bytes);
	ui.selection();
	ui.shape().masterId = '42';
	ui.commands.render(ui.controller.state);
	expect(ui.button('bold').disabled).toBe(true);
	expect(ui.button('fill-red').disabled).toBe(true);
	delete ui.shape().masterId;
	ui.controller.selectShape({ id: '1', name: 'Import test', pageId: 'another-page' });
	expect(ui.button('bold').disabled).toBe(true);
	ui.dispose();
	ui.controller.destroy();
});

it('shares Ctrl+B/I/U routing and preserves text-field editing shortcuts', async () => {
	const ui = await setup();
	ui.selection();
	const key = (target: Element, name: string) =>
		target.dispatchEvent(
			new KeyboardEvent('keydown', {
				key: name,
				ctrlKey: true,
				bubbles: true,
				composed: true,
				cancelable: true,
			}),
		);
	const textarea = document.createElement('textarea');
	ui.viewport.append(textarea);
	expect(key(textarea, 'b')).toBe(true);
	expect(key(ui.combo('font-size'), 'b')).toBe(true);
	expect(ui.edits).toHaveLength(0);
	for (const name of ['b', 'i', 'u']) {
		expect(key(ui.viewport, name)).toBe(false);
		await ui.done();
	}
	expect(ui.shape().text.runs[0]).toMatchObject({ bold: true, italic: true, underline: true });
	ui.dispose();
	ui.controller.destroy();
});

it('rejects invalid formatting without altering bytes and displays the controller error', async () => {
	const ui = await setup();
	ui.selection();
	ui.commands.run({ type: 'font-family', value: 'Missing family' });
	await ui.done();
	expect(ui.controller.state.edit.error).toBeDefined();
	expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
	expect(ui.controller.state.edit.canUndo).toBe(false);
	ui.dispose();
	ui.controller.destroy();
});

it('reorders source-backed shapes with four menu actions and disables edge commands', async () => {
	const ui = await setup();
	await ui.controller.applyEdits([
		{ type: 'create-rectangle', pageId: '1', shapeId: '2', x: 1, y: 1, width: 1, height: 1 },
		{ type: 'create-rectangle', pageId: '1', shapeId: '3', x: 2, y: 2, width: 1, height: 1 },
	]);
	ui.selection();
	const ids = () => ui.controller.state.document!.pages[0]!.shapes.map((shape) => shape.id);
	expect(ui.button('send-to-back').disabled).toBe(true);
	ui.press('bring-to-front');
	await ui.done();
	expect(ids()).toEqual(['2', '3', '1']);
	expect(ui.button('bring-to-front').disabled).toBe(true);
	ui.press('send-to-back');
	await ui.done();
	expect(ids()).toEqual(['1', '2', '3']);
	ui.press('bring-forward');
	await ui.done();
	expect(ids()).toEqual(['2', '1', '3']);
	ui.press('send-backward');
	await ui.done();
	expect(ids()).toEqual(['1', '2', '3']);
	expect(ui.feedback).toContain('Updated shape order.');
	ui.dispose();
	ui.controller.destroy();
});

it('keeps size step direction outside the standard font-size menu', async () => {
	const ui = await setup();
	ui.selection();
	ui.commands.run({ type: 'font-size', value: 100 });
	await ui.done();
	ui.press('grow-font');
	await ui.done();
	expect(ui.shape().text.fontSize * 72).toBeGreaterThan(100);
	ui.commands.run({ type: 'font-size', value: 4 });
	await ui.done();
	ui.press('shrink-font');
	await ui.done();
	expect(ui.shape().text.fontSize * 72).toBeLessThan(4);
	ui.dispose();
	ui.controller.destroy();
});

it('allows fill and line changes on mixed text without enabling text formatting', async () => {
	const ui = await setup(true, true);
	ui.selection();
	expect(ui.shape().text.runs.map((run) => run.bold)).toEqual([false, true]);
	expect(ui.button('bold').disabled).toBe(true);
	expect(ui.button('bold').title).toMatch(/uniform text styles/);
	expect(ui.button('fill-blue').disabled).toBe(false);
	expect(ui.button('fill-blue').title).toMatch(/source formulas and protection/);
	ui.press('fill-blue');
	await ui.done();
	expect(ui.shape().style.fill).toBe('#4472c4');
	ui.press('line-red');
	await ui.done();
	expect(ui.shape().style.lineColor).toBe('#ff0000');
	expect(ui.shape().text.runs.map((run) => run.bold)).toEqual([false, true]);
	ui.dispose();
	ui.controller.destroy();
});
