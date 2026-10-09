import { afterEach, expect, it } from 'vitest';
import type { OfficeUiGallery } from '../ribbon/gallery';
import { setupFormattingViewer as setup } from './__fixtures__/formatting-viewer';
import { BASIC_SHAPES } from './shapes-window';

afterEach(() => document.body.replaceChildren());

const gallery = (root: ShadowRoot) =>
	root.querySelector<OfficeUiGallery>('office-ui-gallery[command="change-shape"]')!;
const pick = (root: ShadowRoot, id: string) => {
	gallery(root).querySelector<HTMLButtonElement>('.trigger')!.click();
	gallery(root).querySelector<HTMLButtonElement>(`[data-gallery-item="${id}"]`)!.click();
};

it('offers every Basic Shapes master and explains why it is disabled', async () => {
	const ui = await setup();
	const element = gallery(ui.root);
	expect(element.state!.sections[0]!.items.map((item) => item.id)).toEqual(
		BASIC_SHAPES.map((master) => master.id),
	);
	expect(element.disabled || element.state!.disabled).toBe(true);
	expect(element.getAttribute('title')).toBe('Change Shape: Select exactly one shape.');
	expect(element.querySelector('.trigger')!.getAttribute('title')).toMatch(/Select exactly one/);
	ui.selection();
	expect(element.state!.disabled).toBe(false);
	expect(element.hasAttribute('disabled')).toBe(false);
	expect(element.getAttribute('title')).toBe('Change Shape');
	ui.dispose();
	ui.controller.destroy();
});

it('changes the selected shape through source history, with undo and redo', async () => {
	const ui = await setup();
	ui.selection();
	const before = ui.shape();
	pick(ui.root, 'ellipse');
	await ui.done();
	expect(ui.edits.at(-1)).toEqual([
		{ type: 'change-shape', pageId: '1', shapeId: '1', shape: 'ellipse' },
	]);
	const after = ui.shape();
	expect(after.geometry).toHaveLength(1);
	expect(after.geometry[0]!.path).toContain(' A ');
	expect([after.width, after.height, after.transform]).toEqual([
		before.width,
		before.height,
		before.transform,
	]);
	expect(after.text.plainText).toBe(before.text.plainText);
	expect(ui.feedback).toContain('Changed the shape to Ellipse.');
	ui.press('undo');
	await ui.done();
	expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
	ui.press('redo');
	await ui.done();
	expect(ui.shape().geometry[0]!.path).toContain(' A ');
	ui.commands.run({ type: 'change-shape', shape: 'triangle' });
	await ui.done();
	expect(ui.shape().geometry[0]!.path).not.toContain(' A ');
	ui.dispose();
	ui.controller.destroy();
});

it('stays disabled for read-only documents and multiple selections', async () => {
	const readOnly = await setup(false);
	readOnly.selection();
	expect(gallery(readOnly.root).getAttribute('title')).toMatch(/Open a \.vsdx file/);
	readOnly.commands.run({ type: 'change-shape', shape: 'star' });
	expect(readOnly.edits).toHaveLength(0);
	readOnly.dispose();
	readOnly.controller.destroy();
	const ui = await setup();
	await ui.controller.applyEdits([
		{ type: 'create-rectangle', pageId: '1', shapeId: '2', x: 1, y: 1, width: 1, height: 1 },
	]);
	ui.controller.selectAll();
	expect(gallery(ui.root).state!.disabled).toBe(true);
	const count = ui.edits.length;
	ui.commands.run({ type: 'change-shape', shape: 'star' });
	expect(ui.edits).toHaveLength(count);
	ui.dispose();
	ui.controller.destroy();
});
