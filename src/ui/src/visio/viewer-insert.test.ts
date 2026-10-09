import { afterEach, expect, it, vi } from 'vitest';
import { parseVsdx } from 'ooxml-core/visio';
import { setupFormattingViewer } from './__fixtures__/formatting-viewer';
import { createContextMenus } from './viewer-context-menu';

const png = Uint8Array.from(
	atob(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAABHNCSVQICAgIfAhkiAAAAAFzUkdCAK7OHOkAAAALSURBVAiZY2AAAgAABQABYlUyiAAAAABJRU5ErkJggg==',
	),
	(char) => char.charCodeAt(0),
);
afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});
const setup = async (source = true) => {
	const view = await setupFormattingViewer(source);
	view.root.append(...createContextMenus(document));
	const dialog = (name: string) =>
		view.root.querySelector<HTMLElement & { open: boolean }>(`.${name}`)!;
	const field = (name: string, value: string) => {
		const input = view.root.querySelector<HTMLInputElement | HTMLSelectElement>(
			`[name="${name}"]`,
		)!;
		input.value = value;
	};
	const dialogButton = (name: string, label: string) =>
		dialog(name)
			.querySelector(`[label="${label}"]`)!
			.shadowRoot!.querySelector<HTMLButtonElement>('button')!
			.click();
	return { ...view, dialog, field, dialogButton };
};
const choose = (input: HTMLInputElement, file: File) => {
	Object.defineProperty(input, 'files', { configurable: true, value: [file] });
	input.dispatchEvent(new Event('change'));
};

it('enables Pictures for editable drawings and Link/ScreenTip only for one selected shape', async () => {
	const view = await setup();
	expect(view.button('pictures').disabled).toBe(false);
	expect(view.button('link').disabled).toBe(true);
	expect(view.button('screen-tip').disabled).toBe(true);
	expect(view.button('field').disabled).toBe(true);
	expect(view.button('chart').title).toMatch(/Excel charts/);
	view.selection();
	expect(view.button('link').disabled).toBe(false);
	expect(view.button('screen-tip').disabled).toBe(false);
	expect(view.button('ctx-hyperlink').disabled).toBe(false);
	view.dispose();
	const readOnly = await setup(false);
	expect(readOnly.button('pictures').disabled).toBe(true);
	readOnly.dispose();
});

it('inserts a local picture at the page centre as a selected, undoable Foreign shape', async () => {
	const view = await setup();
	const input = view.root.querySelector<HTMLInputElement>('[data-insert-picture]')!;
	const click = vi.spyOn(input, 'click').mockImplementation(() => {});
	view.press('pictures');
	expect(click).toHaveBeenCalled();
	choose(input, new File([png], 'dot.png', { type: 'image/png' }));
	await vi.waitFor(() => expect(view.edits.at(-1)?.[0]?.type).toBe('insert-picture'));
	await view.done();
	const page = view.controller.state.document!.pages[0]!;
	const picture = page.shapes.at(-1)!;
	expect(picture).toMatchObject({ kind: 'foreign' });
	expect(picture.image?.mimeType).toBe('image/png');
	expect(picture.rotation).toMatchObject({ pinX: page.width / 2, pinY: page.height / 2 });
	expect(view.controller.state.selectedShapes.map((shape) => shape.id)).toEqual([picture.id]);
	expect(view.feedback.at(-1)).toBe('Inserted picture dot.png.');
	await view.controller.undo();
	expect(view.controller.state.document!.pages[0]!.shapes).toHaveLength(1);
	choose(input, new File(['<svg/>'], 'vector.svg', { type: 'image/svg+xml' }));
	await vi.waitFor(() => expect(view.feedback.at(-1)).toMatch(/vector\.svg cannot be inserted/));
	view.dispose();
});

it('adds, edits and removes a link through the dialog, with Ctrl+K and undo', async () => {
	const view = await setup();
	view.selection();
	view.viewport.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'k', ctrlKey: true, bubbles: true }),
	);
	expect(view.dialog('link-dialog').open).toBe(true);
	view.dialogButton('link-dialog', 'OK');
	expect(view.root.querySelector('.link-dialog [role="alert"]')!.textContent).toMatch(/address/);
	view.field('address', 'example.com/docs');
	view.field('description', 'Docs');
	view.dialogButton('link-dialog', 'OK');
	await vi.waitFor(() => expect(view.dialog('link-dialog').open).toBe(false));
	expect(view.shape().hyperlinks).toEqual([
		expect.objectContaining({
			name: 'Row_1',
			description: 'Docs',
			target: { kind: 'external', href: 'https://example.com/docs' },
		}),
	]);
	view.press('link');
	expect((view.root.querySelector('[name="address"]') as HTMLInputElement).value).toBe(
		'https://example.com/docs',
	);
	const pageName = view.controller.state.document!.pages[0]!.name;
	view.field('subAddress', pageName);
	view.field('address', '');
	view.dialogButton('link-dialog', 'OK');
	await vi.waitFor(() =>
		expect(view.shape().hyperlinks?.[0]?.target).toEqual({
			kind: 'internal',
			subAddress: pageName,
		}),
	);
	view.press('link');
	view.dialogButton('link-dialog', 'Remove Link');
	await vi.waitFor(() => expect(view.shape().hyperlinks).toEqual([]));
	await view.controller.undo();
	expect(view.shape().hyperlinks).toHaveLength(1);
	const saved = await parseVsdx(view.controller.exportVsdx().bytes);
	expect(saved.pages[0]!.shapes[0]!.hyperlinks).toHaveLength(1);
	view.dispose();
});

it('sets the ScreenTip and follows links with Ctrl+click without navigating the app', async () => {
	const view = await setup();
	view.selection();
	view.press('screen-tip');
	view.field('text', 'Hover help');
	view.dialogButton('screen-tip-dialog', 'OK');
	await vi.waitFor(() => expect(view.shape().screenTip).toBe('Hover help'));
	await view.controller.applyEdits([
		{
			type: 'set-shape-hyperlink',
			pageId: '1',
			shapeId: '1',
			hyperlink: { address: 'https://example.com/', subAddress: '', description: '' },
		},
	]);
	const open = vi.spyOn(window, 'open').mockImplementation(() => null);
	const group = document.createElementNS('http://www.w3.org/2000/svg', 'g');
	group.dataset.shapeId = '1';
	group.dataset.pageId = '1';
	view.viewport.append(group);
	const selected = vi.fn();
	view.viewport.addEventListener('click', selected);
	group.dispatchEvent(new MouseEvent('click', { bubbles: true, ctrlKey: true }));
	expect(open).toHaveBeenCalledWith('https://example.com/', '_blank', 'noopener,noreferrer');
	expect(selected).not.toHaveBeenCalled();
	group.dispatchEvent(new MouseEvent('click', { bubbles: true }));
	expect(open).toHaveBeenCalledTimes(1);
	expect(selected).toHaveBeenCalledTimes(1);
	view.dispose();
});
