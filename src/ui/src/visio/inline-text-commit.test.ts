import { afterEach, expect, it, vi } from 'vitest';
import { editVsdx, parseVsdx, type VisioEdit } from 'ooxml-core/visio';
import { ViewerController } from './controller';
import { mountDiagramCanvas } from './building-blocks';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';

afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

async function surface(options: { gate?: Promise<void>; fail?: Error } = {}) {
	const bytes = await createVsdxFixture('Movable');
	const edits: VisioEdit[][] = [];
	const controller = new ViewerController(
		parseVsdx,
		vi.fn(),
		async (source, commands) => {
			edits.push([...commands]);
			await options.gate;
			if (options.fail) throw options.fail;
			const result = await editVsdx(source, commands);
			return { ...result, document: await parseVsdx(result.bytes) };
		},
		async () => 'unused',
	);
	controller.canEdit = true;
	await controller.load(bytes);
	vi.spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (this: Element) {
		const [width, height] = this.hasAttribute('part') ? [800, 600] : [100, 40];
		return { left: 0, top: 0, right: width, bottom: height, x: 0, y: 0, width, height } as DOMRect;
	});
	const host = document.createElement('div');
	document.body.append(host);
	const stop = mountDiagramCanvas(host, controller);
	const viewport = host.shadowRoot!.querySelector<HTMLElement>('[part="canvas"]')!;
	const group = viewport.querySelector<SVGGElement>('[data-shape-id="1"]')!;
	controller.selectShapes([{ id: '1', name: group.dataset.shapeName ?? '', pageId: '1' }]);
	const type = (text: string) => {
		viewport.dispatchEvent(new KeyboardEvent('keydown', { key: 'x', bubbles: true }));
		const input = viewport.querySelector<HTMLTextAreaElement>('textarea.inline-text')!;
		input.value = text;
		return input;
	};
	return { controller, viewport, edits, stop, type };
}

it('reports no draft when nothing is being typed', async () => {
	const ui = await surface();
	const outcome = await ui.controller.commitTextEdit();
	expect(outcome).toMatchObject({ status: 'unchanged', reason: 'no-draft' });
	expect(outcome.sourceGeneration).toBe(ui.controller.sourceGeneration);
	expect(ui.edits).toEqual([]);
	ui.stop();
	ui.controller.destroy();
});

it('commits the visible draft as one edit with page, shape and generation', async () => {
	const ui = await surface();
	ui.type('Saved text');
	const generation = ui.controller.sourceGeneration;
	const outcome = await ui.controller.commitTextEdit();
	expect(outcome).toMatchObject({
		status: 'committed',
		text: 'Saved text',
		pageId: '1',
		shapeId: '1',
		sourceGeneration: generation,
	});
	expect(ui.edits).toHaveLength(1);
	expect(ui.edits[0]![0]).toMatchObject({ type: 'replace-plain-text', text: 'Saved text' });
	expect(
		ui.controller.state.document?.pages[0]?.shapes.find((s) => s.id === '1')?.text,
	).toBeTruthy();
	expect(await ui.controller.commitTextEdit()).toMatchObject({ status: 'unchanged' });
	ui.stop();
	ui.controller.destroy();
});

it('is unchanged when the draft equals the shape text', async () => {
	const ui = await surface();
	const input = ui.type('x');
	input.value = ui.controller.state.document!.pages[0]!.shapes.find(
		(s) => s.id === '1',
	)!.text.plainText;
	expect(await ui.controller.commitTextEdit()).toMatchObject({
		status: 'unchanged',
		reason: 'same-text',
		shapeId: '1',
	});
	expect(ui.edits).toEqual([]);
	ui.stop();
	ui.controller.destroy();
});

it('waits for a commit already in flight instead of repeating it', async () => {
	let release!: () => void;
	const gate = new Promise<void>((resolve) => (release = resolve));
	const ui = await surface({ gate });
	const input = ui.type('Slow');
	// Esc starts the commit; Save asks while it is still running.
	input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
	let settled = false;
	const pending = ui.controller.commitTextEdit().then((outcome) => {
		settled = true;
		return outcome;
	});
	await new Promise((resolve) => setTimeout(resolve, 20));
	expect(settled).toBe(false);
	release();
	expect(await pending).toMatchObject({ status: 'committed', text: 'Slow' });
	expect(ui.edits).toHaveLength(1);
	ui.stop();
	ui.controller.destroy();
});

it('returns the rejected text when the editor refuses the draft', async () => {
	const ui = await surface({ fail: new Error('Text is protected.') });
	ui.type('Nope');
	const outcome = await ui.controller.commitTextEdit();
	expect(outcome).toMatchObject({
		status: 'refused',
		reason: 'rejected',
		message: 'Text is protected.',
		text: 'Nope',
		shapeId: '1',
	});
	ui.stop();
	ui.controller.destroy();
});

it('refuses with the draft when permission was withdrawn', async () => {
	const ui = await surface();
	ui.type('Late');
	ui.controller.canEdit = false;
	const outcome = await ui.controller.commitTextEdit();
	expect(outcome.status === 'refused' || outcome.status === 'unchanged').toBe(true);
	expect(ui.edits).toEqual([]);
	ui.stop();
	ui.controller.destroy();
});

it('never applies an old draft to another document', async () => {
	const ui = await surface();
	ui.type('Old draft');
	const before = ui.controller.sourceGeneration;
	await ui.controller.load(await createVsdxFixture('Other'));
	expect(ui.controller.sourceGeneration).not.toBe(before);
	const outcome = await ui.controller.commitTextEdit();
	// The dropped draft is reported once, with its old identity, and never applied.
	expect(outcome).toMatchObject({
		status: 'cancelled',
		reason: 'stale-document',
		text: 'Old draft',
		pageId: '1',
		shapeId: '1',
		sourceGeneration: before,
	});
	expect(await ui.controller.commitTextEdit()).toMatchObject({
		status: 'unchanged',
		reason: 'no-draft',
	});
	expect(ui.edits).toEqual([]);
	ui.stop();
	ui.controller.destroy();
});

it('drops the draft on unmount and after destroy', async () => {
	const ui = await surface();
	ui.type('Gone');
	ui.stop();
	expect(await ui.controller.commitTextEdit()).toMatchObject({ status: 'unchanged' });
	expect(ui.edits).toEqual([]);
	ui.controller.destroy();
	expect(await ui.controller.commitTextEdit()).toMatchObject({
		status: 'cancelled',
		reason: 'destroyed',
	});
});
