import { afterEach, expect, it, vi } from 'vitest';
import { pointerViewer } from './__fixtures__/pointer-viewer';
import { ViewerTextTool } from './viewer-text-tool';
import { ShapeDrawTool } from './viewer-draw-tool';
import { renderPage } from './render-svg';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});
async function setup() {
	const viewer = await pointerViewer();
	viewer.inactive();
	let active = true;
	const reveal = vi.fn(),
		feedback: string[] = [];
	const tool = new ViewerTextTool(viewer.viewport, viewer.controller, {
		active: () => active,
		announce: (message) => feedback.push(message),
		revealEdit: reveal,
	});
	const dispose = tool.wire(),
		unsubscribe = viewer.controller.subscribe((state) => tool.render(state));
	const draft = () =>
		viewer.viewport.querySelector<HTMLTextAreaElement>('[data-text-draft] textarea')!;
	const draw = async () => {
		viewer.pointer('pointerdown', viewer.svg, 10, 10);
		viewer.pointer('pointermove', viewer.svg, 30, 20);
		viewer.pointer('pointerup', viewer.svg, 30, 20);
		await vi.waitFor(() => expect(draft()).not.toBeNull());
	};
	return {
		...viewer,
		draft,
		draw,
		feedback,
		reveal,
		tool,
		inactive() {
			active = false;
			tool.render(viewer.controller.state);
		},
		dispose() {
			unsubscribe();
			dispose();
			viewer.dispose();
		},
	};
}
it('keeps draft/source separate and creates one logical multiline text box with exact history', async () => {
	const ui = await setup();
	try {
		await ui.draw();
		expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
		expect(ui.edits).toEqual([]);
		const text = ' <&> Ω\n第二行\n';
		ui.draft().value = text;
		ui.draft().dispatchEvent(new KeyboardEvent('keydown', { key: '0', bubbles: true }));
		expect(ui.controller.state.zoom).toBe(1);
		ui.draft().click();
		expect(ui.draft()).not.toBeNull();
		ui.draft().dispatchEvent(
			new KeyboardEvent('keydown', {
				key: 'Enter',
				ctrlKey: true,
				bubbles: true,
				isComposing: true,
			}),
		);
		expect(ui.edits).toEqual([]);
		ui.draft().dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true }),
		);
		await vi.waitFor(() => expect(ui.feedback).toEqual(['Text box 3 added.']));
		const state = ui.controller.state,
			shape = state.document!.pages[0]!.shapes[2]!;
		expect(shape.text.plainText).toBe(text);
		expect([shape.width, shape.height]).toEqual([2, 1]);
		expect(shape.rotation).toMatchObject({ pinX: 2, pinY: 9.5 });
		expect(shape.style).toMatchObject({ fill: 'none', linePattern: 0 });
		expect(state.selectedShapes.map((selection) => selection.id)).toEqual(['3']);
		expect(ui.edits).toHaveLength(1);
		const accepted = ui.controller.exportVsdx().bytes;
		const interactive = renderPage(state.document!, state.document!.pages[0]!);
		expect(interactive.svg.querySelector('[data-shape-id="3"] [data-text-hit]')).not.toBeNull();
		interactive.dispose();
		const portable = renderPage(state.document!, state.document!.pages[0]!, { static: true });
		expect(portable.svg.querySelector('[data-text-hit]')).toBeNull();
		portable.dispose();
		await ui.controller.undo();
		expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
		await ui.controller.redo();
		expect(ui.controller.exportVsdx().bytes).toEqual(accepted);
		expect(ui.controller.state.selectedShapes.map((selection) => selection.id)).toEqual(['3']);
	} finally {
		ui.dispose();
	}
});
it.each(['Escape', 'selection', 'source', 'zoom', 'tool'])(
	'cancels a text draft on %s without writing',
	async (reason) => {
		const ui = await setup();
		try {
			await ui.draw();
			ui.draft().value = 'Uncommitted';
			if (reason === 'Escape')
				ui.draft().dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
			else if (reason === 'selection') ui.select();
			else if (reason === 'source') ui.controller.setDocument(null);
			else if (reason === 'zoom') ui.controller.setZoom(2);
			else ui.inactive();
			expect(ui.draft()).toBeNull();
			expect(ui.svg.querySelector('.text-box-draft-frame')).toBeNull();
			expect(ui.edits).toEqual([]);
		} finally {
			ui.dispose();
		}
	},
);
it('opens existing current-page text editing and ignores background IDs', async () => {
	const ui = await setup();
	try {
		ui.pointer('pointerdown', ui.group());
		expect(ui.reveal).toHaveBeenCalledTimes(1);
		expect(ui.controller.state.selectedShape?.id).toBe('1');
		ui.group().dataset.pageId = 'background';
		ui.pointer('pointerdown', ui.group());
		expect(ui.reveal).toHaveBeenCalledTimes(1);
		expect(ui.draft()).toBeNull();
		expect(ui.edits).toEqual([]);
	} finally {
		ui.dispose();
	}
});
it('preserves intentional whitespace and reports planning refusal without source edits', async () => {
	const ui = await setup();
	try {
		await ui.draw();
		ui.draft().value = '  \n\n';
		ui.draft().dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true }),
		);
		await vi.waitFor(() => expect(ui.feedback).toEqual(['Text box 3 added.']));
		expect(ui.controller.state.document!.pages[0]!.shapes[2]!.text.plainText).toBe('  \n\n');
	} finally {
		ui.dispose();
	}
	const refused = await setup();
	try {
		await refused.draw();
		refused.controller.state.document!.pages[0]!.shapes[0]!.id = '4294967295';
		refused.draft().value = 'Keep me';
		refused
			.draft()
			.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true }));
		await vi.waitFor(() => expect(refused.feedback[0]).toMatch(/IDs/));
		expect(refused.edits).toEqual([]);
		expect(refused.controller.exportVsdx().bytes).toEqual(refused.bytes);
		expect(refused.controller.state.edit.canUndo).toBe(false);
	} finally {
		refused.dispose();
	}
});
it('guards ordinary drawing intent before release and commits creation once when current', async () => {
	const ui = await pointerViewer();
	ui.inactive();
	const feedback: string[] = [];
	const tool = new ShapeDrawTool(ui.viewport, ui.controller, {
		tool: () => 'ellipse',
		announce: (message) => feedback.push(message),
	});
	const dispose = tool.wire(),
		unsubscribe = ui.controller.subscribe((state) => tool.render(state));
	try {
		ui.pointer('pointerdown', ui.svg, 10, 10);
		ui.pointer('pointermove', ui.svg, 30, 20);
		ui.select();
		expect(ui.svg.querySelector('.draw-preview')).toBeNull();
		ui.pointer('pointerup', ui.svg, 30, 20);
		expect(ui.edits).toEqual([]);
		ui.pointer('pointerdown', ui.svg, 10, 10);
		ui.pointer('pointermove', ui.svg, 30, 20);
		ui.pointer('pointerup', ui.svg, 30, 20);
		await vi.waitFor(() => expect(feedback).toEqual(['Shape 3 added.']));
		expect(ui.edits).toHaveLength(1);
		expect(ui.controller.state.selectedShapes.map((shape) => shape.id)).toEqual(['3']);
		await ui.controller.undo();
		expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
		expect(ui.controller.state.selectedShape?.id).toBe('1');
	} finally {
		unsubscribe();
		dispose();
		ui.dispose();
	}
});
