import { afterEach, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { editVsdx, parseVsdx, type VisioEdit } from 'ooxml-core/visio';
import { ViewerController } from './controller';
import { diagramToolbarProps, mountDiagramCanvas, mountDiagramShapes } from './building-blocks';
import { DIAGRAM_TOOLS } from './diagram-tool';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
	vi.restoreAllMocks();
});
async function surface(canEdit: boolean) {
	const zip = await JSZip.loadAsync(await createVsdxFixture('Movable'));
	const page = await zip.file('visio/pages/page1.xml')!.async('string');
	// Formulas that make the shape resizable, as the full viewer's own tests do.
	zip.file(
		'visio/pages/page1.xml',
		page
			.replaceAll('N="X" V="3"/>', 'N="X" V="3" F="Width"/>')
			.replaceAll('N="Y" V="1"/>', 'N="Y" V="1" F="Height"/>'),
	);
	const bytes = await zip.generateAsync({ type: 'uint8array' });
	const edits: VisioEdit[][] = [];
	const controller = new ViewerController(
		parseVsdx,
		vi.fn(),
		async (source, commands) => {
			edits.push([...commands]);
			const result = await editVsdx(source, commands);
			return { ...result, document: await parseVsdx(result.bytes) };
		},
		async () => 'unused',
	);
	controller.canEdit = canEdit;
	await controller.load(bytes);
	const host = document.createElement('div');
	document.body.append(host);
	const stop = mountDiagramCanvas(host, controller);
	const viewport = host.shadowRoot!.querySelector<HTMLElement>('[part="canvas"]')!;
	const svg = viewport.querySelector<SVGSVGElement>('svg.paper')!;
	Object.defineProperty(svg, 'getScreenCTM', {
		value: () => ({ a: 10, b: 0, inverse: () => ({}) }),
		configurable: true,
	});
	vi.stubGlobal(
		'DOMPoint',
		class {
			constructor(
				public x: number,
				public y: number,
			) {}
			matrixTransform() {
				return { x: this.x / 10, y: this.y / 10 };
			}
		},
	);
	const pointer = (type: string, x: number, y: number, target: Element = svg) => {
		const event = new Event(type, { bubbles: true, cancelable: true });
		for (const [key, value] of Object.entries({ pointerId: 7, button: 0, clientX: x, clientY: y }))
			Object.defineProperty(event, key, { value });
		target.dispatchEvent(event);
	};
	const select = () => {
		const group = svg.querySelector<SVGGElement>('[data-shape-id="1"]')!;
		controller.selectShapes([{ id: '1', name: group.dataset.shapeName ?? '', pageId: '1' }]);
	};
	return {
		controller,
		viewport,
		edits,
		pointer,
		select,
		dispose() {
			stop();
			controller.destroy();
		},
	};
}

/** A press on a rotation handle starts a gesture that hides the saved resize overlay. */
function rotationStarts(ui: Awaited<ReturnType<typeof surface>>): boolean {
	const svg = ui.viewport.querySelector<SVGSVGElement>('svg.paper')!;
	const ns = 'http://www.w3.org/2000/svg';
	const overlay = document.createElementNS(ns, 'g');
	overlay.dataset.resizeOverlay = '';
	const handle = document.createElementNS(ns, 'circle');
	handle.dataset.rotationHandle = '1';
	svg.append(overlay, handle);
	ui.pointer('pointerdown', 40, 30, handle);
	const started = overlay.style.visibility === 'hidden';
	ui.viewport.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
	overlay.remove();
	handle.remove();
	return started;
}

it('draws a shape with the selected tool and rejects drawing while read-only', async () => {
	const editable = await surface(true);
	try {
		editable.controller.setTool('rectangle');
		expect(editable.controller.state.tool).toBe('rectangle');
		editable.pointer('pointerdown', 10, 10);
		editable.pointer('pointermove', 30, 20);
		editable.pointer('pointerup', 30, 20);
		await vi.waitFor(() => expect(editable.edits).toHaveLength(1));
		expect(editable.edits[0]![0]).toMatchObject({ type: 'create-rectangle' });
		// Losing permission drops the tool back to the pointer.
		editable.controller.canEdit = false;
		expect(editable.controller.state.tool).toBe('pointer');
	} finally {
		editable.dispose();
	}
	const readOnly = await surface(false);
	try {
		readOnly.controller.setTool('rectangle');
		expect(readOnly.controller.state.tool).toBe('pointer');
		readOnly.pointer('pointerdown', 10, 10);
		readOnly.pointer('pointermove', 30, 20);
		readOnly.pointer('pointerup', 30, 20);
		await Promise.resolve();
		expect(readOnly.edits).toEqual([]);
	} finally {
		readOnly.dispose();
	}
});

it('creates a text box with the text tool and a connector with the connector tool, only when editing', async () => {
	const ui = await surface(true);
	try {
		ui.controller.setTool('text');
		expect(ui.controller.state.tool).toBe('text');
		ui.pointer('pointerdown', 100, 100);
		ui.pointer('pointermove', 140, 120);
		ui.pointer('pointerup', 140, 120);
		const draft = await vi.waitFor(() => {
			const input = ui.viewport.querySelector<HTMLTextAreaElement>('[data-text-draft] textarea');
			expect(input).not.toBeNull();
			return input!;
		});
		draft.value = 'Hello';
		draft.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true }),
		);
		await vi.waitFor(() => expect(ui.edits).toHaveLength(1));
		expect(ui.edits[0]![0]).toMatchObject({ type: 'create-text-box' });
		ui.controller.setTool('connector');
		ui.pointer('pointerdown', 200, 200);
		ui.pointer('pointermove', 260, 240);
		ui.pointer('pointerup', 260, 240);
		await vi.waitFor(() => expect(ui.edits).toHaveLength(2));
		expect(ui.edits[1]![0]).toMatchObject({ type: 'create-line', route: 'right-angle' });
		// Escape abandons the tool and returns to the pointer.
		ui.viewport.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(ui.controller.state.tool).toBe('pointer');
	} finally {
		ui.dispose();
	}
	const readOnly = await surface(false);
	try {
		for (const tool of ['text', 'connector'] as const) {
			readOnly.controller.setTool(tool);
			expect(readOnly.controller.state.tool).toBe('pointer');
			readOnly.pointer('pointerdown', 100, 100);
			readOnly.pointer('pointermove', 140, 120);
			readOnly.pointer('pointerup', 140, 120);
		}
		await Promise.resolve();
		expect(readOnly.edits).toEqual([]);
		expect(readOnly.viewport.querySelector('[data-text-draft]')).toBeNull();
	} finally {
		readOnly.dispose();
	}
});

it('offers resize and rotate handles and inline text only when editing is allowed', async () => {
	const editable = await surface(true);
	try {
		// jsdom has no layout: give the canvas room and the shape a box inside it.
		vi.spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (
			this: Element,
		) {
			const [left, top, right, bottom] = this.hasAttribute('part')
				? [0, 0, 800, 600]
				: [300, 300, 400, 340];
			return {
				left: left!,
				top: top!,
				right: right!,
				bottom: bottom!,
				x: left!,
				y: top!,
				width: right! - left!,
				height: bottom! - top!,
				toJSON: () => ({}),
			};
		});
		editable.select();
		expect(editable.viewport.querySelector('[data-resize-handle]')).not.toBeNull();
		expect(rotationStarts(editable)).toBe(true);
		editable.viewport.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'x', bubbles: true, cancelable: true }),
		);
		const input = editable.viewport.querySelector<HTMLTextAreaElement>('textarea.inline-text');
		expect(input).not.toBeNull();
		expect(input!.hidden).toBe(false);
		editable.controller.canEdit = false;
		expect(editable.viewport.querySelector('[data-resize-handle]')).toBeNull();
		expect(rotationStarts(editable)).toBe(false);
	} finally {
		editable.dispose();
	}
	const readOnly = await surface(false);
	try {
		readOnly.select();
		expect(readOnly.viewport.querySelector('[data-resize-handle]')).toBeNull();
		expect(rotationStarts(readOnly)).toBe(false);
		readOnly.viewport.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'x', bubbles: true, cancelable: true }),
		);
		const input = readOnly.viewport.querySelector<HTMLTextAreaElement>('textarea.inline-text');
		expect(input === null || input.hidden).toBe(true);
	} finally {
		readOnly.dispose();
	}
});

it('deletes the selection from the toolbar callback and the Delete key, gated by permission', async () => {
	const ui = await surface(true);
	try {
		ui.select();
		const props = diagramToolbarProps(ui.controller, () => {});
		expect(props.canDelete).toBe(true);
		expect(props.canCopy).toBe(true);
		await props.onDelete();
		expect(ui.edits.at(-1)).toEqual([{ type: 'delete-shape', pageId: '1', shapeId: '1' }]);
	} finally {
		ui.dispose();
	}
	const readOnly = await surface(false);
	try {
		readOnly.select();
		const props = diagramToolbarProps(readOnly.controller, () => {});
		expect(props.mode).toBe('view');
		expect([props.canDelete, props.canCut, props.canPaste, props.canDraw]).toEqual([
			false,
			false,
			false,
			false,
		]);
		// Copy only reads the drawing, so it stays available.
		expect(props.canCopy).toBe(true);
		await props.onDelete();
		readOnly.viewport.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Delete', bubbles: true, cancelable: true }),
		);
		await Promise.resolve();
		expect(readOnly.edits).toEqual([]);
	} finally {
		readOnly.dispose();
	}
});

it('reports search state, tools and host dialog availability as flat props', async () => {
	const ui = await surface(true);
	try {
		ui.controller.setSearchQuery('Movable');
		const open = vi.fn();
		const props = diagramToolbarProps(ui.controller, () => {}, 'de', { onOpenSettings: open });
		expect(props.search).toMatchObject({ query: 'Movable', resultCount: 1 });
		expect(props.tools.map((tool) => tool.id)).toEqual(['pointer', ...DIAGRAM_TOOLS]);
		expect(props.labels.Undo).toBe('Rückgängig');
		expect(props.labels.Delete).toBe('Löschen');
		expect(props.canOpenSettings).toBe(true);
		props.onOpenSettings?.();
		expect(open).toHaveBeenCalledOnce();
		expect(props.canOpenShareDialog).toBe(false);
		expect('onOpenShareDialog' in props).toBe(false);
		expect(props.canOpenPageSetup || props.canOpenFile).toBe(false);
	} finally {
		ui.dispose();
	}
});

it('mounts the shapes pane on its own and adds a shape through the shared flow', async () => {
	const ui = await surface(true);
	const host = document.createElement('div');
	document.body.append(host);
	const stop = mountDiagramShapes(host, ui.controller);
	try {
		const pane = host.shadowRoot!.querySelector('.shapes-pane')!;
		expect(pane).not.toBeNull();
		const master = pane.querySelector<HTMLElement>('[data-master]');
		expect(master).not.toBeNull();
		master!.click();
		await vi.waitFor(() => expect(ui.edits.length).toBeGreaterThan(0));
	} finally {
		stop();
		expect(host.shadowRoot!.children).toHaveLength(0);
		ui.dispose();
	}
});
