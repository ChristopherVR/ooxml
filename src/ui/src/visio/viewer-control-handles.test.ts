import { afterEach, expect, it, vi } from 'vitest';
import { demoDocument } from 'ooxml-core/visio/ui';
import type { VisioControl } from 'ooxml-core/visio';
import { ViewerController } from './controller';
import { ViewerControlHandles } from './viewer-control-handles';

const SVG = 'http://www.w3.org/2000/svg';
const screen = { a: 10, b: 0, d: 10, inverse: () => ({}) };
Object.defineProperty(SVGElement.prototype, 'getScreenCTM', {
	value: () => screen,
	configurable: true,
	writable: true,
});

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

const free: VisioControl = {
	row: 'Row_1',
	x: 1,
	y: 2,
	prompt: 'Slant',
	canGlue: false,
	drag: { x: true, y: false },
	inherited: true,
};
const fixed: VisioControl = {
	row: 'Row_2',
	x: 3,
	y: 2,
	canGlue: false,
	drag: { x: false, y: false },
	locked: 'Its position is fixed, or follows another handle.',
	inherited: true,
};

async function setup() {
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
	const model = structuredClone(demoDocument);
	const shape = model.pages[0]!.shapes[0]!;
	shape.id = '7';
	shape.controls = [free, fixed];
	const edit = vi.fn(async (bytes: Uint8Array, _edits: readonly unknown[]) => ({
		bytes,
		changedParts: [],
		diagnostics: [],
		document: model,
	}));
	const controller = new ViewerController(async () => model, vi.fn(), edit as never);
	await controller.load(new Uint8Array([1]));
	controller.selectShape({ id: shape.id, name: shape.name, pageId: model.pages[0]!.id });
	const viewport = document.createElement('div');
	const svg = document.createElementNS(SVG, 'svg');
	svg.classList.add('paper');
	const group = document.createElementNS(SVG, 'g');
	group.dataset.shapeId = shape.id;
	group.dataset.pageId = model.pages[0]!.id;
	group.dataset.selected = 'true';
	svg.append(group);
	viewport.append(svg);
	document.body.append(viewport);
	const announce = vi.fn();
	const tool = new ViewerControlHandles(viewport, controller, { active: () => true, announce });
	const dispose = tool.wire();
	tool.render(controller.state);
	const fire = (type: string, target: Element, x: number, y: number) => {
		const event = new Event(type, { bubbles: true, cancelable: true });
		Object.defineProperties(event, {
			pointerId: { value: 7 },
			button: { value: 0 },
			clientX: { value: x },
			clientY: { value: y },
		});
		target.dispatchEvent(event);
	};
	return { controller, svg, viewport, announce, edit, fire, dispose, shape, model };
}

it('draws a yellow handle for every control and marks the ones that cannot move', async () => {
	const ui = await setup();
	const handles = ui.svg.querySelectorAll<SVGRectElement>('[data-control-handle]');
	expect([...handles].map((handle) => handle.dataset.controlHandle)).toEqual(['Row_1', 'Row_2']);
	expect(handles[0]!.hasAttribute('data-control-locked')).toBe(false);
	expect(handles[1]!.hasAttribute('data-control-locked')).toBe(true);
	expect(handles[0]!.getAttribute('aria-label')).toContain('Slant');
	expect(handles[1]!.querySelector('title')!.textContent).toMatch(/fixed/);
	ui.dispose();
});

it('says why a locked handle cannot be dragged and edits nothing', async () => {
	const ui = await setup();
	const locked = ui.svg.querySelector('[data-control-handle="Row_2"]')!;
	ui.fire('pointerdown', locked, 30, 20);
	ui.fire('pointerup', locked, 60, 20);
	expect(ui.announce).toHaveBeenCalledWith(expect.stringContaining('cannot be moved'));
	expect(ui.edit).not.toHaveBeenCalled();
	ui.dispose();
});

it('moves a free handle along its axis and commits one set-control edit on release', async () => {
	const ui = await setup();
	const handle = ui.svg.querySelector<SVGRectElement>('[data-control-handle="Row_1"]')!;
	ui.fire('pointerdown', handle, 10, 20);
	ui.fire('pointermove', ui.svg, 50, 90);
	// The handle follows the pointer on x only.
	expect(Number(handle.getAttribute('x')) + Number(handle.getAttribute('width')) / 2).toBeCloseTo(
		5,
	);
	expect(Number(handle.getAttribute('y')) + Number(handle.getAttribute('height')) / 2).toBeCloseTo(
		2,
	);
	expect(ui.edit).not.toHaveBeenCalled();
	ui.fire('pointerup', ui.svg, 50, 90);
	await vi.waitFor(() => expect(ui.edit).toHaveBeenCalledTimes(1));
	expect(ui.edit.mock.calls[0]![1]).toEqual([
		{
			type: 'set-control',
			pageId: ui.model.pages[0]!.id,
			shapeId: ui.shape.id,
			row: 'Row_1',
			x: 5,
		},
	]);
	ui.dispose();
});
