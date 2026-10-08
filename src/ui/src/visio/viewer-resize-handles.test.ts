import { afterEach, expect, it, vi } from 'vitest';
import { parseVsdx } from 'ooxml-core/visio';
import { ViewerResizeHandles } from './viewer-resize-handles';
import { pointerViewer } from './__fixtures__/pointer-viewer';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});
async function setup(protectedSecond = false, source = true) {
	const ui = await pointerViewer(protectedSecond, source, true);
	Object.defineProperty(ui.svg, 'getScreenCTM', {
		value: () => ({ a: 10, b: 0, inverse: () => ({}) }),
		configurable: true,
	});
	let active = true;
	const tool = new ViewerResizeHandles(ui.viewport, ui.controller, {
		active: () => active,
		announce: (message) => ui.feedback.push(message),
	});
	const dispose = tool.wire();
	const unsubscribe = ui.controller.subscribe((state) => {
		for (const group of ui.svg.querySelectorAll<SVGGElement>('[data-shape-id]'))
			group.dataset.selected = String(
				state.selectedShapes.some((shape) => shape.id === group.dataset.shapeId),
			);
		tool.render(state);
	});
	return {
		...ui,
		tool,
		handle: (id = 'ne') => ui.svg.querySelector<SVGCircleElement>(`[data-resize-handle="${id}"]`)!,
		inactive() {
			active = false;
			tool.render(ui.controller.state);
		},
		dispose() {
			unsubscribe();
			dispose();
			ui.dispose();
		},
	};
}
it('renders eight single-selection handles, previews only a frame and commits anchored source once on release', async () => {
	const ui = await setup();
	ui.select();
	const selection = ui.controller.state.selectedShapes;
	expect(ui.svg.querySelectorAll('[data-resize-handle]')).toHaveLength(8);
	const original = ui.group().outerHTML;
	const rotation = document.createElement('div');
	rotation.className = 'rotation-overlay';
	ui.viewport.append(rotation);
	ui.pointer('pointerdown', ui.handle(), 55, 35);
	ui.pointer('pointermove', ui.svg, 65, 25);
	expect(rotation.style.visibility).toBe('hidden');
	expect(ui.svg.querySelector('[data-resize-frame-preview]')!.getAttribute('width')).toBe('4');
	expect(ui.svg.querySelector('[data-resize-frame-preview]')!.getAttribute('height')).toBe('2');
	expect(ui.group().outerHTML).toBe(original);
	expect(ui.edits).toHaveLength(0);
	expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
	ui.pointer('pointerup', ui.svg, 65, 25);
	await ui.done();
	expect(rotation.style.visibility).toBe('');
	expect(ui.edits).toEqual([
		[
			{
				type: 'resize-shape',
				pageId: '1',
				shapeId: '1',
				width: 4,
				height: 2,
				anchor: { x: 0, y: 0 },
			},
		],
	]);
	const resized = (await parseVsdx(ui.controller.exportVsdx().bytes)).pages[0]!.shapes[0]!;
	expect(resized).toMatchObject({ width: 4, height: 2, rotation: { pinX: 4.5, pinY: 7.5 } });
	expect(resized.transform.slice(4)).toEqual([2.5, 6.5]);
	expect(ui.controller.state.selectedShapes).toBe(selection);
	expect(ui.feedback).toContain('Shape resized.');
	await ui.controller.undo();
	expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
	await ui.controller.redo();
	expect(ui.controller.state.selectedShapes).toBe(selection);
	ui.dispose();
});
it.each(['Escape', 'pointercancel', 'lostpointercapture', 'selection', 'source', 'tool'] as const)(
	'cancels %s with saved content/source intact',
	async (kind) => {
		const ui = await setup();
		ui.select();
		const original = ui.group().outerHTML;
		ui.pointer('pointerdown', ui.handle(), 55, 35);
		ui.pointer('pointermove', ui.svg, 65, 25);
		if (kind === 'Escape')
			ui.viewport.dispatchEvent(
				new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
			);
		else if (kind === 'selection') ui.controller.clearSelection();
		else if (kind === 'source') ui.controller.setDocument(null);
		else if (kind === 'tool') ui.inactive();
		else ui.pointer(kind, ui.svg);
		ui.pointer('pointerup', ui.svg, 65, 25);
		await Promise.resolve();
		expect(ui.svg.querySelector('[data-resize-frame-preview]')).toBeNull();
		expect(ui.edits).toHaveLength(0);
		expect(ui.group().getAttribute('transform')).toBe(
			new DOMParser()
				.parseFromString(original, 'image/svg+xml')
				.documentElement.getAttribute('transform'),
		);
		if (kind !== 'source') expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
		ui.dispose();
	},
);
it('declines multi-selection, model-only and background identities, including forged handles', async () => {
	const ui = await setup();
	ui.select(['1', '2']);
	expect(ui.svg.querySelectorAll('[data-resize-handle]')).toHaveLength(0);
	ui.select();
	const stale = ui.handle().cloneNode(true) as Element;
	ui.svg.append(stale);
	ui.pointer('pointerdown', stale, 55, 35);
	ui.pointer('pointermove', ui.svg, 65, 25);
	ui.pointer('pointerup', ui.svg, 65, 25);
	expect(ui.edits).toHaveLength(0);
	const model = ui.controller.state.document!;
	model.pages.push({ ...structuredClone(model.pages[0]!), id: 'bg', isBackground: true });
	model.pages[0]!.backgroundPageId = 'bg';
	ui.controller.selectShape({ id: '1', name: 'Background', pageId: 'bg' });
	expect(ui.svg.querySelector('[data-resize-overlay]')).toBeNull();
	ui.dispose();
	const readOnly = await setup(false, false);
	readOnly.select();
	expect(readOnly.svg.querySelector('[data-resize-overlay]')).toBeNull();
	readOnly.dispose();
});
it('a handle click makes no edit and source protection rejects resizing without accepting history', async () => {
	const ui = await setup(true);
	ui.select(['2']);
	ui.pointer('pointerdown', ui.handle(), 55, 35);
	ui.pointer('pointerup', ui.svg, 55, 35);
	expect(ui.edits).toHaveLength(0);
	ui.pointer('pointerdown', ui.handle(), 55, 35);
	ui.pointer('pointermove', ui.svg, 65, 25);
	ui.pointer('pointerup', ui.svg, 65, 25);
	await ui.done();
	expect(ui.edits).toHaveLength(1);
	expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
	expect(ui.controller.state.edit.canUndo).toBe(false);
	expect(ui.feedback.join(' ')).toMatch(/LockMoveX|locked/i);
	ui.dispose();
});
