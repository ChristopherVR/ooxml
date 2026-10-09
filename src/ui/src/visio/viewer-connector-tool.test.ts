import { afterEach, expect, it, vi } from 'vitest';
import { pointerViewer } from './__fixtures__/pointer-viewer';
import { ViewerConnectorTool, connectorGlueTarget } from './viewer-connector-tool';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

async function connectorViewer() {
	const ui = await pointerViewer();
	ui.inactive();
	let active = true;
	const announce = vi.fn();
	const tool = new ViewerConnectorTool(ui.viewport, ui.controller, {
		active: () => active,
		announce,
	});
	const dispose = tool.wire();
	return {
		...ui,
		tool,
		announce,
		deactivate() {
			active = false;
			tool.render(ui.controller.state);
		},
		close() {
			dispose();
			ui.dispose();
		},
	};
}

it('highlights glue targets and glues both ends of a drawn connector', async () => {
	const ui = await connectorViewer();
	try {
		ui.pointer('pointermove', ui.group('1'), 40, 40);
		expect(ui.group('1').dataset.connectTarget).toBe('true');
		ui.pointer('pointerdown', ui.group('1'), 40, 40);
		ui.pointer('pointermove', ui.group('2'), 70, 40);
		expect(ui.tool.target).toBe('2');
		expect(ui.group('1').dataset.connectTarget).toBeUndefined();
		expect(ui.group('2').dataset.connectTarget).toBe('true');
		ui.pointer('pointerup', ui.group('2'), 70, 40);
		await ui.done();
		expect(ui.edits).toHaveLength(1);
		expect(ui.edits[0]![0]).toMatchObject({
			type: 'create-line',
			connect: { begin: '1', end: '2' },
		});
		const page = ui.controller.state.document!.pages[0]!;
		expect(page.connectors.map((connection) => connection.toShapeId)).toEqual(['1', '2']);
		expect(ui.announce).toHaveBeenCalledWith(
			expect.stringMatching(/^Connector \d+ added and glued\.$/),
		);
		expect(ui.viewport.querySelector('[data-connect-target]')).toBeNull();
		// Moving a glued shape is now offered and the connector follows it.
		const connector = page.connectors[0]!.fromShapeId;
		const before = page.shapes.find((shape) => shape.id === connector)!.transform;
		await ui.controller.applyEdits([
			{ type: 'move-shape', pageId: page.id, shapeId: '2', x: 7, y: 6 },
		]);
		const moved = ui.controller.state.document!.pages[0]!;
		expect(moved.connectors).toHaveLength(2);
		expect(moved.shapes.find((shape) => shape.id === connector)!.transform).not.toEqual(before);
		await ui.controller.undo();
		await ui.controller.undo();
		expect(ui.controller.state.document!.pages[0]!.connectors).toEqual([]);
	} finally {
		ui.close();
	}
});

it('leaves an end on empty canvas unglued and clears feedback when inactive', async () => {
	const ui = await connectorViewer();
	try {
		ui.pointer('pointerdown', ui.group('1'), 40, 40);
		ui.pointer('pointermove', ui.svg, 40, 90);
		expect(ui.tool.target).toBeUndefined();
		ui.pointer('pointerup', ui.svg, 40, 90);
		await ui.done();
		expect(ui.edits[0]![0]).toMatchObject({ connect: { begin: '1' } });
		expect(ui.controller.state.document!.pages[0]!.connectors).toHaveLength(1);
		ui.pointer('pointermove', ui.group('2'), 70, 40);
		expect(ui.group('2').dataset.connectTarget).toBe('true');
		ui.deactivate();
		expect(ui.viewport.querySelector('[data-connect-target]')).toBeNull();
	} finally {
		ui.close();
	}
});

it('only offers visible local top-level 2D shapes as glue targets', async () => {
	const ui = await connectorViewer();
	try {
		const page = ui.controller.state.document!.pages[0]!;
		expect(connectorGlueTarget(page, ui.group('1'))).toBe('1');
		expect(connectorGlueTarget(page, ui.svg)).toBeUndefined();
		expect(
			connectorGlueTarget(
				{ ...page, shapes: [{ ...page.shapes[0]!, kind: 'connector' }] },
				ui.group('1'),
			),
		).toBeUndefined();
	} finally {
		ui.close();
	}
});
