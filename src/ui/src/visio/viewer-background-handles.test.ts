import { afterEach, expect, it, vi } from 'vitest';
import { demoDocument } from 'ooxml-core/visio/ui';
import { ViewerController } from './controller';
import { ViewerRotationHandle } from './viewer-rotation-handle';
import { ViewerLineEndpoints } from './viewer-line-endpoints';
import { handleGestureIsCurrent } from './viewer-handle-events';

afterEach(() => document.body.replaceChildren());

it.each(['rotation', 'endpoint'] as const)(
	'does not draw or start %s handles for a background shape sharing a foreground ID',
	async (kind) => {
		const model = structuredClone(demoDocument);
		const page = model.pages[0]!,
			background = model.pages[1]!;
		const shape = page.shapes[0]!;
		if (kind === 'endpoint')
			Object.assign(shape, {
				kind: 'connector',
				width: 2,
				height: 0,
				geometry: [{ path: 'M 0 0 L 2 0', fill: false, stroke: true }],
			});
		page.backgroundPageId = background.id;
		background.isBackground = true;
		background.shapes = [structuredClone(shape)];
		const edit = vi.fn();
		const controller = new ViewerController(async () => model, vi.fn(), edit);
		await controller.load(new Uint8Array([1]));
		controller.selectShape({ id: shape.id, name: 'Background', pageId: background.id });
		const viewport = document.createElement('div');
		const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
		svg.classList.add('paper');
		viewport.append(svg);
		document.body.append(viewport);
		const options = { active: () => true, announce: vi.fn() };
		const handles =
			kind === 'rotation'
				? new ViewerRotationHandle(viewport, controller, options)
				: new ViewerLineEndpoints(viewport, controller, options);
		const dispose = handles.wire();
		handles.render(controller.state);
		expect(viewport.querySelectorAll('[data-rotation-handle], [data-line-endpoint]')).toHaveLength(
			0,
		);
		expect(
			handleGestureIsCurrent(
				{ pointer: 7, svg, page, document: model, shapeId: shape.id },
				controller.state,
				true,
			),
		).toBe(false);
		// A stale or synthetic handle must not turn the selected background identity into a local edit.
		const stale = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
		stale.dataset.rotationHandle = shape.id;
		stale.dataset.lineEndpoint = 'begin';
		stale.dataset.lineShapeId = shape.id;
		svg.append(stale);
		for (const type of ['pointerdown', 'pointermove', 'pointerup']) {
			const event = new Event(type, { bubbles: true, cancelable: true });
			Object.defineProperties(event, {
				pointerId: { value: 7 },
				button: { value: 0 },
				clientX: { value: 10 },
				clientY: { value: 10 },
			});
			stale.dispatchEvent(event);
		}
		await Promise.resolve();
		expect(viewport.querySelectorAll('.rotation-preview, .endpoint-preview')).toHaveLength(0);
		expect(edit).not.toHaveBeenCalled();
		expect(controller.state.document!.pages[0]!.shapes[0]).toEqual(shape);
		dispose();
		controller.destroy();
	},
);
