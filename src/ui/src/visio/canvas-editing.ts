import type { ViewerController, ViewerState } from './controller';
import { ViewerInlineText } from './viewer-inline-text';
import { ViewerLineEndpoints } from './viewer-line-endpoints';
import { ViewerPointerGestures } from './viewer-pointer-gestures';
import { ViewerResizeHandles } from './viewer-resize-handles';
import { ViewerRotationHandle } from './viewer-rotation-handle';
import { ShapeDrawTool } from './viewer-draw-tool';
import { ViewerTextTool } from './viewer-text-tool';
import { ViewerConnectorTool } from './viewer-connector-tool';
import type { DiagramTool } from './diagram-tool';

export interface CanvasEditingOptions {
	/** Whether the pointer tool is the active one: handles and gestures stand down otherwise. */
	pointerActive(): boolean;
	announce(message: string): void;
	snap?: ConstructorParameters<typeof ViewerPointerGestures>[2]['snap'];
	clearSnap?: ConstructorParameters<typeof ViewerPointerGestures>[2]['clearSnap'];
	/**
	 * The active tool, present only where the tools live in this wiring (the standalone canvas).
	 * The full viewer owns its tools in its commands and leaves this out.
	 */
	tool?: () => DiagramTool;
}

/**
 * The editing gestures every canvas shares: moving, inline text, line end points, rotation and
 * resize handles, plus ordinary shape drawing when a draw tool is supplied. The full viewer and
 * the standalone building-block canvas both build their surface from this, so the two behave alike
 * and both stay behind the controller's edit permission.
 */
export function createCanvasEditing(
	viewport: HTMLElement,
	controller: ViewerController,
	options: CanvasEditingOptions,
) {
	const gesture = { active: options.pointerActive, announce: options.announce };
	const pointer = new ViewerPointerGestures(viewport, controller, {
		...gesture,
		...(options.snap ? { snap: options.snap } : {}),
		...(options.clearSnap ? { clearSnap: options.clearSnap } : {}),
	});
	const lineEndpoints = new ViewerLineEndpoints(viewport, controller, gesture);
	const rotation = new ViewerRotationHandle(viewport, controller, gesture);
	const resize = new ViewerResizeHandles(viewport, controller, gesture);
	const edit = new ViewerInlineText(viewport, controller, options.announce);
	const current = options.tool;
	const draw = current
		? new ShapeDrawTool(viewport, controller, {
				tool: () => {
					const tool = current();
					return tool === 'pointer' || tool === 'text' || tool === 'connector' ? undefined : tool;
				},
				announce: options.announce,
			})
		: undefined;
	const text = current
		? new ViewerTextTool(viewport, controller, {
				active: () => current() === 'text',
				announce: options.announce,
				revealEdit: () => {},
			})
		: undefined;
	const connector = current
		? new ViewerConnectorTool(viewport, controller, {
				active: () => current() === 'connector',
				announce: options.announce,
			})
		: undefined;
	return {
		pointer,
		lineEndpoints,
		rotation,
		resize,
		edit,
		draw,
		text,
		connector,
		/** Stop any gesture in progress, as when the tool changes or Escape is pressed. */
		cancelTools(): void {
			draw?.cancel();
			text?.cancel();
			connector?.cancel();
		},
		/** Redraw the selection handles after a state or tool change. */
		renderHandles(state: ViewerState): void {
			lineEndpoints.render(state);
			rotation.render(state);
			resize.render(state);
		},
		render(state: ViewerState): void {
			pointer.render(state);
			draw?.render(state);
			text?.render(state);
			connector?.render(state);
			edit.render(state);
			this.renderHandles(state);
		},
		wire(): () => void {
			const stops = [
				pointer.wire(),
				lineEndpoints.wire(),
				rotation.wire(),
				resize.wire(),
				edit.wire(),
				...(draw ? [draw.wire()] : []),
				...(text ? [text.wire()] : []),
				...(connector ? [connector.wire()] : []),
			];
			return () => {
				for (const stop of stops) stop();
			};
		},
	};
}
export type CanvasEditing = ReturnType<typeof createCanvasEditing>;
