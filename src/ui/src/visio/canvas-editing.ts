import type { ViewerController, ViewerState } from './controller';
import { ViewerInlineText } from './viewer-inline-text';
import { ViewerLineEndpoints } from './viewer-line-endpoints';
import { ViewerPointerGestures } from './viewer-pointer-gestures';
import { ViewerResizeHandles } from './viewer-resize-handles';
import { ViewerRotationHandle } from './viewer-rotation-handle';
import { ShapeDrawTool } from './viewer-draw-tool';
import type { DiagramDrawTool } from './diagram-tool';

export interface CanvasEditingOptions {
	/** Whether the pointer tool is the active one: handles and gestures stand down otherwise. */
	pointerActive(): boolean;
	announce(message: string): void;
	snap?: ConstructorParameters<typeof ViewerPointerGestures>[2]['snap'];
	clearSnap?: ConstructorParameters<typeof ViewerPointerGestures>[2]['clearSnap'];
	/** Present only where the draw tool lives outside this wiring (the full viewer's commands). */
	drawTool?: () => DiagramDrawTool | undefined;
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
	const draw = options.drawTool
		? new ShapeDrawTool(viewport, controller, {
				tool: options.drawTool,
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
		/** Redraw the selection handles after a state or tool change. */
		renderHandles(state: ViewerState): void {
			lineEndpoints.render(state);
			rotation.render(state);
			resize.render(state);
		},
		render(state: ViewerState): void {
			pointer.render(state);
			draw?.render(state);
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
			];
			return () => {
				for (const stop of stops) stop();
			};
		},
	};
}
export type CanvasEditing = ReturnType<typeof createCanvasEditing>;
