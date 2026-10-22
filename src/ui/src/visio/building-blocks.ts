import { installOfficeUiThemeFor, prepareOfficeUiBlock } from '../theme';
import type { ViewerController } from './controller';
import { ViewerCanvas } from './viewer-canvas';
import { wireViewerInputs } from './viewer-input';
import { ViewerClipboard } from './viewer-clipboard';
import { fitZoom } from './viewer-fit';
import { viewerStyles } from './styles';
import { canvasAndRibbonStyles, visioThemeAliases } from './styles/index';
import { selectedShape, shapeDetails } from './shape-inspector';
import { createCanvasEditing } from './canvas-editing';
import { registerCanvasServices } from './canvas-services';
import { wireCanvasKeys } from './canvas-keys';
import { isEditCancellation, editErrorMessage } from 'ooxml-core/visio/ui';

export { diagramToolbarProps, diagramMode } from './building-blocks-toolbar';
export type {
	DiagramToolbarProps,
	DiagramToolbarLabels,
	DiagramToolbarHostCallbacks,
	DiagramMode,
	DiagramToolSummary,
} from './building-blocks-toolbar';
export { mountDiagramShapes } from './building-blocks-shapes';
export {
	DIAGRAM_DRAW_TOOLS,
	DIAGRAM_OBJECT_TOOLS,
	DIAGRAM_TOOLS,
	type DiagramTool,
	type DiagramDrawTool,
	type DiagramObjectTool,
} from './diagram-tool';

export function fitDiagramCanvas(
	host: HTMLElement,
	controller: ViewerController,
	mode: 'page' | 'width' = 'page',
): void {
	const viewport = host.shadowRoot?.querySelector<HTMLElement>('[part="canvas"]');
	const state = controller.state;
	const page = state.document?.pages[state.pageIndex];
	if (viewport && page) controller.setZoom(fitZoom(viewport, page, mode));
}

/** The existing non-mutating selection inspector, independently mounted. */
export function mountDiagramInspector(host: HTMLElement, controller: ViewerController): () => void {
	prepareOfficeUiBlock(host, 'inspector', 'visio');
	const render = () => {
		const state = controller.state;
		const shape = selectedShape(state.document, state.selectedShape, state.pageIndex);
		host.replaceChildren(...(shape ? [shapeDetails(shape)] : []));
	};
	const stop = controller.subscribe(render);
	render();
	return () => {
		stop();
		host.replaceChildren();
	};
}

/**
 * The shared drawing engine, without application chrome or file/settings handlers. Moving,
 * inline text, resize, rotate, line end points and the draw tools are the full viewer's own
 * gestures, and all of them stand down unless the controller may edit.
 */
export function mountDiagramCanvas(host: HTMLElement, controller: ViewerController): () => void {
	const doc = host.ownerDocument;
	installOfficeUiThemeFor(host);
	const root = host.shadowRoot ?? host.attachShadow({ mode: 'open' });
	const style = doc.createElement('style');
	style.textContent = `${visioThemeAliases}${viewerStyles}${canvasAndRibbonStyles}`;
	const viewport = doc.createElement('div');
	viewport.className = 'viewport';
	viewport.tabIndex = 0;
	viewport.setAttribute('part', 'canvas');
	// Edit rejections and gesture feedback reach assistive technology; the host owns visible UI.
	const live = doc.createElement('div');
	live.setAttribute('role', 'status');
	live.setAttribute('aria-live', 'polite');
	live.style.cssText =
		'position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap';
	const announce = (message: string) => {
		live.textContent = message;
	};
	const notes = doc.createElement('ul');
	const inspector = doc.createElement('aside');
	inspector.append(doc.createElement('div'));
	const canvas = new ViewerCanvas(viewport, notes, doc.createElement('aside'), inspector);
	const fit = (mode: 'page' | 'width') => fitDiagramCanvas(host, controller, mode);
	root.append(style, viewport, live);
	const slider = Object.assign(doc.createElement('span'), { value: 100 });
	const pointerActive = () => controller.state.tool === 'pointer';
	const editing = createCanvasEditing(viewport, controller, {
		pointerActive,
		announce,
		tool: () => controller.state.tool,
	});
	const clipboard = new ViewerClipboard(root, controller, announce, () => {});
	const report = (error: unknown) => {
		if (!isEditCancellation(error)) announce(editErrorMessage(error));
	};
	const offServices = registerCanvasServices(controller, {
		viewport,
		announce,
		clipboard: (operation) => clipboard.run(operation),
	});
	const offClipboard = clipboard.wire((operation, event) => {
		void clipboard.run(operation, event).catch(report);
	});
	const offInputs = wireViewerInputs(
		{ viewport, zoomSlider: slider },
		controller,
		fit,
		announce,
		(character) => pointerActive() && editing.edit.start(character),
	);
	const offKeys = wireCanvasKeys(viewport, controller, editing, announce);
	const offEditing = editing.wire();
	let lastTool = controller.state.tool;
	const render = () => {
		// A tool change abandons the gesture of the previous tool, as in the full viewer.
		if (controller.state.tool !== lastTool) editing.cancelTools();
		lastTool = controller.state.tool;
		canvas.render(controller.state);
		editing.render(controller.state);
	};
	const stop = controller.subscribe(render);
	render();
	return () => {
		stop();
		offInputs();
		offKeys();
		offEditing();
		offClipboard();
		offServices();
		canvas.dispose();
		root.replaceChildren();
	};
}
