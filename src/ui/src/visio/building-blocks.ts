import type { ViewerController } from './controller';
import { ViewerCanvas } from './viewer-canvas';
import { wireViewerInputs } from './viewer-input';
import { ViewerPointerGestures } from './viewer-pointer-gestures';
import { fitZoom } from './viewer-fit';
import { viewerStyles } from './styles';
import { canvasAndRibbonStyles, visioThemeAliases } from './styles/index';
import { selectedShape, shapeDetails } from './shape-inspector';
import { normalizeVisioLocale, translateRibbonText } from './ribbon-locale';

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

/** The shared drawing engine, without application chrome or file/settings handlers. */
export function mountDiagramCanvas(host: HTMLElement, controller: ViewerController): () => void {
	const doc = host.ownerDocument;
	const root = host.shadowRoot ?? host.attachShadow({ mode: 'open' });
	const style = doc.createElement('style');
	style.textContent = `${visioThemeAliases}${viewerStyles}${canvasAndRibbonStyles}`;
	const viewport = doc.createElement('div');
	viewport.className = 'viewport';
	viewport.tabIndex = 0;
	viewport.setAttribute('part', 'canvas');
	const notes = doc.createElement('ul');
	const inspector = doc.createElement('aside');
	inspector.append(doc.createElement('div'));
	const canvas = new ViewerCanvas(viewport, notes, doc.createElement('aside'), inspector);
	const fit = (mode: 'page' | 'width') => fitDiagramCanvas(host, controller, mode);
	root.append(style, viewport);
	const slider = Object.assign(doc.createElement('span'), { value: 100 });
	const offInputs = wireViewerInputs({ viewport, zoomSlider: slider }, controller, fit);
	const pointer = new ViewerPointerGestures(viewport, controller, {
		active: () => true,
		announce: () => {},
	});
	const offPointer = pointer.wire();
	const render = () => {
		pointer.render(controller.state);
		canvas.render(controller.state);
	};
	const stop = controller.subscribe(render);
	render();
	return () => {
		stop();
		offInputs();
		offPointer();
		canvas.dispose();
		root.replaceChildren();
	};
}

export function diagramToolbarProps(controller: ViewerController, fit: () => void, locale = 'en') {
	const state = controller.state;
	const canEdit = state.canEdit && state.edit.sourceAvailable && !state.loading && !state.edit.busy;
	return {
		labels: Object.fromEntries(
			['Undo', 'Redo', 'Fit to Window', 'Page'].map((label) => [
				label,
				translateRibbonText(normalizeVisioLocale(locale), label),
			]),
		),
		canEdit,
		pageIndex: state.pageIndex,
		pages: state.document?.pages.map((page) => ({ id: page.id, name: page.name })) ?? [],
		zoom: state.zoom,
		selection: state.selectedShapes,
		canUndo: canEdit && state.edit.canUndo,
		canRedo: canEdit && state.edit.canRedo,
		canDuplicate: canEdit && state.selectedShapes.length > 0,
		canGroup: canEdit && state.selectedShapes.length > 1,
		canUngroup: canEdit && state.selectedShapes.length > 0,
		onPageChange: (index: number) => controller.setPage(index),
		onZoomChange: (zoom: number) => controller.setZoom(zoom),
		onFit: fit,
		onUndo: () => controller.undo(),
		onRedo: () => controller.redo(),
		onDuplicate: () => controller.duplicateSelection(),
		onGroup: () => controller.groupSelection('group'),
		onUngroup: () => controller.groupSelection('ungroup'),
		onTextChange: (pageId: string, shapeId: string, text: string) =>
			controller.replacePlainText(pageId, shapeId, text),
		onSearchChange: (query: string) => controller.setSearchQuery(query),
		onNextSearchResult: () => controller.nextSearchResult(),
		onPreviousSearchResult: () => controller.previousSearchResult(),
	};
}
export type DiagramToolbarProps = ReturnType<typeof diagramToolbarProps>;
