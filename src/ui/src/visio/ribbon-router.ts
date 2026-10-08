import type { ViewerController } from './controller';
import type { VisioRibbonAction, CanvasTool, VisioFormattingAction } from './ribbon-action';

/** The controllers a ribbon action can reach. The element supplies each one. */
export interface RibbonTargets {
	controller: ViewerController;
	history(key: 'undo' | 'redo'): void;
	deleteSelection(): void;
	duplicateSelection(): void;
	clipboard(action: Extract<VisioRibbonAction, { type: 'clipboard' }>): void;
	rotateSelection(direction: 'left' | 'right'): void;
	flipSelection(axis: 'horizontal' | 'vertical'): void;
	formatSelection(action: VisioFormattingAction): void;
	arrangeSelection(action: Extract<VisioRibbonAction, { type: 'arrange' }>): void;
	setTool(tool: CanvasTool): void;
	cancelDrawing(): void;
	insertPage(): void;
	showPaintProperties(): void;
	toggleGrid(): void;
	toggleRuler(): void;
	togglePanZoom(): void;
	toggleSizePosition(): void;
	toggleFullscreen(): void;
	togglePane(pane: 'shapes' | 'inspector'): void;
	reveal(panel: 'edit' | 'notes' | 'selection' | 'layers', focusText: boolean): void;
	fit(mode: 'page' | 'width'): void;
	focusSearch(): void;
}

/** Routes a ribbon, status-bar or shortcut action to the controller that owns it. */
export function routeRibbonAction(targets: RibbonTargets, action: VisioRibbonAction): void {
	const { controller } = targets;
	const page = controller.state.document?.pages[controller.state.pageIndex];
	switch (action.type) {
		case 'arrange':
			return targets.arrangeSelection(action);
		case 'selection':
			return action.mode === 'all' ? controller.selectAll() : controller.clearSelection();
		case 'text-toggle':
		case 'font-color':
		case 'text-bullets':
		case 'text-indent':
		case 'font-family':
		case 'font-size':
		case 'font-step':
		case 'text-align':
		case 'shape-format':
		case 'shape-order':
			return targets.formatSelection(action);
		case 'history':
			return targets.history(action.key);
		case 'delete':
			return targets.deleteSelection();
		case 'duplicate':
			return targets.duplicateSelection();
		case 'clipboard':
			return targets.clipboard(action);
		case 'rotate':
			return targets.rotateSelection(action.direction);
		case 'flip':
			return targets.flipSelection(action.axis);
		case 'tool':
			return targets.setTool(action.tool);
		case 'cancel-drawing':
			return targets.cancelDrawing();
		case 'page-insert':
			return targets.insertPage();
		case 'paint-properties':
			return targets.showPaintProperties();
		case 'grid':
			return targets.toggleGrid();
		case 'ruler':
			return targets.toggleRuler();
		case 'panZoom':
			return targets.togglePanZoom();
		case 'sizePosition':
			return targets.toggleSizePosition();
		case 'fullscreen':
			return targets.toggleFullscreen();
		case 'pane':
			return targets.togglePane(action.pane);
		case 'reveal':
			return targets.reveal(action.panel, action.focusText ?? false);
		case 'search':
			return targets.focusSearch();
		case 'page':
			return controller.setPage(controller.state.pageIndex + action.step);
		case 'zoomTo':
			if (page) controller.setZoom(action.percent / 100);
			return;
		case 'zoom':
			if (!page) return;
			if (action.mode === 'actual') return controller.setZoom(1);
			return targets.fit(action.mode === 'fit' ? 'page' : 'width');
	}
}
