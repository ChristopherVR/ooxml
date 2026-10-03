import type { ViewerController } from './controller.js';
import type { VisioRibbonAction } from './ribbon-action.js';

/** The controllers a ribbon action can reach. The element supplies each one. */
export interface RibbonTargets {
	controller: ViewerController;
	history(key: 'undo' | 'redo'): void;
	deleteSelection(): void;
	setTool(tool: 'pointer' | 'rectangle'): void;
	toggleGrid(): void;
	toggleRuler(): void;
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
		case 'history':
			return targets.history(action.key);
		case 'delete':
			return targets.deleteSelection();
		case 'tool':
			return targets.setTool(action.tool);
		case 'grid':
			return targets.toggleGrid();
		case 'ruler':
			return targets.toggleRuler();
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
