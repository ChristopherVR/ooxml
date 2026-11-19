export { BuildingBlocksCanvasComponent } from './canvas.component';
export { BuildingBlocksSlidesPaneComponent } from './slides-pane.component';
export { BuildingBlocksToolbarComponent } from './toolbar.component';
export { ExternalContentBarComponent } from '../viewer/external-content-bar.component';
export {
	createViewerBuildingBlocks,
	VIEWER_BUILDING_BLOCKS_PROVIDERS,
	ViewerBuildingBlocks,
} from './viewer-building-blocks';
export type { ViewerBuildingBlocksOptions, ViewerBuildingBlocksState } from './types';
export type {
	BuildingBlockRequest,
	BuildingBlockResult,
	ViewerBuildingBlockActions,
	ViewerExternalContentBar,
} from 'ooxml-ui/pptx';
// The reporter every viewer and building-blocks entry takes as `onElementRenderError`.
export type { ElementRenderErrorContext, ElementRenderErrorHandler } from 'ooxml-ui/pptx';
