export * from './gallery-types';
export * from './gallery-module';
export * from './gallery-registry';
export * from './gallery-placements';
export * from './gallery-view';
export * from './gallery-preview-svg';
export * from './gallery-element-patch';
export {
	SHAPE_PRESET_ROWS,
	SHAPE_PRESET_STYLE_ROWS,
	SHAPE_STYLE_COLUMNS,
	SHAPE_THEME_STYLE_ROWS,
	type ShapePresetStyleSpec,
	type ShapeStyleColumn,
} from 'ooxml-core/drawingml';
export { themeShapeStyleXml } from './shape-style-xml';
export * from './picture-adjust-catalog';
export * from './picture-adjust-gallery';
export * from './picture-adjust-tile-svg';
export * from './picture-artistic-gallery';
export * from './picture-color-gallery';
export * from './picture-corrections-gallery';
export { smartArtLayoutSwitchPatch, switchSmartArtLayoutData } from './smartart-gallery-patch';
export { tableLayoutCells } from './table-layout-ops';
export { buildQuickStylesGallery, quickStyleShapeStyle } from './quick-styles-gallery';
export { buildTextWarpGallery, textWarpStyle } from './text-warp-gallery';
export type { TableCursorState, TableLayoutCell } from './table-layout-ops';
