/**
 * Geometry module barrel export.
 *
 * Re-exports all public APIs for shape geometry resolution, connector path
 * generation, element transforms, preset shape clip-paths / definitions,
 * and the OOXML DrawingML guide formula evaluator.
 *
 * @module geometry
 */

export {
	getShapeType,
	getShapeClipPath,
	getAdjustmentAwareShapeClipPath,
	getCloudPathForRendering,
	getRoundRectRadiusPx,
	getImageMaskStyle,
	getShapeClipPathFromPreset,
} from './shape-geometry';
export type { ImageMaskStyle } from './shape-geometry';

export {
	PRESET_SHAPE_GEOMETRY_TABLE,
	type PresetShapeGeometryDefinition,
	type PresetPath,
	type PresetPathCommand,
} from '../../../geometry/preset-shape-definitions-table.js';

export {
	ST_SHAPE_TYPE_VALUES,
	PRESET_GEOMETRY_ALIASES,
	isStShapeType,
	normalizeStShapeType,
} from '../../../geometry/preset-geometry-names.js';

export {
	evaluatePresetShape,
	lookupPresetShape,
	type PresetShapeEvaluationResult,
	type PresetSubpathResult,
} from './preset-shape-evaluator';

export { filterValidShapeAdjustmentEntries } from './preset-adjustment-validation';

export {
	getPresetConnectionSites,
	lookupPresetConnectionSites,
	type EvaluatedPresetConnectionSite,
} from '../../../geometry/preset-connection-sites-table.js';
export type {
	PresetConnectionSiteDefinition,
	PresetConnectionSiteToken,
} from '../../../geometry/preset-connection-sites-types.js';

export {
	getPresetTextRect,
	lookupPresetTextRectOverride,
} from '../../../geometry/preset-text-rect-table.js';
export type { PresetTextRectDefinition } from '../../../geometry/preset-text-rect-types.js';

export { customGeometryPathsToSvgSubpaths, type CustomGeometrySubpathSvg } from './custom-geometry';

export {
	resolveCustomGeometryGuideContext,
	resolveCustomGeometryToken,
} from './custom-geometry-guides';

export { applyCustomGeometryGuideOverrides } from './custom-geometry-guide-writeback';

export {
	evaluateCustomGeometryPathData,
	evaluateCustomGeometryPaths,
} from './custom-geometry-live-eval';

export { getAdjustmentAwareClipPath } from '../../../geometry/adjustment-aware-shapes.js';

export {
	getCloudClipPath,
	getCloudCalloutClipPath,
	CLOUD_LOBE_COUNT,
	CLOUD_CALLOUT_TAIL_COUNT,
} from '../../../geometry/cloud-bezier-paths.js';

export { getConnectorAdjustment, getConnectorPathGeometry } from './connector-geometry';
export type { ConnectorPathGeometry } from './connector-geometry';

export {
	TEXT_ORIENTATION_IDENTITY,
	getElementOrientationMatrix,
	getElementTransform,
	getTextCompensationTransform,
	isTextOrientationMatrix,
	multiplyTextOrientationMatrices,
} from './transform-utils';

export {
	PRESET_SHAPE_CLIP_PATHS,
	PRESET_SHAPE_DEFINITIONS,
	PRESET_SHAPE_CATEGORY_LABELS,
	getPresetShapeClipPath,
} from '../../../geometry/preset-shape-paths.js';
export type {
	PresetShapeDefinition,
	PresetShapeCategory,
} from '../../../geometry/preset-shape-paths.js';

export {
	createBuiltinVariables,
	evaluateGuides,
	parseGuideDefinitions,
	parseAdjustmentValues,
	resolveCoordinate,
	evaluateGeometryPaths,
	ooxmlArcToSvg,
} from './guide-formula';
export type { GeometryGuide, GeometryContext } from './guide-formula';
export {
	parseStructuredCustomGeometry,
	buildCustomGeometryPathsFromNodes,
} from './custom-geometry-parser';

export { orderedXmlKey, stripXmlOrderSuffix } from './custom-geometry-command-order';

export {
	unionShapes,
	intersectShapes,
	subtractShapes,
	fragmentShapes,
	combineShapes,
	mergeShapes,
	svgPathToPolygons,
	polygonsToSvgPath,
	unionPolygons,
	intersectPolygons,
	subtractPolygons,
	unionSvgPaths,
	intersectSvgPaths,
	subtractSvgPaths,
} from '../../../geometry/shape-boolean.js';
export type { Vec2, MergeShapeOperation } from '../../../geometry/shape-boolean.js';

export { FreeformPathBuilder, douglasPeucker, catmullRomToBezier } from './freeform-builder';

export {
	isCalloutShape,
	getCalloutTier,
	getCalloutLeaderLineGeometry,
	buildCalloutLeaderLineSvgPath,
	getCalloutViewBoxBounds,
} from '../../../geometry/callout-geometry.js';
export type {
	CalloutPoint,
	CalloutLeaderLineGeometry,
} from '../../../geometry/callout-geometry.js';
