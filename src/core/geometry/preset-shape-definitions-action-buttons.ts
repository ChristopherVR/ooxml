/**
 * ECMA-376 ST_ShapeType preset geometry definitions: action buttons.
 *
 * The 12 `actionButton*` shapes are a flat rectangle plus the glyph (home,
 * help, sound, ...) drawn as shaded sub-paths, exactly as PowerPoint draws
 * them; the geometry itself lives in `preset-shape-definitions-exact-action-buttons`.
 * COM-verified: `Shape.Adjustments.Count` reads 0 for every `actionButton*`, so
 * there is no `avLst` (an authored guide would make PowerPoint refuse the file).
 */

import type { PresetShapeGeometryDefinition } from './preset-shape-definitions-table';
import { EXACT_ACTION_BUTTON_PRESET_DEFINITIONS } from './preset-shape-definitions-exact-action-buttons';

/**
 * Action-button preset definitions exported for aggregation into
 * `PRESET_SHAPE_GEOMETRY_TABLE`. Keys are ECMA-376 ST_ShapeType names.
 */
export const ACTION_BUTTON_PRESET_DEFINITIONS: Record<string, PresetShapeGeometryDefinition> =
	EXACT_ACTION_BUTTON_PRESET_DEFINITIONS;
