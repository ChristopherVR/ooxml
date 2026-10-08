/**
 * chart-palette.ts  -  colour-ramp helpers for chart styling.
 *
 * `tint` (lighten toward white) and `shade` (darken toward black) are the two
 * primitive colour transforms behind the Office chart style-id palettes. They,
 * the style-id → palette mapping (`getChartStylePalette`), the palette
 * precedence (`resolveChartPalette`) and the default fallback palette
 * (`DEFAULT_CHART_PALETTE`) live in `chart-style-palette.ts` and are
 * re-exported here so a binding can pull the whole palette surface from one
 * module.
 *
 * Extracted from the React `viewer/utils/chart-style-palettes.ts`.
 *
 * @module chart-palette
 */

export {
	DEFAULT_CHART_PALETTE,
	getChartStylePalette,
	resolveChartPalette,
	shade,
	tint,
} from './chart-style-palette';
