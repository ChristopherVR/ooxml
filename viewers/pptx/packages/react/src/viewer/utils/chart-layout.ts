/**
 * Thin re-export shim → `pptx-viewer-shared`.
 *
 * Plot-area layout (`computeLayout`, `PlotLayout`) plus the secondary-axis and
 * data-table reservation helpers were consolidated into `pptx-viewer-shared`
 * (`render/chart-helpers.ts` and `render/chart-axis.ts`). This shim preserves
 * the historical React import surface so the chart `.tsx` renderers and
 * colocated tests keep importing the same names unchanged.
 */
export type { PlotLayout } from 'ooxml-ui/pptx';
export type { LayoutOptions } from 'ooxml-ui/pptx';
export {
	computeLayout,
	hasSecondaryValueAxis,
	hasSecondaryCategoryAxis,
	getSecondaryValueAxis,
	getSecondaryCategoryAxis,
	computeLayoutOptions,
	getSecondaryValueAxisId,
	getPrimaryValueAxisId,
	isSeriesOnSecondaryAxis,
	splitSeriesByAxis,
	computeDataTableHeight,
} from 'ooxml-ui/pptx';
