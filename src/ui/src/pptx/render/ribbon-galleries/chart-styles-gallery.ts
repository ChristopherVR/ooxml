/**
 * Chart Design > Chart Styles: the curated recolour presets the on-canvas
 * chart quick actions already offer (`chart-quick-action-styles.ts`), as a
 * ribbon gallery. A pick is `applyChartStylePreset`, then each series' new
 * palette colour is pinned on `series.color` as well: the core save path only
 * rewrites an existing `c:ser/c:spPr` fill when `series.color` is set, so
 * without it the saved file would keep the pre-pick colours. The tiles and
 * their data are the shared Chart Design gallery (`src/ui/src/chart/design-galleries`).
 *
 * @module render/ribbon-galleries/chart-styles-gallery
 */
import type { PptxChartData, PptxElement } from 'ooxml-core/pptx';

import { chartStyleGallerySection } from '../../../chart/design-galleries';
import { applyChartStylePreset, buildChartStylePresets } from '../chart-quick-action-styles';
import type { RibbonGalleryModule } from './gallery-module';

function chartDataOf(element: PptxElement | null): PptxChartData | undefined {
	return element?.type === 'chart' ? element.chartData : undefined;
}

/** A preset applied, with its palette pinned per series so it survives a save. */
export function applyChartStyleGalleryPreset(
	chartData: PptxChartData,
	presetId: string,
): PptxChartData | null {
	const next = applyChartStylePreset(chartData, presetId);
	if (!next?.colorPalette?.length || chartData.varyColors) {
		return next;
	}
	const palette = next.colorPalette;
	return {
		...next,
		series: next.series.map((series, i) => ({ ...series, color: palette[i % palette.length] })),
	};
}

export const CHART_STYLES_GALLERY: RibbonGalleryModule = {
	build(ctx) {
		const chartData = chartDataOf(ctx.element);
		const presets = buildChartStylePresets(chartData ?? { colorPalette: undefined });
		const section = chartStyleGallerySection(presets.find((preset) => preset.applied)?.id);
		return {
			id: 'chartStyles',
			labelKey: 'pptx.gallery.chartStyles.title',
			label: 'Chart Styles',
			disabled: !chartData,
			sections: [
				{
					...section,
					items: section.items.map((item, i) => ({
						id: item.id,
						labelKey: presets[i]!.labelKey,
						label: item.label,
						previewSvg: item.previewSvg,
						applied: presets[i]!.applied,
					})),
				},
			],
		};
	},
	apply(itemId, ctx) {
		const element = ctx.element;
		const chartData = chartDataOf(element);
		const next = chartData ? applyChartStyleGalleryPreset(chartData, itemId) : null;
		if (!element || !next) {
			return null;
		}
		return {
			kind: 'element',
			elementId: element.id,
			patch: { chartData: next } as Partial<PptxElement>,
		};
	},
};
