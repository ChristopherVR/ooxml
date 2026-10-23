// Chart Design galleries of Excel: Chart Styles, Change Colors and Quick Layout, drawn from the
// shared Chart Design gallery data (`../../chart/design-galleries`) and translated here. Picks run
// `chart.styles`, `chart.colors` and `chart.quick-layout`, which apply through the xlsx ChartPatch
// path (`ooxml-core/xlsx` `chartStylePresetPatch`, `chartColorPalettePatch`, `chartQuickLayoutPatch`).
import { chartColorScheme, chartStylePresetOf, chartTakesQuickLayout } from 'ooxml-core/xlsx';
import { activeChart, type EditorContext } from 'ooxml-core/xlsx/ui';
import {
	chartColorGallerySections as sharedColorSections,
	chartQuickLayoutGallerySection,
	chartStyleGallerySection,
	type ChartGallerySection,
} from '../../chart/design-galleries';

/** A shared section translated and shaped for Excel's gallery control. */
function translated(ctx: EditorContext, section: ChartGallerySection) {
	return {
		...(section.title === undefined ? {} : { title: ctx.t(section.title) }),
		columns: section.columns,
		tileWidth: section.tileWidth,
		tileHeight: section.tileHeight,
		items: section.items.map((item) => ({
			id: item.id,
			label: ctx.t(item.labelTemplate, item.labelParams),
			preview: item.previewSvg,
			applied: item.applied,
		})),
	};
}

export function chartColorGallerySections(ctx: EditorContext) {
	const workbook = ctx.workbook();
	const selectedChart = activeChart(ctx);
	if (!workbook || !selectedChart) return [];
	const scheme = chartColorScheme(workbook.theme);
	return sharedColorSections(scheme, selectedChart.chart.colorPalette ?? 10).map((section) =>
		translated(ctx, section),
	);
}

export const chartColorGalleryItems = (ctx: EditorContext) =>
	chartColorGallerySections(ctx).flatMap((section) => section.items);

export function chartStyleGallerySections(ctx: EditorContext) {
	const selectedChart = activeChart(ctx);
	if (!selectedChart) return [];
	return [translated(ctx, chartStyleGallerySection(chartStylePresetOf(selectedChart.chart)))];
}

export const chartStyleGalleryItems = (ctx: EditorContext) =>
	chartStyleGallerySections(ctx).flatMap((section) => section.items);

export function chartQuickLayoutGallerySections(ctx: EditorContext) {
	const selectedChart = activeChart(ctx);
	if (!selectedChart || !chartTakesQuickLayout(selectedChart.chart)) return [];
	return [translated(ctx, chartQuickLayoutGallerySection(selectedChart.chart.quickLayout))];
}

export const chartQuickLayoutGalleryItems = (ctx: EditorContext) =>
	chartQuickLayoutGallerySections(ctx).flatMap((section) => section.items);
