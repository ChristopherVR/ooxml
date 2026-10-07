import { CHART_COLOR_PALETTES, chartPaletteColors } from 'ooxml-core/chart';
import { chartColorScheme } from 'ooxml-core/xlsx';
import { activeChart, type EditorContext } from 'ooxml-core/xlsx/ui';
import { swatchStripSvg } from '../../ribbon/color-preview';

export function chartColorGallerySections(ctx: EditorContext) {
	const workbook = ctx.workbook();
	const selectedChart = activeChart(ctx);
	if (!workbook || !selectedChart) return [];
	const scheme = chartColorScheme(workbook.theme);
	const selected = selectedChart.chart.colorPalette ?? 10;
	return (['colorful', 'monochromatic'] as const).map((group) => ({
		title: ctx.t(group === 'colorful' ? 'Colorful' : 'Monochromatic'),
		columns: 1,
		tileWidth: 96,
		tileHeight: 14,
		items: CHART_COLOR_PALETTES.filter((palette) => palette.group === group).map((palette) => ({
			id: String(palette.id),
			label: ctx.t(group === 'colorful' ? 'Colorful Palette {n}' : 'Monochromatic Palette {n}', {
				n: palette.index,
			}),
			preview: swatchStripSvg(chartPaletteColors(palette, 6, scheme), { width: 96, height: 14 }),
			applied: selected === palette.id,
		})),
	}));
}

export const chartColorGalleryItems = (ctx: EditorContext) =>
	chartColorGallerySections(ctx).flatMap((section) => section.items);
