/**
 * The Chart Design galleries every product shares: Chart Styles (the curated presets of
 * `ooxml-core/chart` `style-presets`), Change Colors (Colorful Palettes 1-4 and Monochromatic
 * Palettes 1-13, `color-palettes`) and Quick Layout (Layout 1-11, `quick-layouts`). Each builder
 * returns DOM-free section data with English labels and SVG tiles; a product maps it onto its own
 * ribbon gallery (PowerPoint's `RibbonGalleryModule`, Excel's gallery control) and translates the
 * label templates, then applies a pick through its chart edit path (`chart/edit`
 * `applyChartStylePreset`, `applyChartColorPalette`, `applyChartQuickLayout`, or the product's own
 * model mapping of them).
 *
 * @module chart/design-galleries
 */
import {
	CHART_COLOR_PALETTES,
	CHART_QUICK_LAYOUTS,
	CHART_STYLE_PRESETS,
	chartPaletteColors,
	type ChartColorPalette,
	type ChartColorScheme,
} from 'ooxml-core/chart';

import { chartLayoutTileSvg, columnChartTileSvg, swatchStripSvg } from './gallery-tiles';

/** One gallery tile. `label` is English; `labelTemplate` + `labelParams` translate it. */
export interface ChartGalleryItem {
	id: string;
	label: string;
	labelTemplate: string;
	labelParams?: Record<string, number>;
	previewSvg: string;
	applied: boolean;
}

export interface ChartGallerySection {
	id: string;
	/** English section title, when the gallery has sections. */
	title?: string;
	columns: number;
	tileWidth: number;
	tileHeight: number;
	items: ChartGalleryItem[];
}

export const CHART_STYLE_TILE = { width: 60, height: 44 } as const;
export const CHART_COLOR_TILE = { width: 96, height: 14 } as const;
export const CHART_LAYOUT_TILE = { width: 56, height: 44 } as const;

/** Chart Styles: one section of the curated presets; `appliedId` marks the current one. */
export function chartStyleGallerySection(appliedId?: string): ChartGallerySection {
	return {
		id: 'styles',
		columns: 6,
		tileWidth: CHART_STYLE_TILE.width,
		tileHeight: CHART_STYLE_TILE.height,
		items: CHART_STYLE_PRESETS.map((preset) => ({
			id: preset.id,
			label: preset.label,
			labelTemplate: preset.label,
			previewSvg: columnChartTileSvg(preset.colors(), CHART_STYLE_TILE),
			applied: preset.id === appliedId,
		})),
	};
}

const COLOR_GROUPS = [
	['colorful', 'Colorful', 'Colorful Palette {n}'],
	['monochromatic', 'Monochromatic', 'Monochromatic Palette {n}'],
] as const;

function colorItem(
	palette: ChartColorPalette,
	scheme: ChartColorScheme,
	template: string,
	appliedId: number | undefined,
): ChartGalleryItem {
	const count = palette.group === 'colorful' ? Math.max(palette.base.length, 6) : 6;
	return {
		id: String(palette.id),
		label: template.replace('{n}', String(palette.index)),
		labelTemplate: template,
		labelParams: { n: palette.index },
		previewSvg: swatchStripSvg(chartPaletteColors(palette, count, scheme), CHART_COLOR_TILE),
		applied: palette.id === appliedId,
	};
}

/** Change Colors: the Colorful and Monochromatic sections resolved under `scheme`. */
export function chartColorGallerySections(
	scheme: ChartColorScheme,
	appliedId?: number,
): ChartGallerySection[] {
	return COLOR_GROUPS.map(([group, title, template]) => ({
		id: group,
		title,
		columns: 1,
		tileWidth: CHART_COLOR_TILE.width,
		tileHeight: CHART_COLOR_TILE.height,
		items: CHART_COLOR_PALETTES.filter((palette) => palette.group === group).map((palette) =>
			colorItem(palette, scheme, template, appliedId),
		),
	}));
}

/** Quick Layout: Layout 1-11; `appliedN` marks the layout the chart shows. */
export function chartQuickLayoutGallerySection(appliedN?: number): ChartGallerySection {
	return {
		id: 'layouts',
		columns: 4,
		tileWidth: CHART_LAYOUT_TILE.width,
		tileHeight: CHART_LAYOUT_TILE.height,
		items: CHART_QUICK_LAYOUTS.map((layout) => ({
			id: `layout${layout.n}`,
			label: `Layout ${layout.n}`,
			labelTemplate: 'Layout {n}',
			labelParams: { n: layout.n },
			previewSvg: chartLayoutTileSvg(layout, CHART_LAYOUT_TILE),
			applied: layout.n === appliedN,
		})),
	};
}

/** The layout number of a Quick Layout item id (`layout5` is 5), or `undefined`. */
export function chartQuickLayoutItemNumber(id: string): number | undefined {
	const match = /^layout(\d+)$/u.exec(id);
	const n = match ? Number(match[1]) : NaN;
	return CHART_QUICK_LAYOUTS.some((layout) => layout.n === n) ? n : undefined;
}
