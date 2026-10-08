// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// What a Word chart shows: the shared chart painter's SVG, drawn from the values cached in the
// chart part with the document theme, or a labelled frame naming a family it does not draw.
import { chartSpaceKind, renderChartSpaceSvg } from '../chart/index';
import type { ThemeColorScheme } from '../drawingml/theme-model';
import type { DocxChart } from './chart';
import type { ThemeCatalog } from './theme-model';

/** The painted chart (`svg`), or only the frame label when the chart is not drawn. */
export interface DocxChartPaint {
	svg?: string;
	/** `Chart`, or `Chart (bubble)` for a family the painter does not draw. */
	label: string;
}

/** The document theme's colour slots as a DrawingML colour scheme. */
export function themeCatalogColorScheme(theme: ThemeCatalog | undefined): ThemeColorScheme {
	const colors: ThemeColorScheme['colors'] = {};
	for (const [slot, value] of Object.entries(theme?.colors ?? {}) as [
		keyof ThemeColorScheme['colors'],
		string | undefined,
	][])
		if (value) colors[slot] = { kind: 'srgb', value: value.replace(/^#/, ''), transforms: [] };
	return { colors };
}

/** Paints a chart at the drawing extent (CSS pixels). Never throws: failures give the frame. */
export function docxChartPaint(
	chart: DocxChart | undefined,
	widthPx: number,
	heightPx: number,
	theme?: ThemeCatalog,
): DocxChartPaint {
	const space = chart?.chartSpace;
	if (!space) return { label: 'Chart' };
	const kind = chartSpaceKind(space);
	if (!kind.drawn) return { label: kind.type ? `Chart (${kind.type})` : 'Chart' };
	try {
		const fonts = {
			...(theme?.fonts.major.latin ? { major: theme.fonts.major.latin } : {}),
			...(theme?.fonts.minor.latin ? { minor: theme.fonts.minor.latin } : {}),
		};
		const svg = renderChartSpaceSvg(space, {
			width: Math.max(1, Math.round(widthPx)),
			height: Math.max(1, Math.round(heightPx)),
			theme: themeCatalogColorScheme(theme),
			fonts,
			...(chart.formatting ? { formatting: chart.formatting } : {}),
			...(chart.colorPalette !== undefined ? { colorPalette: chart.colorPalette } : {}),
		});
		return { svg, label: chart.title ? `Chart: ${chart.title}` : 'Chart' };
	} catch {
		return { label: kind.type ? `Chart (${kind.type})` : 'Chart' };
	}
}
