import {
	CHART_STYLE_PARTS,
	resolveChartStyleDefinition,
	type ChartStyleEntry,
	type ChartStylePart,
	type ResolvedChartStyleEntry,
} from '../../chart/style-definition';
import { resolveDrawingColor } from '../../diagram/drawing-color';
import { drawingColorCss } from '../../diagram/drawing-color-css';
import { EMU_PER_PIXEL, EMU_PER_POINT } from '../../units/constants';
import type { ChartObject, ThemePalette } from '../model';
import { chartColorScheme } from './chart-colors';
import { resolveChartGradient, type ChartGradientFill } from '../../chart/gradient-definition';
import { resolveDrawingShadowXml, type DrawingSvgShadow } from '../../diagram/drawing-shadow';
import { builtInChartTextStyle, effectiveBuiltInChartStyle } from '../../chart/built-in-text-style';
import { inheritChartText } from '../../chart/text-inheritance';

export interface ChartAppearanceEntry extends ResolvedChartStyleEntry {
	textShadow?: DrawingSvgShadow;
	textShadowFilter?: string;
	gradient?: ChartGradientFill;
	typeface?: string;
	axisVisible?: boolean;
	labelsVisible?: boolean;
}
export type ChartAppearance = Partial<Record<ChartStylePart, ChartAppearanceEntry>>;
export const chartPointsToPixels = (points: number): number =>
	(points * EMU_PER_POINT) / EMU_PER_PIXEL;

function typeface(entry: ChartStyleEntry, theme: ThemePalette): string | undefined {
	const name = entry.typeface;
	if (name?.startsWith('+mj-')) return theme.majorFont;
	if (name?.startsWith('+mn-')) return theme.minorFont;
	return (
		name ??
		(entry.fontRef?.index === 'major'
			? theme.majorFont
			: entry.fontRef?.index === 'minor'
				? theme.minorFont
				: undefined)
	);
}

/** Direct properties override native style defaults; theme fonts and colors remain live. */
export function chartAppearance(
	chart: ChartObject,
	theme: ThemePalette,
): ChartAppearance | undefined {
	const defaults = builtInChartTextStyle(effectiveBuiltInChartStyle(chart.formatting));
	const entries: Record<string, ChartStyleEntry> = {};
	for (const part of CHART_STYLE_PARTS) {
		const base = chart.styleDefinition?.entries[part];
		const direct = chart.formatting?.entries[part];
		if (!defaults[part] && !base && !direct) continue;
		const entry: ChartStyleEntry = {
			...inheritChartText(
				part,
				{ ...defaults[part], ...base, sourceXml: base?.sourceXml ?? '' },
				defaults[part],
				chart.formatting?.textDefaults,
			),
			...direct,
			sourceXml: direct?.sourceXml ?? base?.sourceXml ?? '',
		};
		if (base?.line || direct?.line) entry.line = { ...base?.line, ...direct?.line };
		entries[part] = entry;
	}
	const scheme = chartColorScheme(theme) as unknown as Readonly<Record<string, string>>;
	const resolve = (color: Parameters<typeof resolveDrawingColor>[0]) =>
		resolveDrawingColor(
			color,
			{ scheme: (name) => scheme[name] },
			{ hslRounding: 'halfDown', transformOrder: 'document' },
		);
	const out: ChartAppearance =
		resolveChartStyleDefinition({ entries, sourceXml: '' }, (color) =>
			drawingColorCss(resolve(color)),
		) ?? {};
	for (const part of CHART_STYLE_PARTS) {
		const entry = entries[part];
		if (!entry) continue;
		const family = typeface(entry, theme);
		const shadow = resolveDrawingShadowXml(entry.textEffectsXml, {
			scheme: (name) => scheme[name],
		});
		if (shadow) (out[part] ??= {}).textShadow = shadow;
		if (family) (out[part] ??= {}).typeface = family;
		if (entry.axisVisible !== undefined) (out[part] ??= {}).axisVisible = entry.axisVisible;
		if (entry.labelsVisible !== undefined) (out[part] ??= {}).labelsVisible = entry.labelsVisible;
		if (entry.line?.widthEmu === 0) (out[part] ??= {}).lineWidth = 0;
		if (entry.fill?.kind === 'gradient') {
			(out[part] ??= {}).gradient = resolveChartGradient(entry.fill, resolve);
		}
	}
	return out;
}
