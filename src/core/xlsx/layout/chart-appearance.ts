import {
	CHART_STYLE_PARTS,
	resolveChartStyleDefinition,
	type ChartStyleEntry,
	type ChartStylePart,
	type ResolvedChartStyleEntry,
} from '../../chart/style-definition';
import { resolveDrawingColor } from '../../diagram/drawing-color';
import { EMU_PER_PIXEL, EMU_PER_POINT } from '../../units/constants';
import type { ChartObject, ThemePalette } from '../model';
import { chartColorScheme } from './chart-colors';
import type { ChartGradientFill } from '../../chart/gradient-definition';

export interface ChartAppearanceEntry extends ResolvedChartStyleEntry {
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
	if (!chart.styleDefinition && !chart.formatting) return undefined;
	const entries: Record<string, ChartStyleEntry> = {};
	for (const part of CHART_STYLE_PARTS) {
		const base = chart.styleDefinition?.entries[part];
		const direct = chart.formatting?.entries[part];
		if (!base && !direct) continue;
		const entry: ChartStyleEntry = {
			...base,
			...direct,
			sourceXml: direct?.sourceXml ?? base?.sourceXml ?? '',
		};
		if (base?.line || direct?.line) entry.line = { ...base?.line, ...direct?.line };
		entries[part] = entry;
	}
	const scheme = chartColorScheme(theme) as unknown as Readonly<Record<string, string>>;
	const resolve = (color: Parameters<typeof resolveDrawingColor>[0]) =>
		resolveDrawingColor(color, { scheme: (name) => scheme[name] }, { hslRounding: 'halfDown' });
	const out: ChartAppearance =
		resolveChartStyleDefinition({ entries, sourceXml: '' }, (color) => {
			const value = resolve(color);
			if (!value || value.alpha === 1) return value?.hex;
			const channels = [1, 3, 5].map((offset) =>
				Number.parseInt(value.hex.slice(offset, offset + 2), 16),
			);
			return `rgba(${channels.join(',')},${value.alpha})`;
		}) ?? {};
	for (const part of CHART_STYLE_PARTS) {
		const entry = entries[part];
		if (!entry) continue;
		const family = typeface(entry, theme);
		if (family) (out[part] ??= {}).typeface = family;
		if (entry.axisVisible !== undefined) (out[part] ??= {}).axisVisible = entry.axisVisible;
		if (entry.labelsVisible !== undefined) (out[part] ??= {}).labelsVisible = entry.labelsVisible;
		if (entry.line?.widthEmu === 0) (out[part] ??= {}).lineWidth = 0;
		if (entry.fill?.kind === 'gradient') {
			const fill = entry.fill;
			const stops = fill.stops.flatMap((stop) => {
				const color = resolve(stop.color);
				return color ? [{ position: stop.position, color: color.hex, opacity: color.alpha }] : [];
			});
			const focus = fill.fillToRect;
			(out[part] ??= {}).gradient = {
				type: fill.path ? 'radial' : 'linear',
				stops,
				...(fill.angle === undefined ? {} : { angle: fill.angle }),
				...(focus
					? { focalPoint: { x: (1 + focus.l - focus.r) / 2, y: (1 + focus.t - focus.b) / 2 } }
					: {}),
			};
		}
	}
	return out;
}
