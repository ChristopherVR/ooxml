import type { DrawingColor, DrawingFill, DrawingLine, DrawingTextBody } from '../drawingml/types';
import type { ChartManualLayout } from './manual-layout';

/** Chart elements consumed by the current Office chart painters. */
export const CHART_STYLE_PARTS = [
	'title',
	'axisTitle',
	'categoryAxis',
	'valueAxis',
	'legend',
	'dataLabel',
	'dataPoint',
	'dataPointLine',
	'gridlineMajor',
	'gridlineMinor',
	'chartArea',
	'plotArea',
] as const;
export type ChartStylePart = (typeof CHART_STYLE_PARTS)[number];

export interface ChartStyleReference {
	index: string;
	color?: DrawingColor;
}

/** Theme-relative formatting; opaque XML retains properties not yet interpreted. */
export interface ChartStyleEntry {
	/** Imported placement; sourceXml preserves it until layout authoring is implemented. */
	layout?: ChartManualLayout;
	overlay?: boolean;
	textBody?: DrawingTextBody;
	/** Element text defaults before paragraph/run overrides. */
	textBodyStyle?: ChartTextProperties;
	/** Direct axis formatting, when imported from a chart part. */
	axisVisible?: boolean;
	labelsVisible?: boolean;
	fontSize?: number;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	typeface?: string;
	textColor?: DrawingColor;
	fontRef?: ChartStyleReference;
	lineRef?: ChartStyleReference;
	fillRef?: ChartStyleReference;
	effectRef?: ChartStyleReference;
	fill?: DrawingFill;
	line?: DrawingLine;
	/** Direct text-effect XML; an empty effect list suppresses inherited effects. */
	textEffectsXml?: string;
	sourceXml: string;
}

export type ChartTextProperties = Pick<
	ChartStyleEntry,
	| 'fontSize'
	| 'bold'
	| 'italic'
	| 'underline'
	| 'typeface'
	| 'textColor'
	| 'fontRef'
	| 'textEffectsXml'
>;

/** Office chart-style part, distinct from a chart color-style palette. */
export interface ChartStyleDefinition {
	/** Chart-space text body defaults, separate from chart-area paint. */
	textDefaults?: ChartStyleEntry;
	/** Built-in chart-space style, distinct from an external style part ID. */
	builtInStyle?: number;
	id?: number;
	entries: Record<string, ChartStyleEntry>;
	sourceXml: string;
}

/** Resolved subset consumed by existing chart painters, with sizes and widths in points. */
export interface ResolvedChartStyleEntry {
	fontSize?: number;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	color?: string;
	lineColor?: string;
	lineWidth?: number;
	fillColor?: string;
}
export type ResolvedChartStyleDefinition = Partial<Record<ChartStylePart, ResolvedChartStyleEntry>>;

/** One shared precedence rule for explicit formatting and style references. */
export function resolveChartStyleDefinition(
	style: ChartStyleDefinition,
	resolve: (color: DrawingColor) => string | undefined,
): ResolvedChartStyleDefinition | undefined {
	const out: ResolvedChartStyleDefinition = {};
	const fillColor = (fill: DrawingFill | undefined) =>
		fill?.kind === 'none' ? 'none' : fill?.kind === 'solid' ? resolve(fill.color) : undefined;
	for (const name of CHART_STYLE_PARTS) {
		const source = style.entries[name];
		if (!source) continue;
		const entry: ResolvedChartStyleEntry = {};
		for (const key of ['fontSize', 'bold', 'italic', 'underline'] as const) {
			const value = source[key];
			if (value !== undefined) Object.assign(entry, { [key]: value });
		}
		const line =
			fillColor(source.line?.fill) ?? (source.lineRef?.color && resolve(source.lineRef.color));
		const fill = fillColor(source.fill) ?? (source.fillRef?.color && resolve(source.fillRef.color));
		const color =
			(source.textColor && resolve(source.textColor)) ??
			(source.fontRef?.color && resolve(source.fontRef.color));
		if (color !== undefined) entry.color = color;
		if (line !== undefined) entry.lineColor = line;
		if (fill !== undefined) entry.fillColor = fill;
		if (source.line?.widthEmu !== undefined && source.line.widthEmu > 0)
			entry.lineWidth = source.line.widthEmu / 12700;
		if (Object.keys(entry).length) out[name] = entry;
	}
	return Object.keys(out).length ? out : undefined;
}
