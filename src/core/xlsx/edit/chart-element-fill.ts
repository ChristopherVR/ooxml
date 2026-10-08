import type { DiagramFill, DiagramColor } from '../../diagram/types';
import { resolveDrawingColor } from '../../drawingml/drawing-color';
import type { ChartObject, Color } from '../model';
import type { ChartPatch } from './charts';
import { chartDrawingColor } from './chart-series-fill';
import { chartGradientFillEdit, type ChartGradientEdit } from './chart-series-gradient';

export const CHART_FILL_PARTS = ['chartArea', 'plotArea', 'title', 'legend'] as const;
export type ChartFillPart = (typeof CHART_FILL_PARTS)[number];
export type ChartElementFills = Partial<Record<ChartFillPart, DiagramFill>>;
const white = (): DiagramColor => ({ kind: 'scheme', value: 'lt1', transforms: [] });

/** Effective background fill, including imported chart-style defaults. */
export function chartElementFill(chart: ChartObject, part: ChartFillPart): DiagramFill {
	return (
		chart.formatting?.entries[part]?.fill ??
		chart.styleDefinition?.entries[part]?.fill ??
		(part === 'chartArea' ? { kind: 'solid', color: white() } : { kind: 'none' })
	);
}

/** Materialize fill patches for editing and transient views without exposing unsupported metadata edits. */
export function chartWithElementFills(chart: ChartObject, fills: ChartElementFills): ChartObject {
	const formatting = structuredClone(chart.formatting ?? { sourceXml: '', entries: {} });
	for (const [part, fill] of Object.entries(fills)) {
		if (!CHART_FILL_PARTS.includes(part as ChartFillPart))
			throw new RangeError('Unsupported chart fill target');
		if (!fill || !['none', 'solid', 'gradient'].includes(fill.kind))
			throw new RangeError('Unsupported chart fill');
		formatting.entries[part] = {
			...formatting.entries[part],
			sourceXml: formatting.entries[part]?.sourceXml ?? '',
			fill: structuredClone(fill),
		};
	}
	return { ...chart, formatting };
}

function patch(chart: ChartObject, part: ChartFillPart, fill: DiagramFill): ChartPatch | undefined {
	if (JSON.stringify(chartElementFill(chart, part)) === JSON.stringify(fill)) return undefined;
	return { elementFills: { [part]: fill } };
}

export function chartElementFillPatch(
	chart: ChartObject,
	part: ChartFillPart,
	color: Color | null,
): ChartPatch | undefined {
	if (color === null) return patch(chart, part, { kind: 'none' });
	const choice = chartDrawingColor(color);
	if (!choice) throw new RangeError('Invalid chart fill color');
	const previous = chartElementFill(chart, part);
	if (previous.kind === 'solid')
		choice.transforms.push(
			...structuredClone(
				previous.color.transforms.filter((item) =>
					['alpha', 'alphaMod', 'alphaOff'].includes(item.name),
				),
			),
		);
	return patch(chart, part, { kind: 'solid', color: choice });
}

export function chartElementSolidFillPatch(
	chart: ChartObject,
	part: ChartFillPart,
): ChartPatch | undefined {
	const fill = chartElementFill(chart, part);
	if (fill.kind === 'solid') return undefined;
	return patch(chart, part, {
		kind: 'solid',
		color: fill.kind === 'gradient' ? (fill.stops[0]?.color ?? white()) : white(),
	});
}

export function chartElementGradientPatch(
	chart: ChartObject,
	part: ChartFillPart,
	edit: ChartGradientEdit,
): { patch: ChartPatch; stopIndex: number } | undefined {
	const old = chartElementFill(chart, part);
	const result = chartGradientFillEdit(old, edit, old.kind === 'solid' ? old.color : white());
	if (!result) return undefined;
	return { patch: { elementFills: { [part]: result.fill } }, stopIndex: result.stopIndex };
}

export function chartElementTransparency(
	chart: ChartObject,
	part: ChartFillPart,
): number | undefined {
	const fill = chartElementFill(chart, part);
	if (fill.kind !== 'solid') return undefined;
	const resolved = resolveDrawingColor(
		fill.color,
		{ scheme: () => '#000000' },
		{ transformOrder: 'document' },
	);
	return resolved ? Math.round((1 - resolved.alpha) * 100000) / 1000 : undefined;
}

export function chartElementTransparencyPatch(
	chart: ChartObject,
	part: ChartFillPart,
	percent: number,
): ChartPatch | undefined {
	if (!Number.isFinite(percent) || percent < 0 || percent > 100)
		throw new RangeError('Chart transparency must be from 0 to 100');
	const fill = chartElementFill(chart, part);
	if (fill.kind !== 'solid') return undefined;
	const color = structuredClone(fill.color);
	color.transforms = color.transforms.filter(
		(item) => !['alpha', 'alphaMod', 'alphaOff'].includes(item.name),
	);
	if (percent !== 0)
		color.transforms.push({ name: 'alpha', value: String(Math.round((100 - percent) * 1000)) });
	return patch(chart, part, { kind: 'solid', color });
}
