/**
 * The chart elements a Format pane can target, in the order PowerPoint lists
 * them in Chart Format > Current Selection and in the pane's element picker,
 * with their captions and pane titles.
 *
 * @module render/chart-format/chart-format-targets
 */
import type {
	PptxChartAxisFormatting,
	PptxChartData,
	PptxChartFormatTarget,
	PptxChartFormatTargetKey,
} from 'ooxml-core/pptx';
import { chartFormatTargetKey, parseChartFormatTargetKey } from 'ooxml-core/pptx';

/** A translatable caption: key, `{{name}}` parameters and the English fallback. */
export interface ChartFormatText {
	key: string;
	fallback: string;
	params?: Readonly<Record<string, string | number>>;
	/** Parameters that are themselves captions ("{{axis}} Major Gridlines"). */
	paramTexts?: Readonly<Record<string, ChartFormatText>>;
}

/** One entry of the element picker. */
export interface ChartFormatTargetOption {
	key: PptxChartFormatTargetKey;
	target: PptxChartFormatTarget;
	label: ChartFormatText;
}

const text = (
	key: string,
	fallback: string,
	params?: Readonly<Record<string, string | number>>,
): ChartFormatText => (params ? { key, fallback, params } : { key, fallback });

/** Whether an axis is drawn vertically (left/right edge). */
function isVerticalAxis(axis: PptxChartAxisFormatting, data: PptxChartData): boolean {
	if (axis.axPos) {
		return axis.axPos === 'l' || axis.axPos === 'r';
	}
	const horizontalBars = data.barDirection === 'bar';
	return axis.axisType === 'valAx' ? !horizontalBars : horizontalBars;
}

function isSecondaryAxis(data: PptxChartData, index: number): boolean {
	const axis = data.axes?.[index];
	if (!axis) {
		return false;
	}
	return (data.axes ?? []).slice(0, index).some((a) => a.axisType === axis.axisType);
}

/** "Vertical (Value) Axis", "Secondary Horizontal (Category) Axis", ... */
function axisCaption(data: PptxChartData, index: number): ChartFormatText {
	const axis = data.axes?.[index];
	if (!axis) {
		return text('pptx.chartFormat.target.axis', 'Axis');
	}
	if (axis.axisType === 'serAx') {
		return text('pptx.chartFormat.target.depthAxis', 'Depth (Series) Axis');
	}
	const vertical = isVerticalAxis(axis, data);
	const value = axis.axisType === 'valAx';
	const secondary = isSecondaryAxis(data, index);
	const orientation = vertical ? 'Vertical' : 'Horizontal';
	const id = secondary ? `secondary${orientation}` : orientation.toLowerCase();
	const key = `${id}${value ? 'Value' : 'Category'}`;
	const english = `${secondary ? 'Secondary ' : ''}${vertical ? 'Vertical' : 'Horizontal'} (${value ? 'Value' : 'Category'}) Axis`;
	return text(`pptx.chartFormat.target.axis.${key}`, english);
}

function seriesName(data: PptxChartData, index: number): string {
	return data.series[index]?.name || `Series ${index + 1}`;
}

/** The picker caption of `target`. */
export function chartFormatTargetLabel(
	data: PptxChartData,
	target: PptxChartFormatTarget,
): ChartFormatText {
	switch (target.kind) {
		case 'chartArea':
			return text('pptx.chartFormat.target.chartArea', 'Chart Area');
		case 'plotArea':
			return text('pptx.chartFormat.target.plotArea', 'Plot Area');
		case 'title':
			return text('pptx.chartFormat.target.title', 'Chart Title');
		case 'legend':
			return text('pptx.chartFormat.target.legend', 'Legend');
		case 'axis':
			return axisCaption(data, target.axisIndex);
		case 'majorGridlines':
		case 'minorGridlines': {
			const axis = axisCaption(data, target.axisIndex);
			const major = target.kind === 'majorGridlines';
			return {
				...text(
					major
						? 'pptx.chartFormat.target.majorGridlines'
						: 'pptx.chartFormat.target.minorGridlines',
					major ? '{{axis}} Major Gridlines' : '{{axis}} Minor Gridlines',
				),
				paramTexts: { axis },
			};
		}
		case 'series':
			return text('pptx.chartFormat.target.series', 'Series "{{name}}"', {
				name: seriesName(data, target.seriesIndex),
			});
		case 'dataLabels':
			return target.seriesIndex === undefined
				? text('pptx.chartFormat.target.dataLabels', 'Data Labels')
				: text('pptx.chartFormat.target.seriesDataLabels', 'Series "{{name}}" Data Labels', {
						name: seriesName(data, target.seriesIndex),
					});
		default:
			return text('pptx.chartFormat.target.chartArea', 'Chart Area');
	}
}

/** The pane heading for `target` ("Format Chart Area", "Format Axis", ...). */
export function chartFormatPaneTitle(target: PptxChartFormatTarget): ChartFormatText {
	switch (target.kind) {
		case 'plotArea':
			return text('pptx.chartFormat.title.plotArea', 'Format Plot Area');
		case 'title':
			return text('pptx.chartFormat.title.title', 'Format Chart Title');
		case 'legend':
			return text('pptx.chartFormat.title.legend', 'Format Legend');
		case 'axis':
			return text('pptx.chartFormat.title.axis', 'Format Axis');
		case 'majorGridlines':
			return text('pptx.chartFormat.title.majorGridlines', 'Format Major Gridlines');
		case 'minorGridlines':
			return text('pptx.chartFormat.title.minorGridlines', 'Format Minor Gridlines');
		case 'series':
			return text('pptx.chartFormat.title.series', 'Format Data Series');
		case 'dataLabels':
			return text('pptx.chartFormat.title.dataLabels', 'Format Data Labels');
		default:
			return text('pptx.chartFormat.title.chartArea', 'Format Chart Area');
	}
}

function hasTitle(data: PptxChartData): boolean {
	if (data.style?.hasTitle === false) {
		return false;
	}
	return Boolean(data.style?.hasTitle || data.title || data.titleRuns?.length);
}

/** Every element of `data` a Format pane can target, in PowerPoint's picker order. */
export function listChartFormatTargets(data: PptxChartData): ChartFormatTargetOption[] {
	const targets: PptxChartFormatTarget[] = [{ kind: 'chartArea' }];
	if (hasTitle(data)) {
		targets.push({ kind: 'title' });
	}
	(data.axes ?? []).forEach((axis, axisIndex) => {
		if (axis.deleted) {
			return;
		}
		targets.push({ kind: 'axis', axisIndex });
		if (axis.majorGridlines) {
			targets.push({ kind: 'majorGridlines', axisIndex });
		}
		if (axis.minorGridlines) {
			targets.push({ kind: 'minorGridlines', axisIndex });
		}
	});
	if (data.style?.hasLegend) {
		targets.push({ kind: 'legend' });
	}
	targets.push({ kind: 'plotArea' });
	if (data.style?.hasDataLabels) {
		targets.push({ kind: 'dataLabels' });
	}
	data.series.forEach((_, seriesIndex) => {
		targets.push({ kind: 'series', seriesIndex });
		// Each series' own labels (PowerPoint's "Series "X" Data Labels"), listed even while
		// hidden: their Label Options are where one series' labels are switched on.
		targets.push({ kind: 'dataLabels', seriesIndex });
	});
	return targets.map((target) => ({
		key: chartFormatTargetKey(target),
		target,
		label: chartFormatTargetLabel(data, target),
	}));
}

/**
 * `key` when it names an element `data` still has, else the chart area (the
 * target a stale key falls back to after the element was removed).
 */
export function resolveChartFormatTarget(
	data: PptxChartData | undefined,
	key: PptxChartFormatTargetKey | undefined,
): ChartFormatTargetOption {
	const options = data ? listChartFormatTargets(data) : [];
	const match = options.find((option) => option.key === key);
	if (match) {
		return match;
	}
	const parsed = key ? parseChartFormatTargetKey(key) : undefined;
	if (data && parsed) {
		const valid =
			(parsed.kind === 'series' && data.series[parsed.seriesIndex]) ||
			(parsed.kind === 'dataLabels' &&
				parsed.seriesIndex !== undefined &&
				data.series[parsed.seriesIndex]) ||
			((parsed.kind === 'axis' ||
				parsed.kind === 'majorGridlines' ||
				parsed.kind === 'minorGridlines') &&
				data.axes?.[parsed.axisIndex]);
		if (valid) {
			return { key: key!, target: parsed, label: chartFormatTargetLabel(data, parsed) };
		}
	}
	return {
		key: 'chartArea',
		target: { kind: 'chartArea' },
		label: text('pptx.chartFormat.target.chartArea', 'Chart Area'),
	};
}

/** Resolve a {@link ChartFormatText} with a translator that echoes unknown keys. */
export function translateChartFormatText(
	value: ChartFormatText,
	t: (key: string, params?: Readonly<Record<string, string | number>>) => string,
): string {
	const nested = Object.entries(value.paramTexts ?? {}).map(
		([name, caption]) => [name, translateChartFormatText(caption, t)] as const,
	);
	const params =
		value.params || nested.length > 0
			? { ...value.params, ...Object.fromEntries(nested) }
			: undefined;
	const translated = t(value.key, params);
	if (translated && translated !== value.key) {
		return translated;
	}
	return value.fallback.replace(/\{\{(\w+)\}\}/gu, (_, name: string) =>
		String(params?.[name] ?? ''),
	);
}
