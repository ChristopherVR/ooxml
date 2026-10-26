import {
	applyChartAppearance,
	clearChartAreaPaint,
	type ChartAppearance,
	type ChartFillAppearance,
} from '../../../chart';
import type { PptxChartData, PptxChartGradientFill, PptxChartSeries } from '../types';
import { isLineDrawnChartType } from './chart-container-type-map';
import { same } from './chart-neutral-save-common';
import type { NeutralChartEdit } from './chart-neutral-save-plan';

const fill = (color: string | undefined, gradient?: PptxChartGradientFill): ChartFillAppearance =>
	gradient
		? { kind: 'gradient', gradient }
		: color === 'none'
			? { kind: 'none' }
			: color
				? { kind: 'solid', color }
				: { kind: 'automatic' };

/** Old public flat paint fields remain compatibility inputs, with shared appearance semantics. */
export function planNeutralFlatPaint(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	const old = before.style ?? {},
		now = probe.style;
	if (now)
		for (const part of ['chartArea', 'plotArea'] as const) {
			const color = `${part}Fill` as const,
				gradient = `${part}Gradient` as const,
				border = `${part}Border` as const;
			if ([color, gradient, border].every((key) => same(old[key], now[key]))) continue;
			const was: ChartAppearance = {},
				next: ChartAppearance = {};
			if (!same(old[color], now[color]) || !same(old[gradient], now[gradient])) {
				was.fill = fill(old[color], old[gradient]);
				next.fill = fill(now[color], now[gradient]);
			}
			if (old[border] !== now[border]) {
				was.line = fill(old[border]) as ChartAppearance['line'];
				next.line = fill(now[border]) as ChartAppearance['line'];
			}
			const automatic =
				now[color] === undefined && now[gradient] === undefined && now[border] === undefined;
			edits.push((space) =>
				automatic
					? clearChartAreaPaint(space, part)
					: applyChartAppearance(space, { kind: part }, was, next),
			);
			for (const key of [color, gradient, border]) {
				if (old[key] === undefined) delete now[key];
				else Object.assign(now, { [key]: old[key] });
			}
		}
	if (probe.series.length !== before.series.length) return;
	probe.series.forEach((series, index) => {
		const previous = before.series[index]!;
		const lineDrawn = isLineDrawnChartType(series.seriesChartType ?? probe.chartType);
		if (!same(previous.gradientFill, series.gradientFill) && series.gradientFill && !lineDrawn) {
			const was = { fill: fill(previous.color, previous.gradientFill) };
			const next = { fill: fill(series.color, series.gradientFill) };
			edits.push((space) =>
				applyChartAppearance(space, { kind: 'series', seriesIndex: index }, was, next),
			);
			if (previous.gradientFill === undefined) delete series.gradientFill;
			else series.gradientFill = previous.gradientFill;
			if (previous.color === undefined) delete series.color;
			else series.color = previous.color;
		}
		const keys = ['lineNoFill', 'lineWidth', 'lineDashStyle', 'lineGradientFill'] as const;
		if (keys.every((key) => same(previous[key], series[key]))) return;
		const line = (value: PptxChartSeries): NonNullable<ChartAppearance['line']> => ({
			...(fill(
				value.lineNoFill ? 'none' : lineDrawn ? value.color : undefined,
				value.lineGradientFill,
			) as NonNullable<ChartAppearance['line']>),
			...(value.lineWidth === undefined ? {} : { width: value.lineWidth }),
			...(value.lineDashStyle === undefined ? {} : { dash: value.lineDashStyle }),
		});
		const was = { line: line(previous) },
			next = { line: line(series) };
		edits.push((space) =>
			applyChartAppearance(space, { kind: 'series', seriesIndex: index }, was, next),
		);
		for (const key of keys) {
			if (previous[key] === undefined) delete series[key];
			else Object.assign(series, { [key]: previous[key] });
		}
	});
}
