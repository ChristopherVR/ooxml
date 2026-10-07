import type { ChartViewModel } from './chart-view';
import { chartAreaRect, chartTextAttributes } from './chart-svg-appearance';
import { text } from './chart-svg-util';
import { chartTextWidth, chartFontMetrics, type ChartSvgOptions } from './chart-svg-text-metrics';
import type { DiagramTextSpacing } from '../../diagram/types';
import { chartPointsToPixels } from './chart-appearance';

function spacing(
	value: DiagramTextSpacing | undefined,
	naturalHeight: number,
	fallback: number,
): number {
	return value
		? value.unit === 'points'
			? chartPointsToPixels(value.value)
			: naturalHeight * value.value
		: fallback;
}

/** Paint explicit rich lines with host font metrics when available. Automatic wrapping remains open. */
export function chartRichTitleSvg(
	model: ChartViewModel,
	width: number,
	options: ChartSvgOptions = {},
): { markup: string; height: number } | undefined {
	if (!model.titleText) return undefined;
	const lines: (typeof model.titleText.paragraphs)[number][] = [];
	for (const paragraph of model.titleText.paragraphs) {
		let line: (typeof lines)[number] = {
			...paragraph,
			runs: [],
		};
		lines.push(line);
		for (const run of paragraph.runs) {
			const pieces = run.text.split('\n');
			for (const [index, piece] of pieces.entries()) {
				if (index) {
					line = { ...paragraph, runs: [] };
					lines.push(line);
				}
				if (piece) line.runs.push({ ...run, text: piece });
			}
		}
	}
	const base = chartTextAttributes(model, 'title', 14);
	const metrics = lines.map((line) => {
		const runs = line.runs.map((run) => ({
			...run,
			attrs: chartTextAttributes(
				{ ...model, appearance: { title: run.appearance } },
				'title',
				base.size,
			),
		}));
		return {
			...line,
			runs,
			width: runs.reduce((sum, run) => sum + chartTextWidth(run.text, run.attrs, options), 0),
			size: runs.length ? Math.max(...runs.map((run) => run.attrs.size ?? 14)) : (base.size ?? 14),
			fontMetrics: (runs.length ? runs.map((run) => run.attrs) : [base]).map((attrs) =>
				chartFontMetrics(attrs, options),
			),
		};
	});
	// Stack font boxes around a common baseline, including runs smaller than the line's largest font.
	const measured = metrics.every((line) => line.fontMetrics.every((font) => font !== undefined));
	const vertical = metrics.map((line) => {
		const ascent = measured ? Math.max(...line.fontMetrics.map((font) => font!.ascent)) : line.size;
		const descent = measured
			? Math.max(...line.fontMetrics.map((font) => font!.descent))
			: line.size * 0.4 + 4;
		const natural = ascent + descent;
		const height = spacing(line.lineSpacing, natural, natural);
		const leading = (height - natural) / 2;
		return {
			ascent: ascent + leading,
			descent: descent + leading,
			height,
			before: spacing(line.spaceBefore, natural, 0),
			after: spacing(line.spaceAfter, natural, 0),
			leading,
		};
	});
	const boxWidth = Math.min(width - 16, Math.max(0, ...metrics.map((line) => line.width)) + 8);
	const height = vertical.reduce(
		(sum, line) => sum + line.height + line.before + line.after,
		measured ? 4 : 0,
	);
	let baseline =
		8 +
		(metrics[0]?.size ?? base.size ?? 14) +
		(vertical[0]?.before ?? 0) +
		(vertical[0]?.leading ?? 0);
	const spans: string[] = [];
	for (const [index, line] of metrics.entries()) {
		if (index)
			baseline +=
				vertical[index - 1]!.descent +
				vertical[index - 1]!.after +
				vertical[index]!.before +
				vertical[index]!.ascent;
		let x =
			line.align === 'l'
				? (width - boxWidth) / 2 + 4
				: line.align === 'r'
					? (width + boxWidth) / 2 - 4 - line.width
					: (width - line.width) / 2;
		for (const run of line.runs) {
			spans.push(
				text(x, baseline, run.text, {
					...run.attrs,
					anchor: 'start',
				})
					.replace('<text ', '<tspan data-chart-title-run="true" ')
					.replace('</text>', '</tspan>'),
			);
			x += chartTextWidth(run.text, run.attrs, options);
		}
	}
	const parent = text(0, 0, '', { ...base, underline: false }).replace(
		'</text>',
		`${spans.join('')}</text>`,
	);
	return {
		markup:
			chartAreaRect(model, 'title', {
				x: (width - boxWidth) / 2,
				y: 4,
				w: Math.max(0, boxWidth),
				h: height,
			}) + parent,
		height: Math.max(24, height),
	};
}
