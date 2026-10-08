import type { ChartViewModel } from './chart-view';
import { chartAreaRect, chartTextAttributes } from './chart-svg-appearance';
import { text, n } from './chart-svg-util';
import { chartTextWidth, chartFontMetrics, type ChartSvgOptions } from './chart-svg-text-metrics';
import type { DrawingTextSpacing } from '../../drawingml/types';
import { chartPointsToPixels } from './chart-appearance';
import type { ChartTitleText } from './chart-title-text';
import { wrapStyledRuns } from '../../text/wrap-styled-runs';
import { hasManualLayoutFields, resolveManualLayoutRect } from '../../chart/manual-layout';

function spacing(
	value: DrawingTextSpacing | undefined,
	naturalHeight: number,
	fallback: number,
): number {
	return value
		? value.unit === 'points'
			? chartPointsToPixels(value.value)
			: naturalHeight * value.value
		: fallback;
}

/** Paint title runs with shared horizontal flow and optional host font metrics. */
export function chartRichTitleSvg(
	model: ChartViewModel,
	width: number,
	options: ChartSvgOptions = {},
	frameHeight = 400,
): { markup: string; height: number } | undefined {
	if (!model.title) return undefined;
	const base = chartTextAttributes(model, 'title', 14);
	// Native automatic title boxes use 80% of chart width, with 8 CSS pixels of inner padding.
	// Manual text boxes and platform-specific font layout still need separate comparisons.
	const capacity = Math.max(1, width * 0.8 - 8);
	if (
		!model.titleText &&
		!hasManualLayoutFields(model.titleLayout) &&
		!model.titleOverlay &&
		!model.title.includes('\n') &&
		chartTextWidth(model.title, base, options) <= capacity
	)
		return undefined;
	const paragraphs = model.titleText?.paragraphs ?? [
		{
			runs: [{ text: model.title, appearance: model.appearance?.title ?? {} }],
		},
	];
	const hardLines: ChartTitleText['paragraphs'] = [];
	for (const paragraph of paragraphs) {
		let line: ChartTitleText['paragraphs'][number] = {
			...paragraph,
			runs: [],
		};
		hardLines.push(line);
		for (const run of paragraph.runs) {
			const pieces = run.text.split('\n');
			for (const [index, piece] of pieces.entries()) {
				if (index) {
					line = { ...paragraph, runs: [] };
					hardLines.push(line);
				}
				if (piece) line.runs.push({ ...run, text: piece });
			}
		}
	}
	const lines: ChartTitleText['paragraphs'] = hardLines.flatMap((line) => {
		const wrapped = wrapStyledRuns(
			line.runs,
			capacity,
			(content, run) =>
				chartTextWidth(
					content,
					chartTextAttributes(
						{ ...model, appearance: { title: run.appearance } },
						'title',
						base.size,
					),
					options,
				),
			{ breakWords: true },
		);
		const { spaceBefore, spaceAfter, ...properties } = line;
		return wrapped.map((runs, index) => ({
			...properties,
			runs,
			...(index === 0 && spaceBefore ? { spaceBefore } : {}),
			...(index === wrapped.length - 1 && spaceAfter ? { spaceAfter } : {}),
		}));
	});
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
	const markup =
		chartAreaRect(model, 'title', {
			x: (width - boxWidth) / 2,
			y: 4,
			w: Math.max(0, boxWidth),
			h: height,
		}) + parent;
	const automatic = { x: (width - boxWidth) / 2, y: 4, width: boxWidth, height };
	const manual = resolveManualLayoutRect(
		model.titleLayout,
		{ width, height: frameHeight },
		automatic,
	);
	return {
		markup: manual
			? `<g data-chart-title-layout="true" transform="translate(${n(manual.x - automatic.x)} ${n(manual.y - automatic.y)})">${markup}</g>`
			: markup,
		height: Math.max(24, height),
	};
}
