import type { ChartViewModel } from './chart-view';
import { chartAreaRect, chartTextAttributes } from './chart-svg-appearance';
import { text } from './chart-svg-util';
import { chartTextWidth, chartFontMetrics, type ChartSvgOptions } from './chart-svg-text-metrics';

/** Paint explicit rich lines with host font metrics when available. Automatic wrapping remains open. */
export function chartRichTitleSvg(
	model: ChartViewModel,
	width: number,
	options: ChartSvgOptions = {},
): { markup: string; height: number } | undefined {
	if (!model.titleText) return undefined;
	type Run = (typeof model.titleText.paragraphs)[number]['runs'][number];
	const lines: { align?: string; runs: Run[] }[] = [];
	for (const paragraph of model.titleText.paragraphs) {
		let line: (typeof lines)[number] = {
			...(paragraph.align ? { align: paragraph.align } : {}),
			runs: [],
		};
		lines.push(line);
		for (const run of paragraph.runs) {
			const pieces = run.text.split('\n');
			for (const [index, piece] of pieces.entries()) {
				if (index) {
					line = { ...(paragraph.align ? { align: paragraph.align } : {}), runs: [] };
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
	const vertical = metrics.map((line) => ({
		ascent: Math.max(...line.fontMetrics.map((font) => font?.ascent ?? 0)),
		descent: Math.max(...line.fontMetrics.map((font) => font?.descent ?? 0)),
	}));
	const boxWidth = Math.min(width - 16, Math.max(0, ...metrics.map((line) => line.width)) + 8);
	const height = measured
		? vertical.reduce((sum, line) => sum + line.ascent + line.descent, 4)
		: metrics.reduce((sum, line) => sum + line.size * 1.4 + 4, 0);
	let y = 4;
	let baseline = 8 + (metrics[0]?.size ?? base.size ?? 14);
	const spans: string[] = [];
	for (const [index, line] of metrics.entries()) {
		if (measured && index) baseline += vertical[index - 1]!.descent + vertical[index]!.ascent;
		let x =
			line.align === 'l'
				? (width - boxWidth) / 2 + 4
				: line.align === 'r'
					? (width + boxWidth) / 2 - 4 - line.width
					: (width - line.width) / 2;
		for (const run of line.runs) {
			spans.push(
				text(x, measured ? baseline : y + line.size + 4, run.text, {
					...run.attrs,
					anchor: 'start',
				})
					.replace('<text ', '<tspan data-chart-title-run="true" ')
					.replace('</text>', '</tspan>'),
			);
			x += chartTextWidth(run.text, run.attrs, options);
		}
		y += line.size * 1.4 + 4;
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
