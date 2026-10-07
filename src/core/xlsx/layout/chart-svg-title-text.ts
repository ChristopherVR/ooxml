import type { ChartViewModel } from './chart-view';
import { chartAreaRect, chartTextAttributes } from './chart-svg-appearance';
import { text } from './chart-svg-util';
import { chartTextWidth, type ChartSvgOptions } from './chart-svg-text-metrics';

/** Paint explicit rich lines with optional host widths; line height and wrapping remain approximate. */
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
		};
	});
	const boxWidth = Math.min(width - 16, Math.max(0, ...metrics.map((line) => line.width)) + 8);
	const height = metrics.reduce((sum, line) => sum + line.size * 1.4 + 4, 0);
	let y = 4;
	const spans: string[] = [];
	for (const line of metrics) {
		let x =
			line.align === 'l'
				? (width - boxWidth) / 2 + 4
				: line.align === 'r'
					? (width + boxWidth) / 2 - 4 - line.width
					: (width - line.width) / 2;
		for (const run of line.runs) {
			spans.push(
				text(x, y + line.size + 4, run.text, { ...run.attrs, anchor: 'start' })
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
