// Waterfall bars: each step floats from the running total before it to the total after it, coloured
// by role (increase, decrease, total), with dashed connectors between consecutive bars.
import type { Frame } from './chart-svg-cartesian';
import { line, n, rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';

const CONNECTOR_COLOR = '#7F7F7F';
/** Bar width as a share of its category band (Office's default 50% gap). */
const BAR_SHARE = 2 / 3;

export function waterfallBars(model: ChartViewModel, frame: Frame, out: string[]): void {
	const view = model.waterfall;
	if (!view) return;
	const width = frame.bandSize * BAR_SHARE;
	view.steps.forEach((step, i) => {
		const a = frame.value(step.startValue);
		const b = frame.value(step.endValue);
		const color = step.isSubtotal
			? view.colors.total
			: step.value >= 0
				? view.colors.increase
				: view.colors.decrease;
		const x = frame.band(i) - width / 2;
		out.push(`<g data-chart-series="0" data-chart-point="${step.sourceIndex}">`);
		out.push(rect(x, Math.min(a, b), width, Math.max(Math.abs(b - a), 1), color));
		out.push('</g>');
		const next = view.steps[i + 1];
		if (view.connectors && next) {
			const xNext = frame.band(i + 1) - width / 2;
			out.push(
				line(x + width, b, xNext, b, CONNECTOR_COLOR, 0.75).replace(
					'/>',
					` stroke-dasharray="${n(3)} ${n(2)}"/>`,
				),
			);
		}
	});
}
